"""Offline, conservative PDF -> Markdown with a fail-closed evidence pipeline.

No models, network requests, spelling correction, or silent OCR. Run --help.
The original PDF is the authority; text accounting is NOT semantic validation.
"""
from __future__ import annotations

import argparse
from collections import Counter
from dataclasses import asdict, dataclass, field
import gzip
import hashlib
import html
import json
import logging
import math
from pathlib import Path
import re
import shutil
import statistics
import sys
from typing import Any, Iterable, TextIO, TypedDict
import unicodedata

import pdfplumber
from pypdf import PdfReader

LOG = logging.getLogger("pdf_ingest")
BBox = tuple[float, float, float, float]


class Word(TypedDict):
    """A source token; id is unique within its page."""
    id: int
    text: str
    x0: float
    x1: float
    top: float
    bottom: float


@dataclass(frozen=True)
class Config:
    """Conservative defaults. Coordinates are in PDF points."""
    margin_fraction: float = 0.16
    repetition_fraction: float = 0.60
    line_tolerance: float = 3.0
    render_dpi: int = 110
    table_settings: dict[str, Any] = field(default_factory=dict)


@dataclass
class Block:
    """Auditable unit, including source tokens and final Markdown byte range."""
    kind: str
    words: list[Word]
    markdown: str
    bbox: BBox
    context: dict[str, str] = field(default_factory=dict)
    metadata: dict[str, Any] = field(default_factory=dict)
    start: int = 0
    end: int = 0


def sha256(path: Path) -> str:
    """Hash an artifact without loading it into RAM."""
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for part in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(part)
    return digest.hexdigest()


def dump_line(stream: TextIO, value: Any) -> None:
    """Append one UTF-8 JSON record (no non-finite coordinates)."""
    stream.write(json.dumps(value, ensure_ascii=False, allow_nan=False) + "\n")


def canonical(text: str) -> str:
    """Normalize whitespace only; never rewrite spelling or punctuation."""
    return " ".join(text.split())


def character_counts(text: str) -> Counter[str]:
    """Independent-parser comparison only; NFKC is NOT applied to output."""
    return Counter(c for c in unicodedata.normalize("NFKC", text) if not c.isspace())


def escape_text(text: str) -> str:
    """Prevent source content from becoming Markdown/HTML syntax."""
    text = html.escape(canonical(text), quote=False)
    text = re.sub(r"([\\`*_\[\]#|])", r"\\\1", text)
    text = re.sub(r"^(\s*)([-+])(?=\s)", r"\1\\\2", text)
    return re.sub(r"^(\s*\d+)\.(?=\s)", r"\1\\.", text)


def unescape_text(text: str) -> str:
    """Inverse of our emitter, used for strict serialized-content checks."""
    return html.unescape(re.sub(r"\\([\\`*_{}\[\]()#+.!|~>-])", r"\1", text))


def bounds(words: list[Word]) -> BBox:
    return (min(w["x0"] for w in words), min(w["top"] for w in words),
            max(w["x1"] for w in words), max(w["bottom"] for w in words))


def inside(word: Word, bbox: BBox) -> bool:
    x = (word["x0"] + word["x1"]) / 2
    y = (word["top"] + word["bottom"]) / 2
    return bbox[0] <= x < bbox[2] and bbox[1] <= y < bbox[3]


def lines_of(words: Iterable[Word], tolerance: float) -> list[list[Word]]:
    """Group physical lines by baseline; preserve left-to-right token order."""
    lines: list[list[Word]] = []
    for word in sorted(words, key=lambda w: (w["bottom"], w["x0"], w["id"])):
        overlap = 0.0
        if lines:
            top = statistics.median(w["top"] for w in lines[-1])
            bottom = statistics.median(w["bottom"] for w in lines[-1])
            overlap = max(0.0, min(bottom, word["bottom"]) - max(top, word["top"]))
            minimum_height = max(0.01, min(bottom-top, word["bottom"]-word["top"]))
        if lines and (abs(word["bottom"] - bottom) <= tolerance or
                      overlap / minimum_height >= 0.5):
            lines[-1].append(word)
        else:
            lines.append([word])
    for line in lines:
        line.sort(key=lambda w: (w["x0"], w["id"]))
    return lines


def text_of(words: Iterable[Word]) -> str:
    return " ".join(w["text"] for w in words)


def get_words(page: Any) -> list[Word]:
    """Keep ligatures and duplicate glyphs; audit instead of deduplicating."""
    extracted = page.extract_words(x_tolerance=2, y_tolerance=3,
                                   keep_blank_chars=False, expand_ligatures=False)
    return [Word(id=i, text=w["text"], x0=float(w["x0"]), x1=float(w["x1"]),
                 top=float(w["top"]), bottom=float(w["bottom"]))
            for i, w in enumerate(extracted)]


def margin_key(line: list[Word], height: float, config: Config) -> str | None:
    """Learn repeated lines only in narrow page-margin bands."""
    bbox = bounds(line)
    band = "top" if bbox[3] < height * config.margin_fraction else (
        "bottom" if bbox[1] > height * (1 - config.margin_fraction) else "")
    return band + ":" + canonical(text_of(line)) if band else None


def discover_noise(pdf: Any, config: Config) -> set[str]:
    """First pass, storing only repeated-margin signatures, never page text."""
    counts: Counter[str] = Counter()
    for page in pdf.pages:
        keys = {key for line in lines_of(get_words(page), config.line_tolerance)
                if (key := margin_key(line, page.height, config))}
        counts.update(keys)
        page.close()
    threshold = max(3, math.ceil(len(pdf.pages) * config.repetition_fraction))
    return {key for key, count in counts.items() if count >= threshold}


def separate_noise(words: list[Word], height: float, repeated: set[str],
                   config: Config, width: float | None = None) -> tuple[list[Word], list[dict[str, Any]]]:
    """Exclude margin noise from Markdown but retain every excluded token."""
    kept: list[Word] = []
    removed: list[dict[str, Any]] = []
    physical_lines = lines_of(words, config.line_tolerance)
    for line in physical_lines:
        key = margin_key(line, height, config)
        page_number = (bounds(line)[1] > height * (1 - config.margin_fraction) and
                       line is physical_lines[-1] and
                       (width is None or width * 0.35 < (bounds(line)[0]+bounds(line)[2])/2 < width * 0.65) and
                       re.fullmatch(r"\d{1,5}", text_of(line)) is not None)
        if key in repeated or page_number:
            removed.append({"reason": "bottom_page_number" if page_number else
                            "repeated_margin", "words": line})
        else:
            kept.extend(line)
    return kept, removed


def extract_tables(page: Any, words: list[Word], config: Config,
                   issues: list[str]) -> tuple[list[Block], list[Word]]:
    """Assign tokens to geometric cells exactly once; fall back on ambiguity.

    Generic headers prevent promoting the first data row into a false header.
    Merged-cell topology is retained in metadata and requires visual review.
    """
    blocks: list[Block] = []
    consumed: set[int] = set()
    # Filled text/background rectangles are not borders. Thin filled rectangles
    # ARE borders in PDFs exported by Word; lines_strict would miss them entirely.
    geometry = page.filter(lambda obj: obj.get("object_type") != "rect" or
                           min(obj["width"], obj["height"]) <= 2 or obj.get("stroke"))
    tables = sorted(geometry.find_tables(config.table_settings),
                    key=lambda t: (-(t.bbox[2]-t.bbox[0])*(t.bbox[3]-t.bbox[1])))
    for table in tables:
        if len(table.rows) < 2 or len(table.columns) < 2:
            continue
        local = [w for w in words if inside(w, table.bbox)]
        if not local:
            continue
        if any(w["id"] in consumed for w in local):
            issues.append("overlapping_table_candidate")
            continue
        cells = [(r, c, cell) for r, row in enumerate(table.rows)
                 for c, cell in enumerate(row.cells) if cell is not None]
        assigned: dict[tuple[int, int], list[Word]] = {(r, c): [] for r, c, _ in cells}
        valid = True
        for word in local:
            owners = [(r, c, cell) for r, c, cell in cells if inside(word, cell)]
            if len(owners) != 1:
                valid = False
                break
            r, c, cell = owners[0]
            assigned[r, c].append(word)
            if word["x0"] < cell[0]-2 or word["x1"] > cell[2]+2:
                issues.append("word_crosses_cell_boundary")
        if not valid:
            issues.append("ambiguous_table_ownership_fallback_to_text")
            continue
        ordered: list[Word] = []
        grid: list[list[str]] = []
        topology: list[dict[str, Any]] = []
        ncols = max(len(row.cells) for row in table.rows)
        for r, row in enumerate(table.rows):
            output_row: list[str] = []
            for c in range(ncols):
                cell_words = [w for line in lines_of(assigned.get((r, c), []),
                                                    config.line_tolerance) for w in line]
                ordered.extend(cell_words)
                output_row.append(escape_text(text_of(cell_words)))
                cell = row.cells[c] if c < len(row.cells) else None
                topology.append({"row": r, "column": c, "bbox": cell,
                                 "word_ids": [w["id"] for w in cell_words]})
                if cell is None:
                    issues.append("merged_or_missing_table_cell")
            grid.append(output_row)
        markdown = "| " + " | ".join(f"Column {c+1}" for c in range(ncols)) + " |\n"
        markdown += "| " + " | ".join("---" for _ in range(ncols)) + " |\n"
        markdown += "\n".join("| " + " | ".join(row) + " |" for row in grid)
        blocks.append(Block("table", ordered, markdown, tuple(table.bbox),
                            metadata={"cells": topology, "generated_header": True,
                                      "rows": len(grid), "columns": ncols}))
        consumed.update(w["id"] for w in local)
        issues.append("table_layout_requires_review")
    return blocks, [w for w in words if w["id"] not in consumed]


def heading_level(text: str, context: dict[str, str]) -> int | None:
    """Recognize anchored legal markers; never match inline article references."""
    if re.fullmatch(r"(?:CHAPTER|TITLE|PART|ANNEX)\s+[IVXLCDM\d]+(?:\s*[:–—-].*)?",
                    text, re.IGNORECASE):
        return 1
    if re.fullmatch(r"Article\s+\d+(?:\.\d+)*[a-z]?", text, re.IGNORECASE):
        return 2
    if re.match(r"^\d+\.\d+(?:\.\d+)*\s+\S", text):
        return min(3, text.split()[0].count(".") + 1)
    if re.match(r"^(?:\(\d+\)|\d+\.)\s+", text):
        return 3
    marker = re.match(r"^\(([a-z]+)\)\s+", text)
    if marker:
        value = marker.group(1)
        if re.fullmatch(r"[ivxlcdm]+", value) and (len(value) > 1 or "4" in context):
            return 5
        return 4
    return None


def text_blocks(words: list[Word], config: Config, issues: list[str]) -> list[Block]:
    """Join soft line wraps; preserve gaps, clauses, and suspicious flow.

    No dehyphenation: a line-final hyphen may be lexical or discretionary.
    """
    result: list[Block] = []
    current: list[Word] = []
    previous: list[Word] | None = None
    for line in lines_of(words, config.line_tolerance):
        value = text_of(line)
        gap = bounds(line)[1] - bounds(previous)[3] if previous else 0
        size = statistics.median(w["bottom"]-w["top"] for w in line)
        large_horizontal_gap = any(b["x0"]-a["x1"] > max(28, size * 3)
                                   for a, b in zip(line, line[1:]))
        if large_horizontal_gap:
            issues.append("possible_columns_or_unruled_table")
        special = heading_level(value, {}) is not None
        previous_special = (previous is not None and heading_level(text_of(previous), {}) is not None
                            and not re.match(r"^(?:\(\d+\)|\d+\.|\([a-z]+\))\s+", text_of(previous)))
        if current and (gap > size * 0.65 or special or previous_special or large_horizontal_gap):
            result.append(Block("text", current, "", bounds(current)))
            current = []
        if previous and text_of(previous).endswith("-"):
            issues.append("line_final_hyphen_preserved")
        current.extend(line)
        previous = line
    if current:
        result.append(Block("text", current, "", bounds(current)))
    return result


def assemble_blocks(blocks: list[Block], context: dict[str, str], issues: list[str]) -> None:
    """Attach persistent legal context and emit Markdown in geometric order."""
    for block in blocks:
        if block.kind == "text":
            value = text_of(block.words)
            level = heading_level(value, context)
            if level:
                clause = re.match(r"^(\(\d+\)|\d+\.|\([a-z]+\))\s+(.*)$", value)
                title = clause.group(1) if clause else value
                if clause and re.match(r"^\([ivxlcdm]\)$", title):
                    issues.append("ambiguous_single_roman_or_letter_marker")
                for key in list(context):
                    if int(key) >= level:
                        del context[key]
                context[str(level)] = title
                block.kind = "heading"
                block.metadata["level"] = level
                block.markdown = "#" * level + " " + escape_text(title)
                if clause:
                    block.markdown += "\n\n" + escape_text(clause.group(2))
            else:
                block.markdown = escape_text(value)
        block.context = context.copy()


def emitted_text(block: dict[str, Any]) -> str:
    """Recover ONLY source content from our strictly controlled Markdown."""
    md = block["markdown"]
    if block["kind"] == "table":
        values: list[str] = []
        for row in md.splitlines()[2:]:
            # Our escaping always escapes source pipes; unescaped pipes delimit cells.
            cells = re.split(r"(?<!\\)\|", row)[1:-1]
            values.extend(unescape_text(c.strip()) for c in cells)
        return canonical(" ".join(values))
    if block["kind"] == "heading":
        md = re.sub(r"^#{1,6} ", "", md)
    return canonical(unescape_text(md))


def validate_record(record: dict[str, Any]) -> list[str]:
    """Check token coverage, uniqueness, identity, and serialization order."""
    failures: list[str] = []
    source = {w["id"]: w for w in record["words"]}
    accounted = [w for b in record["blocks"] for w in b["words"]]
    accounted += [w for item in record["excluded"] for w in item["words"]]
    if Counter(w["id"] for w in accounted) != Counter(source.keys()):
        failures.append("source_word_coverage_or_duplication_failure")
    if any(source.get(w["id"]) != w for w in accounted):
        failures.append("source_word_identity_failure")
    for block in record["blocks"]:
        if emitted_text(block) != canonical(text_of(block["words"])):
            failures.append("markdown_serialization_content_or_order_failure")
    return sorted(set(failures))


def convert(source: Path, output: Path, config: Config) -> int:
    """Run two extraction passes, streaming evidence and Markdown per page.

    An existing output directory is never overwritten. A failed run leaves
    RUNNING.json and partial evidence for diagnosis; no release file is created.
    """
    source = source.resolve(strict=True)
    output.mkdir(parents=True, exist_ok=False)
    (output / "RUNNING.json").write_text(json.dumps({"source": str(source)}), encoding="utf-8")
    shutil.copyfile(source, output / "source.pdf")
    source_hash = sha256(output / "source.pdf")
    page_reports: list[dict[str, Any]] = []
    context: dict[str, str] = {}
    (output / "review").mkdir()
    # PdfReader caches document objects; content evidence itself is streamed.
    with (output / "source.pdf").open("rb") as independent_stream, \
            pdfplumber.open(output / "source.pdf") as pdf, \
            (output / "candidate.md").open("wb") as markdown, \
            gzip.open(output / "evidence.jsonl.gz", "wt", encoding="utf-8") as evidence:
        independent = PdfReader(independent_stream, strict=False)
        if independent.is_encrypted:
            raise ValueError("Encrypted PDF: provide a locally decrypted copy.")
        LOG.info("Learning margin patterns across %d pages", len(pdf.pages))
        noise = discover_noise(pdf, config)
        for number, page in enumerate(pdf.pages, 1):
            issues: list[str] = []
            words = get_words(page)
            chars = [{key: char.get(key) for key in
                      ("text", "x0", "x1", "top", "bottom", "fontname", "size", "upright")}
                     for char in page.chars]
            raw = "".join(str(c["text"] or "") for c in chars)
            if not words:
                issues.append("no_text_layer_requires_local_ocr")
            if "\ufffd" in raw or re.search(r"\(cid:\d+\)", raw):
                issues.append("unmapped_or_replacement_glyph")
            if any(unicodedata.category(c) in {"Co", "Cs"} or
                   (unicodedata.category(c) == "Cc" and not c.isspace()) for c in raw):
                issues.append("private_use_or_control_glyph_requires_review")
            if any(not c["upright"] for c in chars):
                issues.append("rotated_text_requires_review")
            if character_counts(raw) != character_counts(text_of(words)):
                issues.append("glyph_to_word_accounting_mismatch")
            try:
                other = independent.pages[number-1].extract_text() or ""
                delta_a = character_counts(raw) - character_counts(other)
                delta_b = character_counts(other) - character_counts(raw)
                if delta_a or delta_b:
                    issues.append("independent_parser_character_mismatch")
                independent_check = {"pdfplumber_only": dict(delta_a), "pypdf_only": dict(delta_b)}
            except Exception as error:
                other = ""
                independent_check = {"error": str(error)}
                issues.append("independent_parser_failed")
            # Body images/vectors may carry relationships absent from the text layer.
            body_images = [obj for obj in page.images if
                           page.height * config.margin_fraction <
                           (obj.get("top", 0) + obj.get("bottom", 0)) / 2 <
                           page.height * (1-config.margin_fraction)]
            if body_images or page.curves:
                issues.append("graphics_or_diagram_semantics_require_review")
            kept, excluded = separate_noise(words, page.height, noise, config, page.width)
            try:
                tables, remaining = extract_tables(page, kept, config, issues)
            except Exception as error:
                LOG.warning("Table extraction failed on page %d: %s", number, error)
                issues.append("table_extraction_failed_fallback_to_text")
                tables, remaining = [], kept
            # Partition text at table boundaries so paragraphs cannot jump over a table.
            blocks: list[Block] = list(tables)
            cuts = sorted({b.bbox[1] for b in tables} | {b.bbox[3] for b in tables})
            partitions: dict[int, list[Word]] = {}
            for word in remaining:
                region = sum(word["top"] >= cut for cut in cuts)
                partitions.setdefault(region, []).append(word)
            for group in partitions.values():
                blocks.extend(text_blocks(group, config, issues))
            blocks.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
            assemble_blocks(blocks, context, issues)
            if number > 1 and blocks and blocks[0].kind == "text":
                issues.append("page_boundary_continuation_requires_review")
            for block in blocks:
                block.start = markdown.tell()
                markdown.write((block.markdown + "\n\n").encode("utf-8"))
                block.end = markdown.tell()
            record = {"page": number, "width": page.width, "height": page.height,
                      "chars": chars, "words": words, "excluded": excluded,
                      "blocks": [asdict(b) for b in blocks], "independent_text": other,
                      "independent_check": independent_check}
            failures = validate_record(record)
            record["failures"] = failures
            record["issues"] = sorted(set(issues))
            # All source pages, including apparently clean ones, get a visual reference.
            try:
                page.to_image(resolution=config.render_dpi).save(output / "review" / f"page-{number:04}.png")
            except Exception as error:
                record["issues"].append("page_render_failed")
                record["render_error"] = str(error)
            dump_line(evidence, record)
            page_reports.append({"page": number, "words": len(words),
                                 "excluded_words": sum(len(e["words"]) for e in excluded),
                                 "tables": len(tables), "issues": record["issues"], "failures": failures})
            page.close()
            if number % 10 == 0:
                LOG.info("Processed %d/%d pages", number, len(pdf.pages))
    report = {"schema_version": 1, "status": "blocked" if any(p["failures"] for p in page_reports)
              else "review_required", "source_sha256": source_hash,
              "candidate_sha256": sha256(output / "candidate.md"),
              "evidence_sha256": sha256(output / "evidence.jsonl.gz"),
              "config": asdict(config), "versions": {"python": sys.version,
              "pdfplumber": pdfplumber.__version__, "pypdf": __import__("pypdf").__version__},
              "pages": page_reports, "issue_counts": dict(Counter(i for p in page_reports for i in p["issues"])),
              "note": "Coverage proves emitted text accounting, not visual reading order or semantic fidelity."}
    (output / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    review = {"source_sha256": source_hash, "candidate_sha256": report["candidate_sha256"],
              "reviewer": "", "pages": [{"page": p["page"], "approved": False, "note": ""} for p in page_reports]}
    (output / "review-template.json").write_text(json.dumps(review, indent=2), encoding="utf-8")
    (output / "RUNNING.json").unlink()
    LOG.info("Finished: %s; %s", report["status"], output)
    return 2 if report["status"] == "blocked" else 3


def verify(output: Path) -> dict[str, Any]:
    """Recheck final files, hashes, offsets and per-block source token order."""
    if (output / "RUNNING.json").exists():
        raise ValueError("Incomplete conversion (RUNNING.json exists).")
    report = json.loads((output / "report.json").read_text(encoding="utf-8"))
    for name, key in (("source.pdf", "source_sha256"), ("candidate.md", "candidate_sha256"),
                      ("evidence.jsonl.gz", "evidence_sha256")):
        if sha256(output / name) != report[key]:
            raise ValueError(f"Artifact changed: {name}")
    pages = 0
    end = 0
    with gzip.open(output / "evidence.jsonl.gz", "rt", encoding="utf-8") as evidence, \
            (output / "candidate.md").open("rb") as md:
        for line in evidence:
            record = json.loads(line)
            pages += 1
            if record["page"] != pages or validate_record(record):
                raise ValueError(f"Evidence validation failed on page {pages}")
            for block in record["blocks"]:
                if block["start"] != end:
                    raise ValueError("Non-contiguous Markdown offsets")
                expected = (block["markdown"] + "\n\n").encode("utf-8")
                if block["end"] - block["start"] != len(expected) or md.read(len(expected)) != expected:
                    raise ValueError(f"Markdown differs from source accounting on page {pages}")
                end = block["end"]
        if md.read(1) or pages != len(report["pages"]):
            raise ValueError("Unexpected trailing content or page count")
    return report


def release(output: Path, review_path: Path) -> int:
    """Publish ready.md only with a complete, hash-bound local human review."""
    report = verify(output)
    if report["status"] == "blocked":
        raise ValueError("Accounting failures cannot be overridden by review.")
    review = json.loads(review_path.read_text(encoding="utf-8"))
    if any(review.get(key) != report[key] for key in ("source_sha256", "candidate_sha256")):
        raise ValueError("Review refers to different source/output hashes.")
    if not isinstance(review.get("reviewer"), str) or not review["reviewer"].strip():
        raise ValueError("A named reviewer is required.")
    decisions = review.get("pages", [])
    if [p.get("page") for p in decisions] != [p["page"] for p in report["pages"]]:
        raise ValueError("Every page must be reviewed exactly once, in order.")
    if any(p.get("approved") is not True or not isinstance(p.get("note"), str)
           or not p["note"].strip() for p in decisions):
        raise ValueError("Every page needs explicit approval and a review note.")
    target = output / "ready.md"
    if target.exists():
        raise FileExistsError(target)
    with (output / "accepted-review.json").open("x", encoding="utf-8") as accepted:
        json.dump(review, accepted, ensure_ascii=False, indent=2)
    pending = output / "ready.md.pending"
    with pending.open("xb") as dest, (output / "candidate.md").open("rb") as src:
        shutil.copyfileobj(src, dest)
    pending.rename(target)
    LOG.info("Human-reviewed output released: %s", target)
    return 0


def main(argv: list[str] | None = None) -> int:
    """CLI exit codes: 0 verified/released, 1 error, 2 blocked, 3 needs review."""
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    convert_parser = sub.add_parser("convert", help="Create a new evidence-backed candidate")
    convert_parser.add_argument("source", type=Path)
    convert_parser.add_argument("output", type=Path)
    convert_parser.add_argument("--table-settings", type=Path, help="Local pdfplumber table settings JSON")
    convert_parser.add_argument("--dpi", type=int, default=110)
    convert_parser.add_argument("--margin-fraction", type=float, default=0.16)
    verify_parser = sub.add_parser("verify", help="Verify evidence, not semantic correctness")
    verify_parser.add_argument("output", type=Path)
    release_parser = sub.add_parser("release", help="Accept a completed local visual review")
    release_parser.add_argument("output", type=Path)
    release_parser.add_argument("--review", type=Path, required=True)
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    try:
        if args.command == "convert":
            if not 72 <= args.dpi <= 300:
                raise ValueError("DPI must be between 72 and 300.")
            if not 0.03 <= args.margin_fraction <= 0.20:
                raise ValueError("Margin fraction must be between 0.03 and 0.20.")
            settings = json.loads(args.table_settings.read_text(encoding="utf-8")) if args.table_settings else {}
            if not isinstance(settings, dict):
                raise ValueError("Table settings must be a JSON object.")
            return convert(args.source, args.output, Config(render_dpi=args.dpi, table_settings=settings,
                                                            margin_fraction=args.margin_fraction))
        if args.command == "verify":
            report = verify(args.output)
            LOG.info("Accounting and hashes verified; semantic status: %s", report["status"])
            return 0
        return release(args.output, args.review)
    except (Exception, KeyboardInterrupt) as error:
        LOG.error("%s: %s", type(error).__name__, error)
        return 1


if __name__ == "__main__":
    sys.exit(main())
