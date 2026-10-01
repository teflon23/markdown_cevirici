"""Regression checks for word loss, table corruption, noise, and release gating."""
from __future__ import annotations

from dataclasses import asdict
import gzip
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest

from pdf_to_markdown import (Block, Config, Word, assemble_blocks, emitted_text,
                             escape_text, extract_tables, heading_level, lines_of, release,
                             separate_noise, sha256, validate_record, verify)


def word(i: int, text: str, x: float = 10, y: float = 20) -> Word:
    return Word(id=i, text=text, x0=x, x1=x+10, top=y, bottom=y+10)


class PipelineTests(unittest.TestCase):
    def test_source_syntax_cannot_become_markup(self) -> None:
        values = [r"a|b", r"a\|b", "<script>", "&amp;", "**bold**", "[x](url)", "�", "ﬁ", "#"]
        words = [word(i, text) for i, text in enumerate(values)]
        block = Block("text", words, escape_text(" ".join(values)), (0, 0, 100, 20))
        self.assertEqual(emitted_text(asdict(block)), " ".join(values))

    def test_leading_numbered_text_roundtrip(self) -> None:
        for value in ["1. List of obligations", "- condition", "+ amendment"]:
            block = Block("text", [word(0, value)], escape_text(value), (0, 0, 100, 20))
            self.assertEqual(emitted_text(asdict(block)), value)

    def test_loss_and_duplicate_detection(self) -> None:
        w = word(0, "obligation")
        record = {"words": [w], "blocks": [], "excluded": []}
        self.assertTrue(validate_record(record))
        block = asdict(Block("text", [w], "obligation", (0, 0, 20, 20)))
        record["blocks"] = [block]
        self.assertFalse(validate_record(record))
        record["blocks"] = [block, block]
        self.assertTrue(validate_record(record))

    def test_serialized_word_reordering_is_detected(self) -> None:
        words = [word(0, "shall"), word(1, "not")]
        block = asdict(Block("text", words, "not shall", (0, 0, 20, 20)))
        self.assertIn("markdown_serialization_content_or_order_failure",
                      validate_record({"words": words, "blocks": [block], "excluded": []}))

    def test_word_identity_mutation_is_detected(self) -> None:
        block = asdict(Block("text", [word(0, "may")], "may", (0, 0, 20, 20)))
        self.assertIn("source_word_identity_failure", validate_record(
            {"words": [word(0, "must")], "blocks": [block], "excluded": []}))

    def test_body_page_like_numbers_are_kept(self) -> None:
        words = [word(0, "25", y=300), word(1, "99", y=960)]
        kept, removed = separate_noise(words, 1000, set(), Config())
        self.assertEqual([w["text"] for w in kept], ["25"])
        self.assertEqual(removed[0]["words"][0]["text"], "99")

    def test_footnote_marker_is_not_a_page_number(self) -> None:
        words = [word(0, "3", x=70, y=723), word(1, "51", x=293, y=760)]
        kept, removed = separate_noise(words, 842, set(), Config(), width=595)
        self.assertEqual([w["text"] for w in kept], ["3"])
        self.assertEqual([w["text"] for e in removed for w in e["words"]], ["51"])

    def test_superscript_joins_its_physical_line(self) -> None:
        marker = Word(id=0, text="3", x0=70, x1=74, top=722.7, bottom=729.66)
        body = Word(id=1, text="Footnote", x0=77, x1=110, top=724.824, bottom=733.824)
        self.assertEqual(lines_of([marker, body], 3), [[marker, body]])

    def test_repeated_body_text_is_not_margin_noise(self) -> None:
        words = [word(0, "JOINT", y=300), word(1, "COMMITTEE", x=30, y=300)]
        kept, removed = separate_noise(words, 1000, {"top:JOINT COMMITTEE"}, Config())
        self.assertEqual(len(kept), 2)
        self.assertEqual(removed, [])

    def test_legal_hierarchy(self) -> None:
        context: dict[str, str] = {}
        values = ["CHAPTER II", "Article 25.2", "(1) obligation", "(g) condition", "(iii) detail"]
        blocks = [Block("text", [word(i, v)], "", (0, i*20, 100, i*20+10)) for i, v in enumerate(values)]
        assemble_blocks(blocks, context, [])
        self.assertEqual([b.metadata["level"] for b in blocks], [1, 2, 3, 4, 5])
        self.assertEqual(blocks[-1].context["2"], "Article 25.2")
        self.assertEqual(blocks[-1].context["4"], "(g)")
        self.assertEqual(emitted_text(asdict(blocks[-1])), "(iii) detail")
        self.assertIsNone(heading_level("Article 2 requires compliance.", {}))

    @staticmethod
    def table_page(merged: bool = False) -> SimpleNamespace:
        rows = [SimpleNamespace(cells=[(0, 0, 50, 40), (50, 0, 100, 40)]),
                SimpleNamespace(cells=[(0, 40, 100 if merged else 50, 80),
                                       None if merged else (50, 40, 100, 80)])]
        table = SimpleNamespace(rows=rows, columns=[0, 1], bbox=(0, 0, 100, 80))
        page = SimpleNamespace(find_tables=lambda settings: [table])
        page.filter = lambda predicate: page
        return page

    def test_multiline_cell_and_pipe_roundtrip(self) -> None:
        words = [word(0, "a|b"), word(1, "second", y=31), word(2, "right", x=60), word(3, "last", y=50)]
        issues: list[str] = []
        blocks, remainder = extract_tables(self.table_page(), words, Config(), issues)
        self.assertFalse(remainder)
        self.assertEqual(len(blocks), 1)
        self.assertEqual(len(blocks[0].markdown.splitlines()), 4)
        self.assertEqual(emitted_text(asdict(blocks[0])), "a|b second right last")
        self.assertFalse(validate_record({"words": words, "blocks": [asdict(blocks[0])], "excluded": []}))

    def test_merged_cell_never_duplicates_words(self) -> None:
        words = [word(0, "left"), word(1, "right", x=60), word(2, "merged", x=60, y=50)]
        issues: list[str] = []
        blocks, _ = extract_tables(self.table_page(True), words, Config(), issues)
        self.assertIn("merged_or_missing_table_cell", issues)
        self.assertEqual(blocks[0].markdown.count("merged"), 1)
        self.assertFalse(validate_record({"words": words, "blocks": [asdict(blocks[0])], "excluded": []}))

    def test_unowned_table_word_falls_back(self) -> None:
        page = self.table_page()
        table = page.find_tables({})[0]
        table.rows[0].cells[1] = None
        issues: list[str] = []
        words = [word(0, "preserve", x=60)]
        blocks, remaining = extract_tables(page, words, Config(), issues)
        self.assertEqual(blocks, [])
        self.assertEqual(remaining, words)
        self.assertIn("ambiguous_table_ownership_fallback_to_text", issues)

    def make_output(self, root: Path) -> dict[str, object]:
        (root / "source.pdf").write_bytes(b"source fixture; no PDF parsing in this unit test")
        (root / "candidate.md").write_bytes(b"must\n\n")
        block = asdict(Block("text", [word(0, "must")], "must", (0, 0, 20, 20), end=6))
        with gzip.open(root / "evidence.jsonl.gz", "wt", encoding="utf-8") as out:
            out.write(json.dumps({"page": 1, "words": [word(0, "must")], "blocks": [block], "excluded": []})+"\n")
        report = {"status": "review_required", "pages": [{"page": 1}],
                  "source_sha256": sha256(root / "source.pdf"), "candidate_sha256": sha256(root / "candidate.md"),
                  "evidence_sha256": sha256(root / "evidence.jsonl.gz")}
        (root / "report.json").write_text(json.dumps(report), encoding="utf-8")
        return report

    def test_tampered_markdown_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            self.make_output(root)
            verify(root)
            (root / "candidate.md").write_text("may\n\n", encoding="utf-8")
            with self.assertRaisesRegex(ValueError, "Artifact changed"):
                verify(root)

    def test_incomplete_run_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / "RUNNING.json").write_text("{}")
            with self.assertRaisesRegex(ValueError, "Incomplete"):
                verify(root)

    def test_release_requires_all_page_approvals(self) -> None:
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            report = self.make_output(root)
            review = {**report, "reviewer": "Test reviewer", "pages": [{"page": 1, "approved": False, "note": ""}]}
            path = root / "review.json"
            path.write_text(json.dumps(review))
            with self.assertRaisesRegex(ValueError, "explicit approval"):
                release(root, path)
            self.assertFalse((root / "ready.md").exists())
            review["pages"] = [{"page": 1, "approved": True, "note": "Checked source against output."}]
            path.write_text(json.dumps(review))
            self.assertEqual(release(root, path), 0)
            self.assertEqual((root / "ready.md").read_bytes(), (root / "candidate.md").read_bytes())


if __name__ == "__main__":
    unittest.main()
