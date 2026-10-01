"""Create a static, offline visual review UI from verified conversion evidence."""
from __future__ import annotations

import argparse
import gzip
import html
import json
from pathlib import Path
from typing import Any

from pdf_to_markdown import emitted_text, text_of, unescape_text, verify

STYLE = """
body{font:15px system-ui;margin:20px;background:#f5f7fb;color:#172537}
nav{display:flex;gap:24px;align-items:center;padding:12px;background:white}
.panes{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px}
.pane{background:white;padding:16px;overflow:auto;max-height:85vh}
img{width:100%;height:auto} table{border-collapse:collapse;font-size:12px}
td,th{border:1px solid #9baabe;padding:8px;vertical-align:top;min-width:110px}
th{background:#e4ebf3} pre{white-space:pre-wrap;overflow-wrap:anywhere}
.issues{background:#fff0cc;padding:12px} .context{color:#576984;font-size:12px}
h1,h2,h3,h4,h5{font-size:18px} p{line-height:1.55} li{margin:6px}
@media(max-width:900px){.panes{grid-template-columns:1fr}}
"""


def preview(block: dict[str, Any]) -> str:
    """Render our known block schema, never execute source Markdown or HTML."""
    if block["kind"] == "table":
        meta = block["metadata"]
        words = {w["id"]: w for w in block["words"]}
        cells = {(c["row"], c["column"]): c for c in meta["cells"]}
        parts = ["<table><thead><tr>"]
        parts.extend(f"<th>Column {i+1}</th>" for i in range(meta["columns"]))
        parts.append("</tr></thead><tbody>")
        for r in range(meta["rows"]):
            parts.append("<tr>")
            for c in range(meta["columns"]):
                cell = cells[r, c]
                value = text_of(words[i] for i in cell["word_ids"])
                title = html.escape(str(cell["bbox"]), quote=True)
                parts.append(f'<td title="{title}">{html.escape(value)}</td>')
            parts.append("</tr>")
        parts.append("</tbody></table>")
        return "".join(parts)
    if block["kind"] == "heading":
        level = block["metadata"]["level"]
        title, _, body = block["markdown"].partition("\n\n")
        title = unescape_text(title[level+1:])
        return f"<h{level}>{html.escape(title)}</h{level}><p>{html.escape(unescape_text(body))}</p>"
    return "<p>" + html.escape(emitted_text(block)) + "</p>"


def build(output: Path) -> Path:
    """Create local HTML pages with paired source images and extracted content."""
    report = verify(output)
    destination = output / "review"
    links: list[str] = []
    with gzip.open(output / "evidence.jsonl.gz", "rt", encoding="utf-8") as stream:
        for line in stream:
            record = json.loads(line)
            page = record["page"]
            name = f"page-{page:04}.html"
            issues = ", ".join(record["issues"]) or "Otomatik işaret yok; görsel inceleme yine gereklidir."
            links.append(f'<li><a href="{name}">Sayfa {page}</a> — {html.escape(issues)}</li>')
            navigation = '<a href="index.html">Dizin</a>'
            if page > 1:
                navigation += f'<a href="page-{page-1:04}.html">Önceki</a>'
            if page < len(report["pages"]):
                navigation += f'<a href="page-{page+1:04}.html">Sonraki</a>'
            extracted = "".join(preview(b) for b in record["blocks"])
            markdown = html.escape("\n\n".join(b["markdown"] for b in record["blocks"]))
            excluded = html.escape("; ".join(f"{e['reason']}: {text_of(e['words'])}" for e in record["excluded"]))
            content = f'''<!doctype html><html lang="tr"><meta charset="utf-8">
<title>PDF kontrol — Sayfa {page}</title><style>{STYLE}</style>
<nav>{navigation}<strong>Sayfa {page} / {len(report['pages'])}</strong></nav>
<p class="issues">{html.escape(issues)}</p>
<div class="panes"><section class="pane"><h2>Kaynak PDF</h2>
<img src="page-{page:04}.png" alt="Kaynak sayfa {page}"></section>
<section class="pane"><h2>Çıkarılan içerik — onaylanmamış</h2>{extracted}
<details><summary>Markdown kaynağı</summary><pre>{markdown}</pre></details>
<p class="context">Markdown dışında tutulan kelimeler: {excluded or 'Yok'}</p>
</section></div></html>'''
            (destination / name).write_text(content, encoding="utf-8")
    index = f'''<!doctype html><html lang="tr"><meta charset="utf-8"><title>Yerel PDF kontrolü</title>
<style>{STYLE}</style><h1>Yerel PDF kontrolü</h1><p>Bu arayüz ağ bağlantısı veya JavaScript kullanmaz.
Görsel karşılaştırma içindir; otomatik onay vermez. İnceleme kararlarını review-template.json kopyasına kaydedin.</p>
<p>Kaynak SHA-256: {report['source_sha256']}</p><ol>{''.join(links)}</ol></html>'''
    (destination / "index.html").write_text(index, encoding="utf-8")
    return destination / "index.html"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    print(build(args.output).resolve())


if __name__ == "__main__":
    main()
