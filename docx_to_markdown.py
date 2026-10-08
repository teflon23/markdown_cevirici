"""Local DOCX -> auditable Markdown. No network, model, Word, or extra packages.

WordprocessingML text is accounted for exactly. Visual/semantic correctness
requires human review; generated candidates are never automatically released.
"""
from __future__ import annotations

import argparse
from collections import Counter
from contextlib import closing
from dataclasses import dataclass
import gzip
import hashlib
import html
import json
import logging
from pathlib import Path, PurePosixPath
import re
import shutil
import sqlite3
import sys
from typing import Any, Iterator
from urllib.parse import quote, urlsplit
import xml.etree.ElementTree as ET
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
TEXT_TAGS = {W+'t', W+'delText', W+'instrText', W+'delInstrText', A+'t', M+'t'}
LOG = logging.getLogger("docx_ingest")
BLOCKING_ISSUES = {'unresolved_numbering', 'unsupported_number_format',
                   'broken_note_reference', 'unsupported_embedded_content',
                   'textbox_or_equation_requires_review', 'wrapped_table_fallback',
                   'nested_table_flattened_requires_review', 'no_recognized_text_requires_review'}


def note_info(ancestors: list[tuple[str, dict[str, str]]]) -> dict[str, str] | None:
    """An ID is an identifier, never an implicit separator type."""
    for tag, attrs in ancestors:
        if tag in {W+'footnote', W+'endnote', W+'comment'}:
            return {'kind':tag.split('}')[1], 'id':attrs.get(W+'id',''),
                    'type':attrs.get(W+'type','normal')}
    return None


def story_target(part: str, ancestors: list[tuple[str, dict[str, str]]]) -> str:
    target = 'candidate' if part in {'word/document.xml','word/footnotes.xml','word/endnotes.xml'} else 'supplementary'
    note = note_info(ancestors)
    if note and note['kind'] != 'comment' and note['type'] != 'normal':
        target = 'supplementary'
    if any(tag in {W+'del', W+'moveFrom', W+'txbxContent', M+'oMath', M+'oMathPara'} or tag.endswith('}AlternateContent')
           for tag, _ in ancestors):
        target = 'supplementary'
    return target


def legal_heading(text: str) -> tuple[int,str] | None:
    """Anchored legal structure; generated headings do not replace source text."""
    clean = ' '.join(text.split())
    if re.fullmatch(r'[A-ZÇĞİÖŞÜ0-9 ]+ BÖLÜM',clean):
        return 1,clean
    if re.fullmatch(r'(?:CHAPTER|ANNEX|PART|TITLE)\s+[IVXLCDM\d]+',clean,re.I):
        return 1,clean
    article = re.match(r'((?:(?:GEÇİCİ|EK)\s+)?MADDE\s+\d+[A-Z]?)\s*[-–—:]',clean)
    if article:
        return 2,article[1]
    # Some exported documents put the short caption and article in one paragraph.
    captioned = re.match(r'^[A-Za-zÇĞİÖŞÜçğıöşü ]{1,80}?\s+((?:(?:GEÇİCİ|EK)\s+)?MADDE\s+\d+[A-Z]?)\s*[-–—:]\s*\(1\)',clean)
    if captioned:
        return 2,captioned[1]
    if re.fullmatch(r'Article\s+\d+(?:\.\d+)*',clean,re.I):
        return 2,clean
    return None


def prose_row(row: ET.Element, columns: int) -> bool:
    cells = row.findall(W+'tc')
    if len(cells) != 1 or int(val(cells[0],f'{W}tcPr/{W}gridSpan','1')) != columns:
        return False
    paragraphs = cells[0].findall(W+'p')
    return len(paragraphs) >= 3 and any(legal_heading(''.join(n.text or '' for n in p.iter(W+'t')))
                                        for p in paragraphs)


@dataclass(frozen=True)
class Limits:
    """Bound ZIP/XML expansion; never extract arbitrary archive paths."""
    part_bytes: int = 64 * 1024 * 1024
    total_bytes: int = 512 * 1024 * 1024
    entries: int = 10000
    compression_ratio: int = 1000


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open('rb') as stream:
        for data in iter(lambda: stream.read(1024 * 1024), b''):
            h.update(data)
    return h.hexdigest()


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')


def escape(value: str) -> str:
    """Preserve exact text, including whitespace, while disabling source markup."""
    value = html.escape(value, quote=False)
    value = re.sub(r'([\\`*_\[\]#|])', r'\\\1', value)
    value = re.sub(r'^(\s*\d+)\.(?=\s)', r'\1\\.', value)
    value = re.sub(r'^(\s*)([-+])(?=\s)', r'\1\\\2', value)
    return value.replace('\r', '&#13;').replace('\n', '<br>').replace('\t', '&#9;')


def val(node: ET.Element, path: str, default: str = '') -> str:
    element = node.find(path)
    return element.get(W+'val', default) if element is not None else default


def enabled(node: ET.Element | None) -> bool:
    return node is not None and node.get(W+'val', '1') not in {'0', 'false', 'off'}


def validate_package(z: ZipFile, limits: Limits) -> list[str]:
    """Reject duplicates, traversal, excessive expansion, DTDs and entities."""
    infos = z.infolist()
    names = [i.filename for i in infos]
    if len(names) != len(set(names)) or len(names) > limits.entries:
        raise ValueError('Duplicate ZIP entries or too many parts')
    if sum(i.file_size for i in infos) > limits.total_bytes:
        raise ValueError('DOCX exceeds total uncompressed size limit')
    for item in infos:
        path = PurePosixPath(item.filename)
        if path.is_absolute() or '..' in path.parts or '\\' in item.filename or ':' in item.filename:
            raise ValueError('Unsafe package path')
        if item.flag_bits & 1 or item.file_size > limits.part_bytes:
            raise ValueError('Encrypted or oversized package part')
        if item.file_size / max(1, item.compress_size) > limits.compression_ratio:
            raise ValueError('Suspicious ZIP expansion ratio')
        if item.filename.endswith(('.xml', '.rels')):
            with z.open(item) as stream:
                tail = b''
                for chunk in iter(lambda: stream.read(65536), b''):
                    data = (tail + chunk).upper().replace(b'\x00', b'')
                    if b'<!DOCTYPE' in data or b'<!ENTITY' in data:
                        raise ValueError('DTD/entity declarations are forbidden')
                    tail = chunk[-64:]
    if 'word/document.xml' not in names or '[Content_Types].xml' not in names:
        raise ValueError('Not a supported DOCX package')
    with z.open('word/document.xml') as stream:
        for _, root in ET.iterparse(stream, events=('start',)):
            if root.tag != W+'document':
                raise ValueError('Unsupported DOCX XML namespace (Strict OOXML is not supported)')
            break
    return names


def xml(z: ZipFile, name: str) -> ET.Element:
    if name not in z.namelist():
        return ET.Element('missing')
    with z.open(name) as stream:
        return ET.parse(stream).getroot()


def inventory(z: ZipFile, db: sqlite3.Connection) -> int:
    """Independent streaming inventory of every recognized text node in all XML."""
    db.execute('CREATE TABLE atoms (id TEXT PRIMARY KEY, value TEXT NOT NULL, used INTEGER NOT NULL DEFAULT 0)')
    db.execute('CREATE TABLE positions (part TEXT PRIMARY KEY, last INTEGER NOT NULL)')
    total = 0
    for part in sorted(n for n in z.namelist() if n.endswith('.xml')):
        count = 0
        stack: list[ET.Element] = []
        with z.open(part) as stream:
            for event, node in ET.iterparse(stream, events=('start','end')):
                if event == 'start':
                    stack.append(node)
                    continue
                if node.tag in TEXT_TAGS:
                    count += 1
                    db.execute('INSERT INTO atoms(id,value) VALUES (?,?)',
                               (f'{part}#{count}', node.text or ''))
                    total += 1
                stack.pop()
                if stack:
                    stack[-1].remove(node)
                node.clear()
    db.commit()
    return total


def blocks(z: ZipFile, part: str) -> Iterator[tuple[ET.Element, dict[int, str], list[tuple[str, dict[str, str]]]]]:
    """Yield outer paragraphs/tables, retaining at most one logical block.

    Tables can be large; ZIP limits bound the largest XML part. Text not inside
    a paragraph/table is yielded individually so orphan/unknown text is audited.
    """
    stack: list[ET.Element] = []
    active: ET.Element | None = None
    identifiers: dict[int, str] = {}
    count = 0
    with z.open(part) as stream:
        for event, node in ET.iterparse(stream, events=('start', 'end')):
            if event == 'start':
                stack.append(node)
                if active is None and node.tag in {W+'p', W+'tbl'}:
                    active = node
                continue
            if node.tag in TEXT_TAGS:
                count += 1
                identifiers[id(node)] = f'{part}#{count}'
                if active is None:
                    yield node, identifiers.copy(), [(p.tag, dict(p.attrib)) for p in stack[:-1]]
                    identifiers.clear()
            if node is active:
                yield node, identifiers.copy(), [(p.tag, dict(p.attrib)) for p in stack[:-1]]
                active = None
                identifiers.clear()
            elif active is None and node.tag in {W+'altChunk', W+'object', W+'sym'}:
                yield node, {}, [(p.tag, dict(p.attrib)) for p in stack[:-1]]
            stack.pop()
            if active is None:
                if stack:
                    stack[-1].remove(node)
                node.clear()


class Structure:
    """Resolve inherited paragraph styles and common Word list definitions."""
    def __init__(self, z: ZipFile) -> None:
        styles = xml(z, 'word/styles.xml')
        self.styles = {s.get(W+'styleId', ''): s for s in styles.findall(W+'style')}
        self.default = next((key for key, s in self.styles.items()
                             if s.get(W+'type') == 'paragraph' and s.get(W+'default') == '1'), '')
        numbering = xml(z, 'word/numbering.xml')
        self.abstract = {n.get(W+'abstractNumId'): n for n in numbering.findall(W+'abstractNum')}
        self.nums = {n.get(W+'numId'): n for n in numbering.findall(W+'num')}
        self.counters: dict[tuple[str, str], dict[int, int]] = {}
        self.rels_cache: dict[str, dict[str | None, dict[str, str]]] = {}
        self.notes: dict[str, set[str]] = {}
        for kind, part in [('footnote','word/footnotes.xml'),('endnote','word/endnotes.xml'),('comment','word/comments.xml')]:
            self.notes[kind] = {n.get(W+'id','') for n in xml(z,part).findall(W+kind)
                                if kind == 'comment' or n.get(W+'type','normal') == 'normal'}

    def properties(self, p: ET.Element) -> list[ET.Element]:
        props: list[ET.Element] = []
        direct = p.find(W+'pPr')
        if direct is not None:
            props.append(direct)
        style = val(p, f'{W}pPr/{W}pStyle', self.default)
        visited: set[str] = set()
        while style and style not in visited and style in self.styles:
            visited.add(style)
            element = self.styles[style]
            pr = element.find(W+'pPr')
            if pr is not None:
                props.append(pr)
            style = val(element, W+'basedOn')
        return props

    def heading(self, p: ET.Element) -> int | None:
        for prop in self.properties(p):
            outline = val(prop, W+'outlineLvl')
            if outline:
                return min(6, int(outline)+1) if int(outline) < 9 else None
        style = val(p, f'{W}pPr/{W}pStyle')
        if style.lower() == 'title':
            return 1
        match = re.fullmatch(r'Heading([1-9])', style, re.I)
        return min(6, int(match[1])) if match else None

    def definition(self, num: ET.Element, level: int) -> ET.Element | None:
        override = num.find(f"{W}lvlOverride[@{W}ilvl='{level}']")
        if override is not None and override.find(W+'lvl') is not None:
            return override.find(W+'lvl')
        abstract = self.abstract.get(val(num, W+'abstractNumId'))
        return abstract.find(f"{W}lvl[@{W}ilvl='{level}']") if abstract is not None else None

    def marker(self, p: ET.Element, story: str, issues: set[str]) -> tuple[str, int] | None:
        num_id = level_value = ''
        for prop in self.properties(p):
            num_id = num_id or val(prop, f'{W}numPr/{W}numId')
            level_value = level_value or val(prop, f'{W}numPr/{W}ilvl')
        if not num_id or num_id == '0':
            return None
        issues.add('numbering_requires_review')
        level = int(level_value or '0')
        if not 0 <= level <= 8:
            raise ValueError('Invalid numbering level')
        num = self.nums.get(num_id)
        definition = self.definition(num, level) if num is not None else None
        if definition is None:
            issues.add('unresolved_numbering')
            return '[unresolved list marker]', level
        counts = self.counters.setdefault((story, num_id), {})
        def start(index: int) -> int:
            item = self.definition(num, index)
            override = val(num, f"{W}lvlOverride[@{W}ilvl='{index}']/{W}startOverride")
            return int(override or (val(item, W+'start', '1') if item is not None else '1'))
        counts[level] = counts.get(level, start(level)-1) + 1
        for deeper in list(counts):
            if deeper > level:
                item = self.definition(num, deeper)
                restart = int(val(item, W+'lvlRestart', str(deeper))) if item is not None else deeper
                if restart and level <= restart-1:
                    del counts[deeper]
        pattern = val(definition, W+'lvlText', f'%{level+1}.')
        legal = enabled(definition.find(W+'isLgl'))
        def replace(match: re.Match[str]) -> str:
            index = int(match[1])-1
            item = self.definition(num, index)
            fmt = 'decimal' if legal else (val(item, W+'numFmt', 'decimal') if item is not None else 'decimal')
            return number(counts.get(index, start(index)), fmt, issues)
        return re.sub(r'%([1-9])', replace, pattern), level


def number(value: int, fmt: str, issues: set[str]) -> str:
    if fmt in {'decimal', 'decimalZero'}:
        return str(value).zfill(2) if fmt == 'decimalZero' else str(value)
    if fmt in {'lowerLetter', 'upperLetter'} and value > 0:
        result = ''
        while value:
            value, rest = divmod(value-1, 26)
            result = chr(97+rest) + result
        return result.upper() if fmt == 'upperLetter' else result
    if fmt in {'lowerRoman', 'upperRoman'} and 0 < value < 4000:
        result = ''
        for n, token in [(1000,'M'),(900,'CM'),(500,'D'),(400,'CD'),(100,'C'),(90,'XC'),
                         (50,'L'),(40,'XL'),(10,'X'),(9,'IX'),(5,'V'),(4,'IV'),(1,'I')]:
            while value >= n:
                result += token
                value -= n
        return result.lower() if fmt == 'lowerRoman' else result
    issues.add('unsupported_number_format')
    return str(value)


class Emitter:
    """Keep source fragments separate from generated Markdown structure."""
    def __init__(self, z: ZipFile, part: str, ids: dict[int, str], structure: Structure,
                 assets: dict[str, str], default_target: str, ancestors: list[tuple[str, dict[str, str]]]) -> None:
        self.part, self.ids, self.structure, self.assets = part, ids, structure, assets
        self.fragments: list[dict[str, Any]] = []
        self.issues: set[str] = set()
        self.metadata: dict[str, Any] = {'ancestors': ancestors}
        self.default_target = default_target
        relpart = str(PurePosixPath(part).parent / '_rels' / (PurePosixPath(part).name + '.rels'))
        if part not in structure.rels_cache:
            structure.rels_cache[part] = {r.get('Id'): dict(r.attrib) for r in xml(z, relpart)}
        self.rels = structure.rels_cache[part]
        self.parents: dict[int, ET.Element] = {}

    def paragraph_prefix(self, node: ET.Element, target: str, inline: bool) -> None:
        """One numbering event for EVERY paragraph, including cells and textboxes."""
        heading = self.structure.heading(node)
        marker = self.structure.marker(node, self.part+':'+target, self.issues)
        spec = {'heading':heading, 'label':marker[0] if marker else None,
                'level':marker[1] if marker else None, 'target':target}
        index = len(self.metadata.setdefault('paragraphs',[]))
        self.metadata['paragraphs'].append(spec)
        if marker:
            label, level = marker
            markup = (escape(label)+' ') if heading or inline else ('    '*level+'- **'+escape(label)+'** ')
            self.fragments.append({'markup':markup,'target':target,'semantic':'numbering',
                                   'paragraph':index,'label':label})
            if not inline:
                self.metadata['numbering'] = {'label':label,'level':level}

    def generated(self, value: str, target: str | None = None) -> None:
        self.fragments.append({'markup': value, 'target': target or self.default_target})

    def emit(self, node: ET.Element, target: str | None = None) -> None:
        target = target or self.default_target
        tag = node.tag
        if tag == W+'p':
            self.paragraph_prefix(node,target,inline=True)
        if tag in {W+'del', W+'moveFrom'}:
            self.issues.add('tracked_changes_require_review')
            target = 'supplementary'
            self.generated('\n\n[Deleted or moved-from text]\n\n', target)
        if tag in {W+'ins', W+'moveTo'}:
            self.issues.add('tracked_changes_require_review')
        if tag in {W+'txbxContent', M+'oMath', M+'oMathPara'}:
            self.issues.add('textbox_or_equation_requires_review')
            target = 'supplementary'
            self.generated('\n\n[Textbox or equation text; original structure requires review]\n\n', target)
        if tag in TEXT_TAGS:
            if tag in {W+'delText', W+'instrText', W+'delInstrText', A+'t', M+'t'}:
                target = 'supplementary'
                self.issues.add('non_body_text_retained_separately')
            if tag in {W+'instrText', W+'delInstrText'}:
                self.generated('\n\n[Field instruction]\n\n', target)
            self.fragments.append({'id': self.ids[id(node)], 'text': node.text or '', 'target': target})
            return
        if tag in {W+'br', W+'cr', W+'tab', W+'noBreakHyphen', W+'softHyphen'}:
            value = {W+'br':'<br>', W+'cr':'<br>', W+'tab':'&#9;', W+'noBreakHyphen':'‑', W+'softHyphen':'\u00ad'}[tag]
            self.generated(value, target)
            return
        if tag in {W+'footnoteReference', W+'endnoteReference', W+'commentReference'}:
            kind = tag.split('}')[1].replace('Reference', '')
            identifier = node.get(W+'id', '')
            if identifier not in self.structure.notes.get(kind,set()):
                self.issues.add('broken_note_reference')
            self.metadata.setdefault('references', []).append({'kind': kind, 'id': identifier})
            destination = 'supplementary.md' if kind == 'comment' else ('candidate.md' if target == 'supplementary' else '')
            self.fragments.append({'markup':f' [{kind} {escape(identifier)}]({destination}#{kind}-{quote(identifier)}) ',
                                   'target':target,'semantic':'note_reference','kind':kind,'note_id':identifier})
        if tag == W+'hyperlink':
            relation = self.rels.get(node.get(R+'id'), {})
            url = relation.get('Target', '') if relation.get('TargetMode') == 'External' else ''
            anchor = node.get(W+'anchor')
            self.metadata.setdefault('hyperlinks', []).append({'url': url, 'anchor': anchor})
            allowed = urlsplit(url).scheme.lower() in {'https', 'http', 'mailto'}
            if allowed:
                self.generated('[', target)
            else:
                self.issues.add('internal_or_unsupported_hyperlink')
            for child in node:
                self.emit(child, target)
            if allowed:
                self.generated('](' + quote(url, safe=':/?#@&=+%') + ')', target)
            return
        if tag == W+'r':
            prop = node.find(W+'rPr')
            if prop is not None:
                if enabled(prop.find(W+'vanish')) or enabled(prop.find(W+'webHidden')):
                    target = 'supplementary'
                    self.issues.add('hidden_text_requires_review')
                if enabled(prop.find(W+'strike')) or enabled(prop.find(W+'dstrike')):
                    self.issues.add('strikethrough_requires_review')
                marker = ('**' if enabled(prop.find(W+'b')) else '') + ('*' if enabled(prop.find(W+'i')) else '')
            else:
                marker = ''
            if marker:
                self.fragments.append({'markup':marker,'target':target,'format_edge':'open'})
            for child in node:
                self.emit(child, target)
            if marker:
                self.fragments.append({'markup':marker,'target':target,'format_edge':'close'})
            return
        if tag == A+'blip':
            rid = node.get(R+'embed') or node.get(R+'link')
            relation = self.rels.get(rid, {})
            self.issues.add('image_requires_visual_review')
            if relation.get('TargetMode') == 'External':
                self.metadata.setdefault('external_images', []).append(relation.get('Target', ''))
            else:
                import posixpath
                path = posixpath.normpath(posixpath.join(str(PurePosixPath(self.part).parent), relation.get('Target',''))).lstrip('/')
                if path in self.assets:
                    self.generated(f' ![Image]({self.assets[path]}) ', target)
        if tag in {W+'fldChar', W+'fldSimple'}:
            self.issues.add('field_cached_result_requires_review')
            self.metadata.setdefault('fields', []).append(dict(node.attrib))
        if tag in {W+'sym', W+'object', W+'altChunk', W+'pict'}:
            self.issues.add('unsupported_embedded_content')
            self.metadata.setdefault('unsupported', []).append({'tag': tag, 'attributes': dict(node.attrib)})
        if tag.endswith('}AlternateContent'):
            self.issues.add('alternate_content_requires_review')
            target = 'supplementary'
        if tag == W+'tbl' and node is not self.metadata.get('_root'):
            self.issues.add('nested_table_flattened_requires_review')
        for child in node:
            self.emit(child, target)
        if tag == W+'p':
            self.generated('<br>', target)

    def paragraph(self, node: ET.Element) -> None:
        heading = self.structure.heading(node)
        raw = ''.join(n.text or '' for n in node.iter(W+'t'))
        legal = legal_heading(raw)
        if legal:
            level,title = legal
            markup = '#'*level+' ' if raw.strip()==title else '#'*level+' '+escape(title)+'\n\n'
            self.fragments.append({'markup':markup,
                                   'target':self.default_target,'semantic':'legal_heading',
                                   'level':level,'title':title})
            self.metadata.setdefault('legal_headings',[]).append({'level':level,'title':title})
            heading = None
        if heading:
            self.fragments.append({'markup':'#'*heading+' ','target':self.default_target,
                                   'semantic':'heading','level':heading})
            self.metadata['heading_level'] = heading
        self.paragraph_prefix(node,self.default_target,inline=False)
        for child in node:
            self.emit(child)
        self.generated('\n\n')

    def block(self, node: ET.Element) -> dict[str, Any]:
        self.parents = {id(child): parent for parent in node.iter() for child in parent}
        for tag, _ in self.metadata['ancestors']:
            if tag in {W+'del', W+'moveFrom'}:
                self.default_target = 'supplementary'
                self.issues.add('tracked_changes_require_review')
            elif tag in {W+'ins', W+'moveTo'}:
                self.issues.add('tracked_changes_require_review')
        if node.tag == W+'tbl':
            self.table(node)
        elif node.tag == W+'p':
            self.paragraph(node)
        else:
            self.issues.add('orphan_text_retained')
            self.emit(node, 'supplementary')
            self.generated('\n\n', 'supplementary')
        seen = [f['id'] for f in self.fragments if 'id' in f]
        expected = [self.ids[id(n)] for n in node.iter() if n.tag in TEXT_TAGS]
        if seen != expected:
            raise ValueError('Block source text order/coverage failure')
        return {'part': self.part, 'kind': 'table' if node.tag == W+'tbl' else 'paragraph',
                'fragments': self.fragments, 'metadata': self.metadata, 'issues': sorted(self.issues)}

    def table(self, node: ET.Element) -> None:
        """Emit direct grid cells only; preserve merge topology without duplication."""
        self.issues.add('table_requires_visual_review')
        rows = node.findall(W+'tr')
        # Wrapped rows/cells cannot be guessed safely: preserve their text linearly.
        direct_ids = {id(n) for row in rows for cell in row.findall(W+'tc') for n in cell.iter() if n.tag in TEXT_TAGS}
        all_ids = {id(n) for n in node.iter() if n.tag in TEXT_TAGS}
        if direct_ids != all_ids or not rows:
            self.issues.add('wrapped_table_fallback')
            self.emit(node)
            self.generated('\n\n')
            return
        widths = [int(val(row, f'{W}trPr/{W}gridBefore', '0')) +
                  sum(int(val(c, f'{W}tcPr/{W}gridSpan', '1')) for c in row.findall(W+'tc')) +
                  int(val(row, f'{W}trPr/{W}gridAfter', '0')) for row in rows]
        columns = max(widths)
        if not 1 <= columns <= 256:
            raise ValueError('Unsupported table grid width')
        table_open = False
        topology: list[dict[str, Any]] = []
        for r, row in enumerate(rows):
            unwrap = prose_row(row,columns)
            if unwrap:
                self.issues.add('full_width_prose_row_unwrapped_requires_review')
                self.metadata.setdefault('unwrapped_rows',[]).append(r)
                self.generated('\n\n')
                table_open = False
            elif not table_open:
                self.generated('| '+' | '.join(f'Column {n+1}' for n in range(columns))+' |\n')
                self.generated('| '+' | '.join('---' for _ in range(columns))+' |\n')
                table_open = True
            column = int(val(row, f'{W}trPr/{W}gridBefore', '0'))
            if not unwrap:
                self.generated('| ' + ' | '*column)
            for cell in row.findall(W+'tc'):
                span = int(val(cell, f'{W}tcPr/{W}gridSpan', '1'))
                if not 1 <= span <= 256:
                    raise ValueError('Invalid grid span')
                merge = cell.find(f'{W}tcPr/{W}vMerge')
                topology.append({'row': r, 'column': column, 'span': span,
                                 'vmerge': merge.get(W+'val', 'continue') if merge is not None else None,
                                 'atom_ids': [self.ids[id(n)] for n in cell.iter() if n.tag in TEXT_TAGS]})
                if span != 1 or merge is not None:
                    self.issues.add('merged_table_cells')
                if unwrap:
                    for child in cell:
                        if child.tag == W+'p':
                            self.paragraph(child)
                        else:
                            self.emit(child)
                else:
                    self.emit(cell)
                    self.generated(' | '*span)
                column += span
            if not unwrap:
                self.generated(' | '*(columns-column)+'\n')
        self.generated('\n')
        self.metadata['table'] = {'columns': columns, 'rows': len(rows), 'cells': topology,
                                  'generated_header': True}


def render(record: dict[str, Any], target: str) -> str:
    """Coalesce equal adjacent Word run formatting without changing source text."""
    selected = [f for f in record['fragments'] if f['target']==target]
    filtered = []
    index = 0
    while index < len(selected):
        f = selected[index]
        if (f.get('format_edge')=='close' and index+1<len(selected) and
                selected[index+1].get('format_edge')=='open' and f['markup']==selected[index+1]['markup']):
            index += 2
            continue
        filtered.append(f)
        index += 1
    output: list[str] = []
    active: str | None = None
    content: list[str] = []
    for f in filtered:
        if f.get('format_edge')=='open':
            if active is not None:
                raise ValueError('Overlapping formatting spans')
            active = f['markup']
            content = []
        elif f.get('format_edge')=='close':
            if active != f['markup']:
                raise ValueError('Unbalanced formatting spans')
            value = ''.join(content)
            core = value.strip(' ')
            before = value[:len(value)-len(value.lstrip(' '))]
            after = value[len(value.rstrip(' ')):]
            output.append(before+active+core+active+after if core else value)
            active = None
        else:
            value = escape(f['text']) if 'id' in f else f['markup']
            (content if active is not None else output).append(value)
    if active is not None:
        raise ValueError('Unclosed formatting span')
    return ''.join(output)


def account(db: sqlite3.Connection, record: dict[str, Any]) -> None:
    for f in record['fragments']:
        if 'id' not in f:
            continue
        source = db.execute('SELECT value,used FROM atoms WHERE id=?', (f['id'],)).fetchone()
        if source is None or source[0] != f['text'] or source[1] != 0:
            raise ValueError('Source text changed, duplicated or invented: '+f['id'])
        part, index = f['id'].rsplit('#',1)
        previous = db.execute('SELECT last FROM positions WHERE part=?',(part,)).fetchone()
        if int(index) != (previous[0] if previous else 0)+1:
            raise ValueError('Source text order changed: '+f['id'])
        db.execute('INSERT OR REPLACE INTO positions(part,last) VALUES (?,?)',(part,int(index)))
        db.execute('UPDATE atoms SET used=1 WHERE id=?', (f['id'],))


def structural_expectations(z: ZipFile) -> Iterator[dict[str, Any]]:
    """Independent traversal: derive paragraph/list/note/grid facts from XML.

    Number-format resolution is shared and separately regression-tested. This
    pass does not trust emitted metadata or infer structure from output text.
    """
    resolver = Structure(z)
    parts = sorted(n for n in z.namelist() if n.endswith('.xml'))
    parts.remove('word/document.xml')
    parts.insert(0,'word/document.xml')
    for part in parts:
        for root, ids, ancestors in blocks(z,part):
            parents = {id(child):parent for parent in root.iter() for child in parent}
            def lineage(n: ET.Element) -> list[tuple[str,dict[str,str]]]:
                chain: list[tuple[str,dict[str,str]]] = []
                parent = parents.get(id(n))
                while parent is not None:
                    chain.append((parent.tag,dict(parent.attrib)))
                    parent = parents.get(id(parent))
                return ancestors+list(reversed(chain))
            paragraphs = []
            for p in root.iter(W+'p'):
                target = story_target(part,lineage(p))
                label = resolver.marker(p,part+':'+target,set())
                paragraphs.append({'heading':resolver.heading(p),'label':label[0] if label else None,
                                   'level':label[1] if label else None,'target':target})
            destinations = {}
            for n in root.iter():
                if n.tag not in TEXT_TAGS:
                    continue
                destination = story_target(part,lineage(n))
                if n.tag in {W+'delText',W+'instrText',W+'delInstrText',A+'t',M+'t'}:
                    destination = 'supplementary'
                parent = parents.get(id(n))
                while parent is not None:
                    if parent.tag == W+'r':
                        pr = parent.find(W+'rPr')
                        if pr is not None and (enabled(pr.find(W+'vanish')) or enabled(pr.find(W+'webHidden'))):
                            destination = 'supplementary'
                    parent = parents.get(id(parent))
                destinations[ids[id(n)]] = destination
            references = [{'kind':n.tag.split('}')[1].replace('Reference',''),'id':n.get(W+'id','')}
                          for n in root.iter() if n.tag in {W+'footnoteReference',W+'endnoteReference',W+'commentReference'}]
            table = None
            unwrapped = []
            heading_paragraphs = [root] if root.tag==W+'p' else []
            if root.tag == W+'tbl':
                rows = root.findall(W+'tr')
                cells = []
                widths = []
                for r,row in enumerate(rows):
                    col = int(val(row,f'{W}trPr/{W}gridBefore','0'))
                    for tc in row.findall(W+'tc'):
                        span = int(val(tc,f'{W}tcPr/{W}gridSpan','1'))
                        merge = tc.find(f'{W}tcPr/{W}vMerge')
                        cells.append({'row':r,'column':col,'span':span,
                                      'vmerge':merge.get(W+'val','continue') if merge is not None else None,
                                      'atom_ids':[ids[id(n)] for n in tc.iter() if n.tag in TEXT_TAGS]})
                        col += span
                    widths.append(col+int(val(row,f'{W}trPr/{W}gridAfter','0')))
                if rows:
                    table = {'columns':max(widths),'rows':len(rows),'cells':cells,'generated_header':True}
                    unwrapped = [i for i,row in enumerate(rows) if prose_row(row,table['columns'])]
                    heading_paragraphs.extend(p for i in unwrapped for tc in rows[i].findall(W+'tc') for p in tc.findall(W+'p'))
            legal = [{'level':h[0],'title':h[1]} for p in heading_paragraphs
                     if (h:=legal_heading(''.join(n.text or '' for n in p.iter(W+'t'))))]
            yield {'part':part,'paragraphs':paragraphs,'references':references,'table':table,
                   'note':note_info(ancestors),'target':story_target(part,ancestors),
                   'destinations':destinations,'legal_headings':legal,'unwrapped_rows':unwrapped}


def audit_structure(record: dict[str, Any], expected: dict[str, Any]) -> None:
    """A missing generated list marker must fail even when all text survived."""
    meta = record['metadata']
    if record['part'] != expected['part'] or meta.get('paragraphs',[]) != expected['paragraphs']:
        raise ValueError('Source paragraph/list structure mismatch')
    if meta.get('note') != expected['note']:
        raise ValueError('Source note type/identity mismatch')
    for f in record['fragments']:
        if 'id' in f and f['target'] != expected['destinations'].get(f['id']):
            raise ValueError('Source text destination mismatch')
    if meta.get('unwrapped_rows',[]) != expected['unwrapped_rows']:
        raise ValueError('Layout table row classification mismatch')
    if meta.get('legal_headings',[]) != expected['legal_headings']:
        raise ValueError('Legal heading structure mismatch')
    heading_fragments = [f for f in record['fragments'] if f.get('semantic')=='legal_heading']
    if len(heading_fragments) != len(expected['legal_headings']):
        raise ValueError('Generated legal heading missing')
    for f,h in zip(heading_fragments,expected['legal_headings']):
        if (f.get('level')!=h['level'] or f.get('title')!=h['title'] or
                f['markup'] not in {'#'*h['level']+' ', '#'*h['level']+' '+escape(h['title'])+'\n\n'}):
            raise ValueError('Generated legal heading mismatch')
    if meta.get('references',[]) != expected['references']:
        raise ValueError('Source note reference mismatch')
    if expected['table'] and 'wrapped_table_fallback' not in record['issues'] and meta.get('table') != expected['table']:
        raise ValueError('Source table grid mismatch')
    numbered = [f for f in record['fragments'] if f.get('semantic') == 'numbering']
    expected_labels = [(i,p) for i,p in enumerate(expected['paragraphs']) if p['label'] is not None]
    if len(numbered) != len(expected_labels):
        raise ValueError('Generated numbering marker missing or duplicated')
    for actual,(index,p) in zip(numbered,expected_labels):
        forms = {escape(p['label'])+' ', '    '*p['level']+'- **'+escape(p['label'])+'** '}
        if actual.get('paragraph') != index or actual.get('label') != p['label'] or actual['markup'] not in forms or actual['target'] != p['target']:
            raise ValueError('Generated numbering marker differs from source')
    reference_fragments = [f for f in record['fragments'] if f.get('semantic') == 'note_reference']
    if len(reference_fragments) != len(expected['references']):
        raise ValueError('Generated note link missing')
    for f,ref in zip(reference_fragments,expected['references']):
        kind, identifier = ref['kind'],ref['id']
        destination = 'supplementary.md' if kind=='comment' else ('candidate.md' if f['target']=='supplementary' else '')
        if f.get('kind') != kind or f.get('note_id') != identifier or f['markup'] != f' [{kind} {escape(identifier)}]({destination}#{kind}-{quote(identifier)}) ':
            raise ValueError('Generated note link differs from source')
    note = expected['note']
    if note:
        for fragment in record['fragments']:
            # Deletions/fields may intentionally move to the supplement; a normal
            # note's ordinary text must not be moved wholesale by its numeric ID.
            if 'id' in fragment and note['type'] != 'normal' and note['kind'] != 'comment' and fragment['target'] != 'supplementary':
                raise ValueError('Separator note incorrectly included in body')


def build_review(output: Path) -> None:
    """Static escaped-text review; does not pretend to render Word pagination."""
    with (output/'review.html').open('w',encoding='utf-8') as out, \
            gzip.open(output/'evidence.jsonl.gz','rt',encoding='utf-8') as evidence:
        out.write('''<!doctype html><html lang="tr"><meta charset="utf-8"><title>DOCX kontrolü</title>
<style>body{font:15px system-ui;margin:24px;color:#183047;background:#f5f7fb}
section{background:white;padding:20px;margin:20px 0;border:1px solid #ccd6df}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}pre{white-space:pre-wrap;overflow-wrap:anywhere}
.warning{color:#795500}details{margin:12px 0}@media(max-width:800px){.grid{display:block}}</style>
<h1>DOCX metin denetimi</h1><p>Bu ekran XML metni ile üretilen Markdown'ı gösterir; Word sayfa düzenini render etmez.
Görsel doğrulama için source.docx dosyasını yerel Word/LibreOffice uygulamanızda ayrıca inceleyin.</p>''')
        for line in evidence:
            record = json.loads(line)
            n = record['block']
            raw = ''.join(f['text'] for f in record['fragments'] if 'id' in f)
            md = render(record,'candidate')
            supplementary = render(record,'supplementary')
            out.write(f'<section id="block-{n}"><h2>Blok {n} — {html.escape(record["part"])}</h2>')
            out.write('<p class="warning">'+html.escape(', '.join(record['issues']) or 'Otomatik uyarı yok; inceleme gerekli.')+'</p>')
            out.write('<div class="grid"><div><h3>Kaynak metin parçaları</h3><pre>'+html.escape(raw)+
                      '</pre></div><div><h3>Aday Markdown</h3><pre>'+html.escape(md)+'</pre></div></div>')
            if supplementary:
                out.write('<details open><summary>Ayrı korunan içerik</summary><pre>'+html.escape(supplementary)+'</pre></details>')
            out.write('<details><summary>Yapı ve kaynak bilgileri</summary><pre>'+
                      html.escape(json.dumps(record['metadata'],ensure_ascii=False,indent=2))+'</pre></details></section>')
        out.write('</html>')


def convert(source: Path, output: Path, limits: Limits = Limits()) -> int:
    """Stream logical blocks, preserve original package, and require review."""
    source = source.resolve(strict=True)
    with ZipFile(source) as check:
        validate_package(check, limits)
    output.mkdir(parents=True, exist_ok=False)
    write_json(output/'RUNNING.json', {'source': str(source)})
    shutil.copyfile(source, output/'source.docx')
    (output/'assets').mkdir()
    issue_counts: Counter[str] = Counter()
    assets: dict[str, str] = {}
    hashes: dict[str, str] = {}
    block_count = 0
    text_counts: Counter[str] = Counter()
    contexts: dict[str, dict[str, Any]] = {}
    dbpath = output/'audit.sqlite3'
    with ZipFile(output/'source.docx') as z, closing(sqlite3.connect(dbpath)) as db, \
            gzip.open(output/'evidence.jsonl.gz', 'wt', encoding='utf-8') as evidence, \
            (output/'candidate.md').open('wb') as candidate, (output/'supplementary.md').open('wb') as supplement:
        names = validate_package(z, limits)
        atom_count = inventory(z, db)
        structure = Structure(z)
        expected_blocks = structural_expectations(z)
        for name in names:
            if name.startswith('word/media/') and not name.endswith('/'):
                suffix = PurePosixPath(name).suffix.lower()
                suffix = suffix if suffix in {'.png','.jpg','.jpeg','.gif','.webp'} else '.bin'
                asset = 'assets/'+hashlib.sha256(name.encode()).hexdigest()[:20]+suffix
                with z.open(name) as src, (output/asset).open('wb') as dst:
                    shutil.copyfileobj(src, dst)
                hashes[asset] = digest(output/asset)
                if suffix != '.bin':
                    assets[name] = asset
                issue_counts['media_preserved_requires_review'] += 1
            elif name.startswith('word/embeddings/') or name.endswith('vbaProject.bin'):
                issue_counts['embedded_binary_retained_in_source'] += 1
        parts = sorted(n for n in names if n.endswith('.xml'))
        parts.remove('word/document.xml')
        parts.insert(0, 'word/document.xml')
        for part in parts:
            known = (part in {'[Content_Types].xml','word/document.xml','word/styles.xml','word/numbering.xml',
                             'word/settings.xml','word/webSettings.xml','word/fontTable.xml',
                             'word/footnotes.xml','word/endnotes.xml','word/comments.xml'} or
                     re.fullmatch(r'word/(?:header|footer)\d+\.xml',part) or
                     part.startswith(('docProps/','word/theme/')))
            if not known:
                issue_counts['additional_xml_part_retained_in_source'] += 1
            announced: set[str] = set()
            for node, ids, ancestors in blocks(z, part):
                default = story_target(part,ancestors)
                emitter = Emitter(z, part, ids, structure, assets, default, ancestors)
                note = note_info(ancestors)
                emitter.metadata['note'] = note
                key = note['kind']+'-'+note['id'] if note else part
                if key not in announced:
                    announced.add(key)
                    title = f'{note["kind"]} {note["id"]}' if note else part
                    if part != 'word/document.xml':
                        emitter.generated('\n\n## '+escape(title)+'\n\n')
                record = emitter.block(node)
                context = contexts.setdefault(part, {'headings':{},'list':{}})
                heading = record['metadata'].get('heading_level')
                if heading:
                    context['headings'] = {k:v for k,v in context['headings'].items() if int(k)<heading}
                    context['headings'][str(heading)] = ''.join(f['text'] for f in record['fragments']
                                                               if 'id' in f and f['target']=='candidate')
                    context['list'] = {}
                marker = record['metadata'].get('numbering')
                if marker:
                    level = marker['level']
                    context['list'] = {k:v for k,v in context['list'].items() if int(k)<level}
                    context['list'][str(level)] = marker['label']
                record['metadata']['context'] = {'headings':dict(context['headings']), 'list':dict(context['list'])}
                audit_structure(record,next(expected_blocks))
                block_count += 1
                record['block'] = block_count
                if any(f['target'] == 'supplementary' and 'id' in f for f in record['fragments']):
                    record['fragments'].insert(0, {'target':'supplementary',
                        'markup':f'\n\n### Source block {block_count} — {escape(part)}\n\n'})
                account(db, record)
                for target, stream in [('candidate',candidate), ('supplementary',supplement)]:
                    payload = render(record, target).encode('utf-8')
                    record[target+'_range'] = [stream.tell(), stream.tell()+len(payload)]
                    stream.write(payload)
                    text_counts[target] += sum(len(f['text']) for f in record['fragments']
                                               if f['target'] == target and 'id' in f)
                evidence.write(json.dumps(record, ensure_ascii=False)+'\n')
                issue_counts.update(record['issues'])
                if block_count % 200 == 0:
                    LOG.info('Processed %d logical blocks', block_count)
            db.commit()
        missing = db.execute('SELECT count(*) FROM atoms WHERE used=0').fetchone()[0]
        if missing:
            raise ValueError(f'{missing} source atoms were not accounted for')
        if next(expected_blocks,None) is not None:
            raise ValueError('Missing structural source blocks')
    if not atom_count:
        issue_counts['no_recognized_text_requires_review'] += 1
    build_review(output)
    for filename in ['source.docx','candidate.md','supplementary.md','evidence.jsonl.gz','review.html']:
        hashes[filename] = digest(output/filename)
    blocking = {k:v for k,v in issue_counts.items() if k in BLOCKING_ISSUES}
    report = {'schema_version':2, 'format':'docx', 'status':'blocked' if blocking else 'review_required', 'blocks':block_count,
              'blocking_issues':blocking,
              'source_atoms':atom_count, 'source_characters_by_destination':dict(text_counts),
              'issues':dict(issue_counts), 'hashes':hashes,
              'limitations':'Text accounting is not proof of visual fidelity. No Word pagination is inferred.'}
    write_json(output/'report.json', report)
    write_json(output/'review-template.json', {'hashes':hashes, 'reviewer':'',
               'blocks':[{'block':i,'approved':False,'note':''} for i in range(1,block_count+1)]})
    (output/'RUNNING.json').unlink()
    dbpath.unlink()
    LOG.info('Converted %d blocks; %s: %s', block_count, report['status'], output)
    return 2 if blocking else 3


def verify(output: Path) -> dict[str, Any]:
    """Independently re-inventory source XML and check every serialized fragment."""
    if (output/'RUNNING.json').exists():
        raise ValueError('Incomplete conversion')
    report = json.loads((output/'report.json').read_text(encoding='utf-8'))
    for name, expected in report['hashes'].items():
        path = (output/name).resolve()
        if not path.is_relative_to(output.resolve()) or digest(path) != expected:
            raise ValueError('Artifact hash/path mismatch: '+name)
    import tempfile
    with tempfile.TemporaryDirectory() as temporary, closing(sqlite3.connect(Path(temporary)/'audit.db')) as db, \
            ZipFile(output/'source.docx') as z, gzip.open(output/'evidence.jsonl.gz','rt',encoding='utf-8') as evidence, \
            (output/'candidate.md').open('rb') as candidate, (output/'supplementary.md').open('rb') as supplement:
        validate_package(z, Limits())
        total = inventory(z, db)
        expected_blocks = structural_expectations(z)
        count = 0
        for line in evidence:
            record = json.loads(line)
            count += 1
            if record['block'] != count:
                raise ValueError('Nonsequential block identity')
            account(db, record)
            expected = next(expected_blocks,None)
            if expected is None:
                raise ValueError('Extra output block')
            audit_structure(record,expected)
            for target, stream in [('candidate',candidate),('supplementary',supplement)]:
                payload = render(record,target).encode('utf-8')
                if record[target+'_range'] != [stream.tell(),stream.tell()+len(payload)] or stream.read(len(payload)) != payload:
                    raise ValueError('Serialized Markdown differs from evidence')
        if candidate.read(1) or supplement.read(1) or count != report['blocks'] or total != report['source_atoms']:
            raise ValueError('Trailing content or counts mismatch')
        if db.execute('SELECT count(*) FROM atoms WHERE used=0').fetchone()[0]:
            raise ValueError('Unaccounted source atoms')
        if next(expected_blocks,None) is not None:
            raise ValueError('Missing structural source blocks')
    return report


def release(output: Path, review_path: Path) -> None:
    report = verify(output)
    if report.get('blocking_issues'):
        raise ValueError('Unresolved structural content blocks release; repair the source or converter first')
    review = json.loads(review_path.read_text(encoding='utf-8'))
    if review.get('hashes') != report['hashes'] or not str(review.get('reviewer','')).strip():
        raise ValueError('Review identity/hash mismatch')
    decisions = review.get('blocks',[])
    if [d.get('block') for d in decisions] != list(range(1,report['blocks']+1)):
        raise ValueError('All blocks must be reviewed exactly once')
    if any(d.get('approved') is not True or not str(d.get('note','')).strip() for d in decisions):
        raise ValueError('Every block needs approval and a note')
    with (output/'accepted-review.json').open('x',encoding='utf-8') as stream:
        json.dump(review,stream,ensure_ascii=False,indent=2)
    with (output/'ready.md.pending').open('xb') as dst, (output/'candidate.md').open('rb') as src:
        shutil.copyfileobj(src,dst)
    if (output/'ready.md').exists():
        raise FileExistsError('ready.md already exists')
    (output/'ready.md.pending').rename(output/'ready.md')


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest='command',required=True)
    c = sub.add_parser('convert')
    c.add_argument('source',type=Path)
    c.add_argument('output',type=Path)
    v = sub.add_parser('verify')
    v.add_argument('output',type=Path)
    r = sub.add_parser('release')
    r.add_argument('output',type=Path)
    r.add_argument('--review',type=Path,required=True)
    args = parser.parse_args(argv)
    logging.basicConfig(level=logging.INFO,format='%(levelname)s %(message)s')
    try:
        if args.command == 'convert':
            return convert(args.source,args.output)
        if args.command == 'verify':
            report = verify(args.output)
            LOG.info('Source accounting verified; status: %s',report['status'])
        else:
            release(args.output,args.review)
        return 0
    except (Exception,KeyboardInterrupt) as error:
        LOG.error('%s: %s',type(error).__name__,error)
        return 1


if __name__ == '__main__':
    sys.exit(main())
