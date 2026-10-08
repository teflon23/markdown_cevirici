"""Synthetic OOXML regression fixtures. No external document or network needed."""
from __future__ import annotations

import gzip
import json
import hashlib
from pathlib import Path
import tempfile
import unittest
from zipfile import ZipFile, ZIP_DEFLATED

from docx_to_markdown import W, A, M, R, Limits, convert, verify, release, number


def paragraph(value: str, properties: str = '') -> str:
    return f'<w:p>{properties}<w:r><w:t xml:space="preserve">{value}</w:t></w:r></w:p>'


def package(path: Path, body: str, extras: dict[str, str | bytes] | None = None) -> Path:
    document = (f'<w:document xmlns:w="{W[1:-1]}" xmlns:a="{A[1:-1]}" xmlns:m="{M[1:-1]}" '
                f'xmlns:r="{R[1:-1]}"><w:body>{body}</w:body></w:document>')
    with ZipFile(path, 'w', compression=ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr('word/document.xml', document)
        for name, value in (extras or {}).items():
            z.writestr(name,value)
    return path


class DocxTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.out = self.root/'out'

    def tearDown(self) -> None:
        self.temp.cleanup()

    def run_doc(self, body: str, extras: dict[str, str | bytes] | None = None) -> tuple[str, str, dict]:
        source = package(self.root/'source.docx',body,extras)
        status = convert(source,self.out)
        report = verify(self.out)
        self.assertEqual(status,2 if report['blocking_issues'] else 3)
        return ((self.out/'candidate.md').read_text('utf-8'),
                (self.out/'supplementary.md').read_text('utf-8'),report)

    def records(self) -> list[dict]:
        with gzip.open(self.out/'evidence.jsonl.gz','rt',encoding='utf-8') as stream:
            return [json.loads(line) for line in stream]

    def test_split_runs_do_not_insert_spaces(self) -> None:
        md,_,r = self.run_doc('<w:p><w:r><w:t>comp</w:t></w:r><w:r><w:t>liance</w:t></w:r></w:p>')
        self.assertEqual(md,'compliance\n\n')
        self.assertEqual(r['source_atoms'],2)

    def test_literal_markup_and_unicode_are_preserved(self) -> None:
        value = '  Türkçe &amp; &lt;tag&gt; | [a] **b**  '
        md,_,r = self.run_doc(paragraph(value))
        self.assertIn('Türkçe',md)
        self.assertIn(r'\|',md)
        self.assertNotIn('<tag>',md)
        self.assertEqual(r['source_atoms'],1)

    def test_inherited_style_heading(self) -> None:
        styles = f'<w:styles xmlns:w="{W[1:-1]}"><w:style w:styleId="Base"><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style><w:style w:styleId="Legal"><w:basedOn w:val="Base"/></w:style></w:styles>'
        md,_,_ = self.run_doc(paragraph('Article 25.2','<w:pPr><w:pStyle w:val="Legal"/></w:pPr>'),{'word/styles.xml':styles})
        self.assertEqual(md,'## Article 25.2\n\n')

    def test_body_table_body_order(self) -> None:
        table = '<w:tbl><w:tr><w:tc>'+paragraph('A|B')+paragraph('C')+'</w:tc><w:tc>'+paragraph('D')+'</w:tc></w:tr></w:tbl>'
        md,_,_ = self.run_doc(paragraph('Before')+table+paragraph('After'))
        self.assertLess(md.index('Before'),md.index('| Column 1'))
        self.assertLess(md.index('D<br>'),md.index('After'))
        self.assertIn(r'A\|B<br>C<br>',md)
        separator = next(line for line in md.splitlines() if line.startswith('| ---'))
        self.assertEqual(len(separator.split('|')),4)

    def test_merged_cells_keep_text_once(self) -> None:
        table = '<w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr>'+paragraph('Merged')+'</w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr>'+paragraph('')+'</w:tc></w:tr></w:tbl>'
        md,_,report = self.run_doc(table)
        self.assertEqual(md.count('Merged'),1)
        self.assertIn('merged_table_cells',report['issues'])
        self.assertEqual(self.records()[0]['metadata']['table']['columns'],2)

    def test_nested_table_text_is_accounted(self) -> None:
        table = '<w:tbl><w:tr><w:tc>'+paragraph('Outer')+'<w:tbl><w:tr><w:tc>'+paragraph('Inner')+'</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>'
        md,_,r = self.run_doc(table)
        self.assertIn('Inner',md)
        self.assertIn('nested_table_flattened_requires_review',r['issues'])

    def test_tracked_deletion_is_separated(self) -> None:
        md,sup,r = self.run_doc('<w:p><w:del><w:r><w:delText>old</w:delText></w:r></w:del><w:ins><w:r><w:t>new</w:t></w:r></w:ins></w:p>')
        self.assertNotIn('old',md)
        self.assertIn('new',md)
        self.assertIn('old',sup)
        self.assertIn('tracked_changes_require_review',r['issues'])

    def test_paragraph_inside_deleted_wrapper(self) -> None:
        md,sup,_ = self.run_doc('<w:del>'+paragraph('deleted')+'</w:del>'+paragraph('current'))
        self.assertNotIn('deleted',md)
        self.assertIn('deleted',sup)

    def test_header_footer_and_comments_retained(self) -> None:
        extras = {'word/header1.xml':f'<w:hdr xmlns:w="{W[1:-1]}">{paragraph("Header")}</w:hdr>',
                  'word/comments.xml':f'<w:comments xmlns:w="{W[1:-1]}"><w:comment w:id="0">{paragraph("Comment")}</w:comment></w:comments>'}
        md,sup,r = self.run_doc(paragraph('Body'),extras)
        self.assertNotIn('Header',md)
        self.assertIn('Header',sup)
        self.assertIn('Comment',sup)
        self.assertEqual(r['source_atoms'],3)

    def test_footnote_reference_and_definition(self) -> None:
        extras = {'word/footnotes.xml':f'<w:footnotes xmlns:w="{W[1:-1]}"><w:footnote w:id="1">{paragraph("Note text")}</w:footnote></w:footnotes>'}
        md,_,_ = self.run_doc('<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="1"/></w:r></w:p>',extras)
        self.assertIn('](#footnote-1)',md)
        self.assertIn('## footnote 1',md)
        self.assertIn('Note text',md)

    def test_hyperlink_and_fields(self) -> None:
        rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="https://example.com/a" TargetMode="External"/></Relationships>'
        body = '<w:p><w:hyperlink r:id="r1"><w:r><w:t>Link</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> PAGE </w:instrText><w:t>7</w:t></w:r></w:p>'
        md,sup,r = self.run_doc(body,{'word/_rels/document.xml.rels':rels})
        self.assertIn('[Link](https://example.com/a)',md)
        self.assertNotIn(' PAGE ',md)
        self.assertIn(' PAGE ',sup)
        self.assertIn('field_cached_result_requires_review',r['issues'])

    def test_image_is_local_and_hashed(self) -> None:
        rels = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Target="media/image.png"/></Relationships>'
        md,_,r = self.run_doc('<w:p><w:r><w:drawing><a:blip r:embed="r1"/></w:drawing></w:r></w:p>',{'word/_rels/document.xml.rels':rels,'word/media/image.png':b'fixture'})
        self.assertIn('![Image](assets/',md)
        self.assertEqual(len(list((self.out/'assets').iterdir())),1)
        self.assertIn('image_requires_visual_review',r['issues'])
        asset = next((self.out/'assets').iterdir())
        asset.write_bytes(b'changed')
        with self.assertRaisesRegex(ValueError,'hash/path mismatch'):
            verify(self.out)

    def test_hidden_text_is_retained_separately(self) -> None:
        md,sup,r = self.run_doc('<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r><w:r><w:t>visible</w:t></w:r></w:p>')
        self.assertNotIn('hidden',md)
        self.assertIn('hidden',sup)
        self.assertIn('hidden_text_requires_review',r['issues'])

    def test_review_html_escapes_source(self) -> None:
        self.run_doc(paragraph('&lt;script&gt;alert(1)&lt;/script&gt;'))
        review = (self.out/'review.html').read_text('utf-8')
        self.assertNotIn('<script>',review)
        self.assertIn('&lt;script&gt;',review)

    def test_strict_ooxml_fails_explicitly(self) -> None:
        source = package(self.root/'strict.docx','')
        with ZipFile(source,'w') as z:
            z.writestr('[Content_Types].xml','<Types/>')
            z.writestr('word/document.xml','<document xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"/>')
        with self.assertRaisesRegex(ValueError,'namespace'):
            convert(source,self.out)

    def test_numbering_nested_and_restart(self) -> None:
        numbering = f'<w:numbering xmlns:w="{W[1:-1]}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="(%2)"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
        body = ''.join(paragraph('Item',f'<w:pPr><w:numPr><w:ilvl w:val="{level}"/><w:numId w:val="1"/></w:numPr></w:pPr>') for level in [0,1,1,0,1])
        md,_,_ = self.run_doc(body,{'word/numbering.xml':numbering})
        self.assertIn('- **1.** Item',md)
        self.assertIn('    - **(b)** Item',md)
        self.assertEqual(md.count('**(a)**'),2)

    def test_math_and_unknown_text_not_lost(self) -> None:
        md,sup,r = self.run_doc('<w:p><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></w:p>',{'word/charts/chart1.xml':f'<a:root xmlns:a="{A[1:-1]}"><a:t>Chart title</a:t></a:root>'})
        self.assertNotIn('x',md)
        self.assertIn('x',sup)
        self.assertIn('Chart title',sup)
        self.assertEqual(r['source_atoms'],2)

    def test_doctype_is_rejected(self) -> None:
        source = package(self.root/'bad.docx','',{'evil.xml':'<!DOCTYPE x [<!ENTITY e "payload">]><x>&e;</x>'})
        with self.assertRaisesRegex(ValueError,'DTD'):
            convert(source,self.out)
        self.assertFalse(self.out.exists())

    def test_unsafe_zip_path_is_rejected(self) -> None:
        source = package(self.root/'bad.docx','',{'../outside':'bad'})
        with self.assertRaisesRegex(ValueError,'Unsafe'):
            convert(source,self.out)

    def test_expansion_limit(self) -> None:
        source = package(self.root/'large.docx',paragraph('too large'))
        with self.assertRaisesRegex(ValueError,'size limit'):
            convert(source,self.out,Limits(total_bytes=10))

    def test_output_tampering_rejected(self) -> None:
        self.run_doc(paragraph('must'))
        (self.out/'candidate.md').write_text('may',encoding='utf-8')
        with self.assertRaisesRegex(ValueError,'hash/path mismatch'):
            verify(self.out)

    def replace_evidence(self, records: list[dict]) -> None:
        """Refresh file hashes to test inventory checks independently of hashing."""
        from docx_to_markdown import render
        offsets = {'candidate':0,'supplementary':0}
        buffers = {'candidate':b'','supplementary':b''}
        for record in records:
            for target in offsets:
                data = render(record,target).encode('utf-8')
                record[target+'_range'] = [offsets[target],offsets[target]+len(data)]
                offsets[target] += len(data)
                buffers[target] += data
        for target,data in buffers.items():
            (self.out/(target+'.md')).write_bytes(data)
        with gzip.open(self.out/'evidence.jsonl.gz','wt',encoding='utf-8') as stream:
            for record in records:
                stream.write(json.dumps(record)+'\n')
        report = json.loads((self.out/'report.json').read_text('utf-8'))
        for name in ['candidate.md','supplementary.md','evidence.jsonl.gz']:
            report['hashes'][name] = hashlib.sha256((self.out/name).read_bytes()).hexdigest()
        (self.out/'report.json').write_text(json.dumps(report),encoding='utf-8')

    def test_reordered_source_is_rejected_even_with_updated_hashes(self) -> None:
        self.run_doc('<w:p><w:r><w:t>shall </w:t><w:t>not</w:t></w:r></w:p>')
        records = self.records()
        records[0]['fragments'][0],records[0]['fragments'][1] = records[0]['fragments'][1],records[0]['fragments'][0]
        self.replace_evidence(records)
        with self.assertRaisesRegex(ValueError,'order changed'):
            verify(self.out)

    def test_missing_source_is_rejected_even_with_updated_hashes(self) -> None:
        self.run_doc(paragraph('obligation'))
        records = self.records()
        records[0]['fragments'] = [f for f in records[0]['fragments'] if 'id' not in f]
        self.replace_evidence(records)
        with self.assertRaisesRegex(ValueError,'Unaccounted'):
            verify(self.out)

    def test_top_level_altchunk_is_flagged(self) -> None:
        _,_,r = self.run_doc('<w:altChunk r:id="r99"/>')
        self.assertIn('unsupported_embedded_content',r['issues'])

    def test_no_output_overwrite(self) -> None:
        self.run_doc(paragraph('body'))
        with self.assertRaises(FileExistsError):
            convert(self.root/'source.docx',self.out)

    def test_release_gate(self) -> None:
        self.run_doc(paragraph('body'))
        path = self.out/'review-template.json'
        with self.assertRaises(ValueError):
            release(self.out,path)
        review = json.loads(path.read_text('utf-8'))
        review['reviewer'] = 'Synthetic test reviewer'
        for block in review['blocks']:
            block.update(approved=True,note='Compared against source')
        path.write_text(json.dumps(review),encoding='utf-8')
        release(self.out,path)
        self.assertEqual((self.out/'ready.md').read_bytes(),(self.out/'candidate.md').read_bytes())

    def test_number_formats(self) -> None:
        issues: set[str] = set()
        self.assertEqual(number(27,'upperLetter',issues),'AA')
        self.assertEqual(number(49,'lowerRoman',issues),'xlix')
        self.assertFalse(issues)

    def numbered_fixture(self) -> tuple[str, dict[str,str]]:
        props = '<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>'
        xml = f'<w:numbering xmlns:w="{W[1:-1]}"><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
        return props, {'word/numbering.xml':xml}

    def test_table_numbering_continues_into_body(self) -> None:
        props,extras = self.numbered_fixture()
        table = '<w:tbl><w:tr><w:tc>'+paragraph('First',props)+'</w:tc></w:tr></w:tbl>'
        md,_,_ = self.run_doc(table+paragraph('Second',props),extras)
        self.assertIn('1. First',md)
        self.assertIn('**2.** Second',md)

    def test_generated_number_deletion_fails_even_with_new_hashes(self) -> None:
        props,extras = self.numbered_fixture()
        self.run_doc(paragraph('Item',props),extras)
        records = self.records()
        records[0]['fragments'] = [f for f in records[0]['fragments'] if f.get('semantic')!='numbering']
        self.replace_evidence(records)
        with self.assertRaisesRegex(ValueError,'numbering marker missing'):
            verify(self.out)

    def test_zero_note_id_is_not_separator(self) -> None:
        extras = {'word/footnotes.xml':f'<w:footnotes xmlns:w="{W[1:-1]}"><w:footnote w:id="0">{paragraph("Ordinary note")}</w:footnote><w:footnote w:id="7" w:type="separator">{paragraph("Separator")}</w:footnote></w:footnotes>'}
        md,sup,_ = self.run_doc(paragraph('Body'),extras)
        self.assertIn('Ordinary note',md)
        self.assertNotIn('Separator',md)
        self.assertIn('Separator',sup)

    def test_missing_note_blocks_release(self) -> None:
        _,_,report = self.run_doc('<w:p><w:r><w:t>Body</w:t><w:footnoteReference w:id="9"/></w:r></w:p>')
        self.assertEqual(report['status'],'blocked')
        with self.assertRaisesRegex(ValueError,'structural content blocks'):
            release(self.out,self.out/'review-template.json')

    def test_bold_runs_join_and_spaces_stay_outside(self) -> None:
        body = '<w:p>'+''.join(f'<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{s}</w:t></w:r>' for s in [' Alpha',' ','Beta '])+'</w:p>'
        md,_,_ = self.run_doc(body)
        self.assertEqual(md,' **Alpha Beta** \n\n')

    def test_turkish_captioned_article_and_chapter(self) -> None:
        md,_,_ = self.run_doc(paragraph('BİRİNCİ BÖLÜM')+paragraph('Tanımlar MADDE 4- (1) İçerik'))
        self.assertIn('# BİRİNCİ BÖLÜM',md)
        self.assertIn('## MADDE 4\n',md)
        self.assertIn('Tanımlar MADDE 4- (1) İçerik',md)

    def test_layout_prose_row_unwraps(self) -> None:
        row = '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr>'+paragraph('BİRİNCİ BÖLÜM')+paragraph('MADDE 1- (1) İçerik')+paragraph('Devam')+'</w:tc></w:tr>'
        md,_,report = self.run_doc('<w:tbl>'+row+'</w:tbl>')
        self.assertIn('\n## MADDE 1\n',md)
        self.assertNotIn('| Column',md)
        self.assertIn('full_width_prose_row_unwrapped_requires_review',report['issues'])

    def test_table_geometry_tamper_fails(self) -> None:
        self.run_doc('<w:tbl><w:tr><w:tc>'+paragraph('Cell')+'</w:tc></w:tr></w:tbl>')
        records = self.records()
        records[0]['metadata']['table']['columns'] = 2
        self.replace_evidence(records)
        with self.assertRaisesRegex(ValueError,'table grid mismatch'):
            verify(self.out)

    def test_heading_deletion_fails(self) -> None:
        self.run_doc(paragraph('MADDE 1- (1) İçerik'))
        records = self.records()
        records[0]['fragments'] = [f for f in records[0]['fragments'] if f.get('semantic')!='legal_heading']
        self.replace_evidence(records)
        with self.assertRaisesRegex(ValueError,'legal heading missing'):
            verify(self.out)


if __name__ == '__main__':
    unittest.main()
