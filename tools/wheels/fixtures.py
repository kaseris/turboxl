"""Portable, dependency-free ZIP/XLSX fixtures for native and installed-wheel tests."""
from pathlib import Path
import argparse
import zipfile


def archive(root, output, members):
    root, output = Path(root), Path(output)
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for member in members:
            path = root / member
            if not path.exists():
                raise FileNotFoundError(path)
            files = sorted(path.rglob('*')) if path.is_dir() else [path]
            for file in files:
                if file.is_file() and file.resolve() != output.resolve():
                    z.write(file, file.relative_to(root).as_posix())


def workbook(output, *, date1904=False, scalars=False):
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
    workbook_properties = '<workbookPr date1904="1"/>' if date1904 else ''
    data_sheet = (
        f'<worksheet xmlns="{main}"><sheetData><row r="1">'
        '<c r="B1" t="b"><v>1</v></c><c r="C1"><v>42</v></c>'
        '<c r="D1"><v>42.5</v></c><c r="E1" t="inlineStr"><is><t>text</t></is></c>'
        '<c r="F1" t="e"><v>#N/A</v></c></row><row r="2">'
        '<c r="A2" s="1"><v>59</v></c><c r="B2" s="1"><v>60</v></c>'
        '<c r="C2" s="1"><v>61</v></c>'
        '<c r="D2" s="2"><v>45292.123456789</v></c>'
        '<c r="E2" s="3"><v>0.999999999999</v></c>'
        '<c r="F2" s="1"><v>4000000</v></c>'
        '<c r="G2"><f>20+22</f><v>42</v></c>'
        '</row></sheetData></worksheet>'
        if scalars else
        f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c><c r="C1" s="1"><v>45306</v></c><c r="D1" s="2"><v>45306.57326388889</v></c></row></sheetData></worksheet>'
    )
    parts = {
        '[Content_Types].xml': '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
<Override PartName="/xl/dialogsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.dialogsheet+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet4.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>''',
        '_rels/.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{main}" xmlns:r="{rel}">{workbook_properties}<sheets><sheet name="Chart" sheetId="1" r:id="rId1"/><sheet name="Data" sheetId="2" r:id="rId2"/><sheet name="Hidden" sheetId="3" state="hidden" r:id="rId3"/><sheet name="VeryHidden" sheetId="4" state="veryHidden" r:id="rId4"/><sheet name="Dialog" sheetId="5" r:id="rId5"/><sheet name="Sparse" sheetId="6" r:id="rId6"/></sheets></workbook>',
        # Exercise both OPC target forms: openpyxl commonly writes package-absolute
        # worksheet targets, while other producers use workbook-relative targets.
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/chartsheet" Target="chartsheets/sheet1.xml"/><Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId3" Type="{rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId4" Type="{rel}/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId5" Type="{rel}/dialogsheet" Target="dialogsheets/sheet1.xml"/><Relationship Id="rId6" Type="{rel}/worksheet" Target="worksheets/sheet4.xml"/><Relationship Id="rId7" Type="{rel}/styles" Target="styles.xml"/><Relationship Id="rId8" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        'xl/styles.xml': f'<styleSheet xmlns="{main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="18"/></cellXfs></styleSheet>',
        'xl/sharedStrings.xml': f'<sst xmlns="{main}" count="1" uniqueCount="1"><si><t>café, "quoted"</t></si></sst>',
        'xl/chartsheets/sheet1.xml': f'<chartsheet xmlns="{main}"/>',
        'xl/dialogsheets/sheet1.xml': f'<dialogsheet xmlns="{main}"/>',
        'xl/worksheets/sheet1.xml': data_sheet,
        'xl/worksheets/sheet2.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>secret</t></is></c></row></sheetData></worksheet>',
        'xl/worksheets/sheet3.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>deep secret</t></is></c></row></sheetData></worksheet>',
        'xl/worksheets/sheet4.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>origin</t></is></c></row><row r="5"><c r="D5" t="inlineStr"><is><t>gap</t></is></c></row><row r="50"><c r="Z50" t="inlineStr"><is><t>far</t></is></c></row></sheetData></worksheet>',
    }
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data.encode('utf-8'))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--workbook', type=Path)
    group.add_argument('--archive', nargs='+')
    args = parser.parse_args()
    if args.workbook:
        workbook(args.workbook)
    else:
        archive(args.archive[0], args.archive[1], args.archive[2:] or ['.'])
