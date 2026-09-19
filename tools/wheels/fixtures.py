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


def workbook(output):
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    pkg = 'http://schemas.openxmlformats.org/package/2006/relationships'
    parts = {
        '[Content_Types].xml': '''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>''',
        '_rels/.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        'xl/workbook.xml': f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/><sheet name="Sparse" sheetId="3" r:id="rId3"/></sheets></workbook>',
        # Exercise both OPC target forms: openpyxl commonly writes package-absolute
        # worksheet targets, while other producers use workbook-relative targets.
        'xl/_rels/workbook.xml.rels': f'<Relationships xmlns="{pkg}"><Relationship Id="rId1" Type="{rel}/worksheet" Target="/xl/worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{rel}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{rel}/worksheet" Target="worksheets/sheet3.xml"/><Relationship Id="rId4" Type="{rel}/styles" Target="styles.xml"/><Relationship Id="rId5" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        'xl/styles.xml': f'<styleSheet xmlns="{main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts><cellXfs count="3"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/></cellXfs></styleSheet>',
        'xl/sharedStrings.xml': f'<sst xmlns="{main}" count="1" uniqueCount="1"><si><t>café, "quoted"</t></si></sst>',
        'xl/worksheets/sheet1.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>42</v></c><c r="C1" s="1"><v>45306</v></c><c r="D1" s="2"><v>45306.57326388889</v></c></row></sheetData></worksheet>',
        'xl/worksheets/sheet2.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>secret</t></is></c></row></sheetData></worksheet>',
        'xl/worksheets/sheet3.xml': f'<worksheet xmlns="{main}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>origin</t></is></c></row><row r="5"><c r="D5" t="inlineStr"><is><t>gap</t></is></c></row><row r="50"><c r="Z50" t="inlineStr"><is><t>far</t></is></c></row></sheetData></worksheet>',
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
