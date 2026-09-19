#!/usr/bin/env python3
"""Generate deterministic mixed-type XLSX files for performance testing."""

from __future__ import annotations

import argparse
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def static_parts(shared: bool) -> dict[str, str]:
    shared_override = (
        '<Override PartName="/xl/sharedStrings.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        if shared
        else ""
    )
    shared_rel = (
        f'<Relationship Id="rId3" Type="{REL}/sharedStrings" Target="sharedStrings.xml"/>'
        if shared
        else ""
    )
    return {
        "[Content_Types].xml": f'''<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{shared_override}</Types>''',
        "_rels/.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<workbook xmlns="{MAIN}" xmlns:r="{REL}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{PKG}"><Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/>{shared_rel}</Relationships>',
        "xl/styles.xml": f'<styleSheet xmlns="{MAIN}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>',
    }


def inline_cell(reference: str, value: str) -> str:
    return f'<c r="{reference}" t="inlineStr"><is><t>{value}</t></is></c>'


def shared_cell(reference: str, index: int) -> str:
    return f'<c r="{reference}" t="s"><v>{index}</v></c>'


def generate(output: Path, rows: int, mode: str) -> None:
    shared = mode == "shared"
    output.parent.mkdir(parents=True, exist_ok=True)
    headers = ("text", "int", "float", "bool", "date", "text2", "int2", "float2")
    pool = list(headers) + [f"label-{index}" for index in range(256)]
    with ZipFile(output, "w", ZIP_DEFLATED, compresslevel=6) as archive:
        for name, content in static_parts(shared).items():
            archive.writestr(name, content.encode())
        if shared:
            items = "".join(f"<si><t>{value}</t></si>" for value in pool)
            archive.writestr(
                "xl/sharedStrings.xml",
                f'<sst xmlns="{MAIN}" count="{rows * 2 + 8}" uniqueCount="{len(pool)}">{items}</sst>'.encode(),
            )
        with archive.open("xl/worksheets/sheet1.xml", "w") as sheet:
            sheet.write(f'<worksheet xmlns="{MAIN}"><sheetData>'.encode())
            header_cells = "".join(
                shared_cell(f"{chr(65 + index)}1", index)
                if shared
                else inline_cell(f"{chr(65 + index)}1", value)
                for index, value in enumerate(headers)
            )
            sheet.write(f'<row r="1">{header_cells}</row>'.encode())
            for row in range(2, rows + 2):
                first = shared_cell(f"A{row}", 8 + row % 256) if shared else inline_cell(f"A{row}", f"label-{row % 256}")
                second = shared_cell(f"F{row}", 8 + (row * 7) % 256) if shared else inline_cell(f"F{row}", f"label-{(row * 7) % 256}")
                xml = (
                    f'<row r="{row}">{first}<c r="B{row}"><v>{row}</v></c>'
                    f'<c r="C{row}"><v>{row / 7:.7f}</v></c><c r="D{row}" t="b"><v>{row & 1}</v></c>'
                    f'<c r="E{row}" s="1"><v>{45000 + row % 1000}</v></c>{second}'
                    f'<c r="G{row}"><v>{row * 3}</v></c><c r="H{row}"><v>{row / 13:.7f}</v></c></row>'
                )
                sheet.write(xml.encode())
            sheet.write(b"</sheetData></worksheet>")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    parser.add_argument("--rows", type=int, default=60_000)
    parser.add_argument("--mode", choices=("inline", "shared"), default="inline")
    args = parser.parse_args()
    if args.rows < 1:
        parser.error("--rows must be positive")
    generate(args.output, args.rows, args.mode)


if __name__ == "__main__":
    main()
