"""Patch Table_Production.xlsm for the table QC flow — zip-level, no Excel, no openpyxl.

The workbook carries form-control buttons, charts, a Power Pivot model, slicers
and a VBA project. openpyxl drops most of that on a round-trip, so the package is
edited at the XML level and re-zipped with every other part byte-identical.

What it changes (and nothing else):

TABLE Floor (xl/worksheets/sheet2.xml)
  * R3 = "QC Status", S3 = "QC Date"  (headers for the columns the macro writes)
  * conditional formatting A4:Q500 — grey fill + grey text when QC Status is set,
    plus strikethrough when it is REJECTED. Greying is data-driven, so clearing
    the status un-greys the row; the macro never paints cells directly.
  * the 16 stacked "RELEASE CARTS" buttons now call TableQC_Process and read
    "QC / RELEASE TABLE". They used to call ReleaseToShippingAndReceiving, the
    REVO shaft flow, which reads a "REVO Floor" sheet this workbook does not have.
  * column widths for R:S

xl/styles.xml
  * two new differential formats (dxf) for the rules above

Usage:
    python table_production/tools/patch_workbook.py <in.xlsm> <out.xlsm>
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

SHEET_PART = "xl/worksheets/sheet2.xml"          # TABLE Floor (codename Sheet7)
STYLES_PART = "xl/styles.xml"
VML_PART = "xl/drawings/vmlDrawing1.vml"
DRAWING_PART = "xl/drawings/drawing1.xml"

OLD_MACRO = "[0]!ReleaseToShippingAndReceiving"
NEW_MACRO = "[0]!TableQC_Process"
OLD_CAPTION = "RELEASE CARTS"
NEW_CAPTION = "QC / RELEASE TABLE"

HEADER_ROW = 3
FIRST_DATA_ROW = 4
LAST_CF_ROW = 500                                  # matches the Holds!$A$2:$A$500 horizon
HEADER_STYLE = "363"                               # same xf as the O3:Q3 headers

GREY_FILL = "FFD9D9D9"
GREY_FONT = "FF808080"


class PatchError(RuntimeError):
    pass


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PatchError(message)


def patch_styles(xml: str) -> tuple[str, int, int]:
    """Append two dxfs; return (xml, grey_index, reject_index)."""
    match = re.search(r'<dxfs count="(\d+)">', xml)
    _expect(match is not None, "styles.xml: no <dxfs count=...> element")
    count = int(match.group(1))
    grey = (
        f'<dxf><font><color rgb="{GREY_FONT}"/></font>'
        f'<fill><patternFill patternType="solid"><fgColor indexed="64"/>'
        f'<bgColor rgb="{GREY_FILL}"/></patternFill></fill></dxf>'
    )
    reject = (
        f'<dxf><font><strike/><color rgb="{GREY_FONT}"/></font>'
        f'<fill><patternFill patternType="solid"><fgColor indexed="64"/>'
        f'<bgColor rgb="{GREY_FILL}"/></patternFill></fill></dxf>'
    )
    xml = xml.replace(match.group(0), f'<dxfs count="{count + 2}">', 1)
    xml = xml.replace("</dxfs>", grey + reject + "</dxfs>", 1)
    return xml, count, count + 1


def patch_sheet(xml: str, grey_dxf: int, reject_dxf: int) -> str:
    # --- headers R3 / S3 ---------------------------------------------------
    row3 = re.search(rf'<row r="{HEADER_ROW}"[^>]*>.*?</row>', xml, re.S)
    _expect(row3 is not None, f"sheet: row {HEADER_ROW} not found")
    row_xml = row3.group(0)
    _expect(f'<c r="R{HEADER_ROW}" s="10"/>' in row_xml, "sheet: R3 is not the expected empty cell")
    _expect(f'<c r="S{HEADER_ROW}" s="10"/>' in row_xml, "sheet: S3 is not the expected empty cell")
    new_row = row_xml.replace(
        f'<c r="R{HEADER_ROW}" s="10"/>',
        f'<c r="R{HEADER_ROW}" s="{HEADER_STYLE}" t="inlineStr"><is><t>QC Status</t></is></c>',
    ).replace(
        f'<c r="S{HEADER_ROW}" s="10"/>',
        f'<c r="S{HEADER_ROW}" s="{HEADER_STYLE}" t="inlineStr"><is><t>QC Date</t></is></c>',
    )
    xml = xml.replace(row_xml, new_row, 1)

    # --- column widths for R:S --------------------------------------------
    tail_col = '<col min="18" max="16384" width="9.42578125"'
    _expect(tail_col in xml, "sheet: <cols> tail entry not as expected")
    cols_match = re.search(r'<col min="18" max="16384" width="9.42578125"[^>]*/>', xml)
    assert cols_match is not None
    attrs = cols_match.group(0)[len('<col min="18" max="16384" width="9.42578125"'):]
    xml = xml.replace(
        cols_match.group(0),
        '<col min="18" max="18" width="14.7109375" customWidth="1"/>'
        '<col min="19" max="19" width="11.7109375" customWidth="1"/>'
        '<col min="20" max="16384" width="9.42578125"' + attrs,
        1,
    )

    # --- conditional formatting ------------------------------------------
    # Schema order: ... phoneticPr, conditionalFormatting, dataValidations,
    # hyperlinks, printOptions ... so it goes right before <printOptions>.
    _expect("<conditionalFormatting" not in xml, "sheet: already has conditional formatting; patched twice?")
    _expect("<printOptions" in xml, "sheet: no <printOptions> anchor")
    sqref = f"A{FIRST_DATA_ROW}:Q{LAST_CF_ROW}"
    cf = (
        f'<conditionalFormatting sqref="{sqref}">'
        f'<cfRule type="expression" dxfId="{reject_dxf}" priority="1" stopIfTrue="1">'
        f'<formula>$R{FIRST_DATA_ROW}="REJECTED"</formula></cfRule>'
        f'<cfRule type="expression" dxfId="{grey_dxf}" priority="2">'
        f'<formula>$R{FIRST_DATA_ROW}&lt;&gt;""</formula></cfRule>'
        f'</conditionalFormatting>'
    )
    xml = xml.replace("<printOptions", cf + "<printOptions", 1)

    # --- buttons -> new macro ----------------------------------------------
    n = xml.count(f'macro="{OLD_MACRO}"')
    _expect(n > 0, "sheet: no buttons bound to the old macro")
    xml = xml.replace(f'macro="{OLD_MACRO}"', f'macro="{NEW_MACRO}"')
    return xml


def patch_caption(xml: str) -> str:
    _expect(OLD_CAPTION in xml, f"caption '{OLD_CAPTION}' not found")
    return xml.replace(OLD_CAPTION, NEW_CAPTION)


def _well_formed(name: str, xml: str) -> None:
    try:
        ET.fromstring(xml.encode("utf-8"))
    except ET.ParseError as exc:
        raise PatchError(f"{name}: patched XML is not well-formed ({exc})") from exc


def patch(src: Path, dst: Path) -> dict[str, int]:
    with zipfile.ZipFile(src, "r") as zin:
        names = zin.namelist()
        for part in (SHEET_PART, STYLES_PART, VML_PART, DRAWING_PART):
            _expect(part in names, f"{part} missing from package")
        styles, grey, reject = patch_styles(zin.read(STYLES_PART).decode("utf-8"))
        sheet = patch_sheet(zin.read(SHEET_PART).decode("utf-8"), grey, reject)
        vml = patch_caption(zin.read(VML_PART).decode("latin-1"))
        drawing = patch_caption(zin.read(DRAWING_PART).decode("utf-8"))

        _well_formed(STYLES_PART, styles)
        _well_formed(SHEET_PART, sheet)
        _well_formed(DRAWING_PART, drawing)

        replaced = {
            STYLES_PART: styles.encode("utf-8"),
            SHEET_PART: sheet.encode("utf-8"),
            VML_PART: vml.encode("latin-1"),
            DRAWING_PART: drawing.encode("utf-8"),
        }
        dst.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(dst, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():                       # same order, same names
                payload = replaced.get(info.filename, zin.read(info.filename))
                zout.writestr(info, payload)
    return {
        "buttons_repointed": sheet.count(f'macro="{NEW_MACRO}"'),
        "grey_dxf": grey,
        "reject_dxf": reject,
    }


if __name__ == "__main__":
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    result = patch(Path(sys.argv[1]), Path(sys.argv[2]))
    print(f"patched -> {sys.argv[2]}: {result}")
