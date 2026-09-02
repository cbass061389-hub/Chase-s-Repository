"""Build real Office packages to test the discovery probes against.

These are genuine OOXML/BIFF12 files assembled byte-for-byte — a real base64
DataMashup blob wrapping a real nested zip, real BIFF12 record framing, a real
PROJECT stream — not mocks. The probes are therefore exercised on the same code
paths they will hit on the actual estate.

The one thing these cannot substitute for is the estate itself: layout quirks in
the real workbooks (merged header bands, stray totals rows, renamed columns) can
only be validated by running `python -m sc.cli discover` on the OneDrive tree.
"""

from __future__ import annotations

import base64
import io
import struct
import zipfile
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

CONTENT_TYPES: str = """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>
</Types>"""

REL_NS: str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS: str = "http://schemas.openxmlformats.org/package/2006/relationships"
MAIN_NS: str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


# --------------------------------------------------------------------------
# OOXML (.xlsx / .xlsm)
# --------------------------------------------------------------------------

def _shared_strings(strings: Sequence[str]) -> str:
    items: str = "".join(f"<si><t xml:space=\"preserve\">{s}</t></si>" for s in strings)
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>'
        f'<sst xmlns="{MAIN_NS}" count="{len(strings)}" uniqueCount="{len(strings)}">{items}</sst>'
    )


def _column_letter(index: int) -> str:
    letters: str = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _sheet_xml(
    rows: Sequence[Sequence[object]],
    string_index: Dict[str, int],
    formula_cells: Sequence[Tuple[int, int, str]] = (),
    inline_strings: bool = False,
) -> str:
    """Build a worksheet part. ``rows`` holds str or numeric values; None = blank."""
    formula_map: Dict[Tuple[int, int], str] = {(r, c): f for r, c, f in formula_cells}
    body: List[str] = []
    for row_index, row in enumerate(rows, start=1):
        cells: List[str] = []
        for column_index, value in enumerate(row):
            if value is None:
                continue
            reference: str = f"{_column_letter(column_index)}{row_index}"
            formula: str = formula_map.get((row_index, column_index), "")
            formula_xml: str = f"<f>{formula}</f>" if formula else ""
            if isinstance(value, str):
                if inline_strings:
                    cells.append(f'<c r="{reference}" t="inlineStr">{formula_xml}<is><t>{value}</t></is></c>')
                else:
                    cells.append(f'<c r="{reference}" t="s">{formula_xml}<v>{string_index[value]}</v></c>')
            else:
                cells.append(f'<c r="{reference}">{formula_xml}<v>{value}</v></c>')
        width: int = max((len(r) for r in rows), default=1)
        body.append(f'<row r="{row_index}" spans="1:{width}">{"".join(cells)}</row>')

    last: str = f"{_column_letter(max((len(r) for r in rows), default=1) - 1)}{len(rows)}"
    return (
        f'<?xml version="1.0" encoding="UTF-8"?>'
        f'<worksheet xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
        f'<dimension ref="A1:{last}"/><sheetData>{"".join(body)}</sheetData></worksheet>'
    )


def build_mashup_blob(section_text: str) -> str:
    """A real DataMashup part: base64( header + nested zip( Formulas/Section1.m ) )."""
    inner: io.BytesIO = io.BytesIO()
    with zipfile.ZipFile(inner, "w", zipfile.ZIP_DEFLATED) as package:
        package.writestr("Config/Package.xml", '<?xml version="1.0"?><Package/>')
        package.writestr("Formulas/Section1.m", section_text)
    package_bytes: bytes = inner.getvalue()

    blob: bytes = (
        struct.pack("<II", 0, len(package_bytes))
        + package_bytes
        + struct.pack("<I", 0)                     # empty permissions section
        + struct.pack("<I", 0)                     # empty metadata section
    )
    encoded: str = base64.b64encode(blob).decode("ascii")
    return (
        f'<?xml version="1.0" encoding="utf-8"?>'
        f'<DataMashup xmlns="http://schemas.microsoft.com/DataMashup">{encoded}</DataMashup>'
    )


def build_vba_project(project_name: str, modules: Sequence[str], classes: Sequence[str],
                      protected: bool) -> bytes:
    """A stand-in vbaProject.bin whose PROJECT stream text matches what Excel writes."""
    lines: List[str] = [
        "ID=\"{00000000-0000-0000-0000-000000000000}\"",
        *[f"Module={name}" for name in modules],
        *[f"Class={name}" for name in classes],
        "Document=Sheet1/&H00000000",
        f"Name=\"{project_name}\"",
        "HelpContextID=\"0\"",
    ]
    if protected:
        lines += ["CMG=\"AABBCCDDEE\"", "DPB=\"1122334455\"", "GC=\"9988776655\""]
    header: bytes = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 56   # OLE signature
    return header + ("\r\n".join(lines) + "\r\n").encode("latin-1") + b"\x00" * 64


def write_xlsx(
    path: Path,
    sheets: Sequence[Tuple[str, Sequence[Sequence[object]], str]],
    *,
    mashup_section: str = "",
    vba: bytes | None = None,
    tables: Dict[str, Tuple[str, str, Sequence[str]]] | None = None,
    external_link_target: str = "",
    connections: Sequence[Dict[str, str]] = (),
    defined_names: Sequence[Tuple[str, str]] = (),
    formula_cells: Dict[str, Sequence[Tuple[int, int, str]]] | None = None,
    inline_string_sheets: Sequence[str] = (),
) -> Path:
    """Assemble a valid workbook package. ``sheets`` is ``(name, rows, state)``."""
    path.parent.mkdir(parents=True, exist_ok=True)
    tables = tables or {}
    formula_cells = formula_cells or {}

    all_strings: List[str] = []
    for _name, rows, _state in sheets:
        for row in rows:
            for value in row:
                if isinstance(value, str) and value not in all_strings:
                    all_strings.append(value)
    string_index: Dict[str, int] = {s: i for i, s in enumerate(all_strings)}

    workbook_rels: List[str] = []
    sheet_entries: List[str] = []
    rel_number: int = 0

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr(
            "_rels/.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL_NS}">'
            f'<Relationship Id="rIdWb" Type="{REL_NS}/officeDocument" Target="xl/workbook.xml"/>'
            f'</Relationships>',
        )

        for index, (name, rows, state) in enumerate(sheets, start=1):
            rel_number += 1
            rel_id: str = f"rId{rel_number}"
            part: str = f"xl/worksheets/sheet{index}.xml"
            zf.writestr(
                part,
                _sheet_xml(rows, string_index, formula_cells.get(name, ()), name in inline_string_sheets),
            )
            state_attr: str = f' state="{state}"' if state != "visible" else ""
            sheet_entries.append(
                f'<sheet name="{name}" sheetId="{index}" r:id="{rel_id}"{state_attr}/>'
            )
            workbook_rels.append(
                f'<Relationship Id="{rel_id}" Type="{REL_NS}/worksheet" '
                f'Target="worksheets/sheet{index}.xml"/>'
            )

            if name in tables:
                table_name, reference, columns = tables[name]
                table_part: str = f"xl/tables/table{index}.xml"
                column_xml: str = "".join(
                    f'<tableColumn id="{c + 1}" name="{column}"/>' for c, column in enumerate(columns)
                )
                zf.writestr(
                    table_part,
                    f'<?xml version="1.0"?><table xmlns="{MAIN_NS}" id="{index}" '
                    f'name="{table_name}" displayName="{table_name}" ref="{reference}">'
                    f'<tableColumns count="{len(columns)}">{column_xml}</tableColumns></table>',
                )
                zf.writestr(
                    f"xl/worksheets/_rels/sheet{index}.xml.rels",
                    f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL_NS}">'
                    f'<Relationship Id="rIdT{index}" Type="{REL_NS}/table" '
                    f'Target="../tables/table{index}.xml"/></Relationships>',
                )

        if all_strings:
            rel_number += 1
            zf.writestr("xl/sharedStrings.xml", _shared_strings(all_strings))
            workbook_rels.append(
                f'<Relationship Id="rId{rel_number}" Type="{REL_NS}/sharedStrings" '
                f'Target="sharedStrings.xml"/>'
            )

        if connections:
            rel_number += 1
            connection_xml: str = "".join(
                f'<connection id="{i + 1}" name="{c["name"]}" type="{c.get("type", "5")}" '
                f'refreshOnLoad="1" background="1" description="{c.get("description", "")}">'
                f'<dbPr connection="{c.get("connection_string", "")}" '
                f'command="{c.get("command", "")}"/></connection>'
                for i, c in enumerate(connections)
            )
            zf.writestr(
                "xl/connections.xml",
                f'<?xml version="1.0"?><connections xmlns="{MAIN_NS}">{connection_xml}</connections>',
            )
            workbook_rels.append(
                f'<Relationship Id="rId{rel_number}" Type="{REL_NS}/connections" '
                f'Target="connections.xml"/>'
            )

        if external_link_target:
            rel_number += 1
            zf.writestr(
                "xl/externalLinks/externalLink1.xml",
                f'<?xml version="1.0"?><externalLink xmlns="{MAIN_NS}"><externalBook/></externalLink>',
            )
            zf.writestr(
                "xl/externalLinks/_rels/externalLink1.xml.rels",
                f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL_NS}">'
                f'<Relationship Id="rIdEL" Type="{REL_NS}/externalLinkPath" '
                f'Target="{external_link_target}" TargetMode="External"/></Relationships>',
            )
            workbook_rels.append(
                f'<Relationship Id="rId{rel_number}" Type="{REL_NS}/externalLink" '
                f'Target="externalLinks/externalLink1.xml"/>'
            )

        if mashup_section:
            rel_number += 1
            zf.writestr("customXml/item1.xml", build_mashup_blob(mashup_section))
            workbook_rels.append(
                f'<Relationship Id="rId{rel_number}" Type="{REL_NS}/customXml" '
                f'Target="../customXml/item1.xml"/>'
            )

        if vba is not None:
            rel_number += 1
            zf.writestr("xl/vbaProject.bin", vba)
            workbook_rels.append(
                f'<Relationship Id="rId{rel_number}" Type="{REL_NS}/vbaProject" '
                f'Target="vbaProject.bin"/>'
            )

        names_xml: str = ""
        if defined_names:
            names_xml = "<definedNames>" + "".join(
                f'<definedName name="{name}">{refers}</definedName>' for name, refers in defined_names
            ) + "</definedNames>"

        zf.writestr(
            "xl/workbook.xml",
            f'<?xml version="1.0"?><workbook xmlns="{MAIN_NS}" xmlns:r="{REL_NS}">'
            f'<sheets>{"".join(sheet_entries)}</sheets>{names_xml}</workbook>',
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            f'<?xml version="1.0"?><Relationships xmlns="{PKG_REL_NS}">'
            f'{"".join(workbook_rels)}</Relationships>',
        )
    return path


# --------------------------------------------------------------------------
# BIFF12 (.xlsb)
# --------------------------------------------------------------------------

def _biff_id(record_id: int) -> bytes:
    if record_id < 0x80:
        return bytes([record_id])
    return bytes([(record_id & 0x7F) | 0x80, record_id >> 7])


def _biff_len(length: int) -> bytes:
    out: bytearray = bytearray()
    while True:
        byte: int = length & 0x7F
        length >>= 7
        if length:
            byte |= 0x80
        out.append(byte)
        if not length:
            return bytes(out)


def _biff_record(record_id: int, payload: bytes) -> bytes:
    return _biff_id(record_id) + _biff_len(len(payload)) + payload


def _xl_wide(text: str) -> bytes:
    return struct.pack("<I", len(text)) + text.encode("utf-16-le")


def write_xlsb(path: Path, sheets: Sequence[Tuple[str, Sequence[Sequence[str]], str]]) -> Path:
    """Assemble a real BIFF12 package with a shared string table."""
    path.parent.mkdir(parents=True, exist_ok=True)
    state_codes: Dict[str, int] = {"visible": 0, "hidden": 1, "veryHidden": 2}

    strings: List[str] = []
    for _name, rows, _state in sheets:
        for row in rows:
            for value in row:
                if value not in strings:
                    strings.append(value)
    index_of: Dict[str, int] = {s: i for i, s in enumerate(strings)}

    workbook_stream: bytearray = bytearray()
    workbook_stream += _biff_record(143, b"")                      # BrtBeginBundleShs
    for tab, (name, _rows, state) in enumerate(sheets):
        payload: bytes = (
            struct.pack("<II", state_codes.get(state, 0), tab)
            + _xl_wide(f"rId{tab + 1}")
            + _xl_wide(name)
        )
        workbook_stream += _biff_record(156, payload)               # BrtBundleSh
    workbook_stream += _biff_record(144, b"")                      # BrtEndBundleShs

    shared_stream: bytearray = bytearray()
    for value in strings:
        shared_stream += _biff_record(19, b"\x00" + _xl_wide(value))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("xl/workbook.bin", bytes(workbook_stream))
        zf.writestr("xl/sharedStrings.bin", bytes(shared_stream))
        for index, (_name, rows, _state) in enumerate(sheets, start=1):
            stream: bytearray = bytearray()
            width: int = max((len(r) for r in rows), default=1)
            stream += _biff_record(148, struct.pack("<IIII", 0, max(len(rows) - 1, 0), 0, max(width - 1, 0)))
            for row_index, row in enumerate(rows):
                stream += _biff_record(0, struct.pack("<IIHH", row_index, 0, 300, 0) + b"\x00")
                for column_index, value in enumerate(row):
                    cell: bytes = struct.pack("<I", column_index) + b"\x00\x00\x00\x00"
                    stream += _biff_record(7, cell + struct.pack("<I", index_of[value]))
            zf.writestr(f"xl/worksheets/sheet{index}.bin", bytes(stream))
    return path
