"""Minimal BIFF12 (.xlsb) reader for discovery.

.xlsb stores its workbook and sheets as binary record streams instead of XML,
so the OOXML probe cannot see sheet names or headers. Since the Overstock /
static-inventory workbook is .xlsb, treating that format as a blind spot would
leave a named source undiscovered.

This reads only what Phase 0 needs — sheet names, hidden state, row counts,
and the header band — using the documented BIFF12 record framing. It is
best-effort by design: any structural surprise is reported as an error on the
sheet rather than guessed around.

Record framing
--------------
    record id:     1-2 bytes, 7 bits per byte, high bit = continuation
    record length: 1-4 bytes, 7 bits per byte, high bit = continuation
    payload:       <length> bytes
"""

from __future__ import annotations

import struct
import zipfile
from dataclasses import dataclass, field
from typing import Dict, Iterator, List, Optional, Tuple

from .ooxml import NotAnOoxmlPackage, open_workbook

# Record identifiers used here (MS-XLSB).
BRT_ROW_HDR: int = 0
BRT_CELL_BLANK: int = 1
BRT_CELL_RK: int = 2
BRT_CELL_ERROR: int = 3
BRT_CELL_BOOL: int = 4
BRT_CELL_REAL: int = 5
BRT_CELL_ST: int = 6            # inline string
BRT_CELL_ISST: int = 7          # shared-string index
BRT_FMLA_STRING: int = 8
BRT_FMLA_NUM: int = 9
BRT_FMLA_BOOL: int = 10
BRT_FMLA_ERROR: int = 11
BRT_SST_ITEM: int = 19
BRT_WS_DIM: int = 148
BRT_BUNDLE_SH: int = 156

FORMULA_RECORDS: frozenset[int] = frozenset(
    {BRT_FMLA_STRING, BRT_FMLA_NUM, BRT_FMLA_BOOL, BRT_FMLA_ERROR}
)
HEADER_SCAN_ROWS: int = 25
CELL_STRUCT_BYTES: int = 8      # column (4) + iStyleRef (3) + flags (1)


class BiffError(RuntimeError):
    """Raised when a BIFF12 stream cannot be framed."""


@dataclass
class XlsbSheet:
    name: str
    part: str
    state: str = "visible"
    tab_id: int = 0
    row_count: int = 0
    last_row: int = -1
    last_column: int = -1
    formula_count: int = 0
    header_row: Optional[int] = None
    headers: List[str] = field(default_factory=list)
    error: Optional[str] = None


def _read_varint(data: bytes, offset: int, max_bytes: int) -> Tuple[int, int]:
    """Read a 7-bit continuation-encoded integer. Returns ``(value, new_offset)``."""
    value: int = 0
    for index in range(max_bytes):
        if offset >= len(data):
            raise BiffError(f"truncated varint at offset {offset}")
        byte: int = data[offset]
        offset += 1
        value |= (byte & 0x7F) << (7 * index)
        if not byte & 0x80:
            return value, offset
    return value, offset


def iter_records(data: bytes) -> Iterator[Tuple[int, bytes]]:
    """Yield ``(record_id, payload)`` for a BIFF12 stream."""
    offset: int = 0
    total: int = len(data)
    while offset < total:
        record_id, offset = _read_varint(data, offset, 2)
        length, offset = _read_varint(data, offset, 4)
        if offset + length > total:
            raise BiffError(
                f"record {record_id} declares {length} bytes but only "
                f"{total - offset} remain at offset {offset}"
            )
        yield record_id, data[offset : offset + length]
        offset += length


def _wide_string(payload: bytes, offset: int) -> Tuple[str, int]:
    """Read an XLWideString (uint32 char count + UTF-16LE chars)."""
    if offset + 4 > len(payload):
        raise BiffError("truncated XLWideString length")
    count: int = struct.unpack_from("<I", payload, offset)[0]
    offset += 4
    if count == 0xFFFFFFFF:                      # XLNullableWideString null marker
        return "", offset
    end: int = offset + (count * 2)
    if end > len(payload):
        raise BiffError(f"XLWideString claims {count} chars, payload too short")
    return payload[offset:end].decode("utf-16-le", errors="replace"), end


def read_sheet_bundle(data: bytes) -> List[Tuple[str, int, str]]:
    """Sheet entries from ``xl/workbook.bin`` as ``[(name, tab_id, state)]``."""
    states: Dict[int, str] = {0: "visible", 1: "hidden", 2: "veryHidden"}
    out: List[Tuple[str, int, str]] = []
    for record_id, payload in iter_records(data):
        if record_id != BRT_BUNDLE_SH:
            continue
        if len(payload) < 8:
            continue
        hs_state, tab_id = struct.unpack_from("<II", payload, 0)
        offset: int = 8
        _rel_id, offset = _wide_string(payload, offset)   # strRelID (nullable)
        name, _offset = _wide_string(payload, offset)
        out.append((name, tab_id, states.get(hs_state, f"unknown({hs_state})")))
    return out


def read_shared_strings(data: bytes, cap: int) -> List[str]:
    """Shared string table from ``xl/sharedStrings.bin``."""
    strings: List[str] = []
    for record_id, payload in iter_records(data):
        if record_id != BRT_SST_ITEM or len(payload) < 5:
            continue
        text, _offset = _wide_string(payload, 1)          # 1 flags byte precedes the string
        strings.append(text)
        if len(strings) >= cap:
            break
    return strings


def _cell_column(payload: bytes) -> int:
    if len(payload) < 4:
        return -1
    return int(struct.unpack_from("<I", payload, 0)[0])


def _probe_sheet_stream(sheet: XlsbSheet, data: bytes, shared: List[str], row_cap: int) -> None:
    """Fill counts and the header band for one sheet stream."""
    scanned: Dict[int, Dict[int, str]] = {}
    current_row: int = -1
    physical_rows: int = 0

    for record_id, payload in iter_records(data):
        if record_id == BRT_WS_DIM and len(payload) >= 16:
            row_first, row_last, col_first, col_last = struct.unpack_from("<IIII", payload, 0)
            sheet.last_row = int(row_last)
            sheet.last_column = int(col_last)
            del row_first, col_first
            continue

        if record_id == BRT_ROW_HDR:
            if len(payload) >= 4:
                current_row = int(struct.unpack_from("<I", payload, 0)[0])
            physical_rows += 1
            if physical_rows >= row_cap:
                break
            continue

        if record_id in FORMULA_RECORDS:
            sheet.formula_count += 1

        if current_row < 0 or current_row >= HEADER_SCAN_ROWS:
            continue

        text: Optional[str] = None
        if record_id == BRT_CELL_ISST and len(payload) >= CELL_STRUCT_BYTES + 4:
            index: int = int(struct.unpack_from("<I", payload, CELL_STRUCT_BYTES)[0])
            text = shared[index] if 0 <= index < len(shared) else ""
        elif record_id == BRT_CELL_ST and len(payload) > CELL_STRUCT_BYTES:
            text, _offset = _wide_string(payload, CELL_STRUCT_BYTES)
        elif record_id == BRT_CELL_REAL and len(payload) >= CELL_STRUCT_BYTES + 8:
            text = str(struct.unpack_from("<d", payload, CELL_STRUCT_BYTES)[0])
        elif record_id in (BRT_CELL_RK, BRT_CELL_BOOL, BRT_CELL_ERROR, BRT_CELL_BLANK):
            text = ""

        if text:
            column: int = _cell_column(payload)
            scanned.setdefault(current_row, {})[column if column >= 0 else 0] = text

    sheet.row_count = physical_rows
    # Reuse the XML probe's validated header scoring so both formats agree.
    from .workbook_probe import _choose_header_row

    sheet.header_row, _confidence, sheet.headers = _choose_header_row(
        {row + 1: cells for row, cells in scanned.items()}   # BIFF rows are zero-based
    )


def probe_xlsb(path: str, row_cap: int = 2_000_000, shared_cap: int = 1_000_000) -> List[XlsbSheet]:
    """Sheet-level probe of an .xlsb workbook. Raises for non-package files."""
    sheets: List[XlsbSheet] = []
    with open_workbook(str(path)) as zf:
        names: List[str] = zf.namelist()
        if "xl/workbook.bin" not in names:
            raise NotAnOoxmlPackage(f"{path}: no xl/workbook.bin; not an .xlsb package")

        bundle: List[Tuple[str, int, str]] = read_sheet_bundle(zf.read("xl/workbook.bin"))
        shared: List[str] = (
            read_shared_strings(zf.read("xl/sharedStrings.bin"), shared_cap)
            if "xl/sharedStrings.bin" in names
            else []
        )

        # workbook.bin.rels maps sheets to parts, but BundleSh order matches the
        # sheetN.bin ordering in every file Excel writes, so pair positionally and
        # record a mismatch rather than silently misattributing a sheet.
        parts: List[str] = sorted(
            (n for n in names if n.startswith("xl/worksheets/sheet") and n.endswith(".bin")),
            key=lambda n: int("".join(c for c in n.rsplit("/", 1)[-1] if c.isdigit()) or 0),
        )
        for index, (name, tab_id, state) in enumerate(bundle):
            part: str = parts[index] if index < len(parts) else ""
            sheet: XlsbSheet = XlsbSheet(name=name, part=part, state=state, tab_id=tab_id)
            if not part:
                sheet.error = f"no worksheet stream for sheet '{name}' (found {len(parts)} streams)"
                sheets.append(sheet)
                continue
            try:
                _probe_sheet_stream(sheet, zf.read(part), shared, row_cap)
            except (BiffError, struct.error) as exc:
                sheet.error = f"BIFF12 framing failure in {part}: {exc}"
            sheets.append(sheet)
    return sheets
