"""Structural probe for a single OOXML workbook (.xlsx/.xlsm/.xltx/.xltm).

Answers, without opening Excel: what sheets exist, which are hidden, how many
rows each holds, what the header row actually is, which named tables and ranges
are defined, what external data each query/connection/link points at, and
whether the file carries VBA.

Design notes
------------
* Single streaming pass per sheet. Row bodies are only materialized for the
  first ``header_scan_rows`` rows; everything after that is counted, not stored.
* Header detection is *validated*, not assumed to be row 1. The spec forbids
  fragile positional cell references, so the probe reports which row it chose
  and how confident it is, and downstream readers anchor on those names.
* Formula density is recorded because it is the strongest available signal for
  "downstream copy" versus "true source".
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Set, Tuple
from xml.etree import ElementTree as ET

from .ooxml import (
    NS_MAIN,
    NS_REL,
    NotAnOoxmlPackage,
    Relationship,
    cell_column_letters,
    column_index,
    iter_shared_strings,
    local_name,
    open_workbook,
    qn,
    read_rels,
    resolve_part,
)

HEADER_SCAN_ROWS: int = 25
SHARED_STRING_CAP: int = 1_000_000
MIN_HEADER_STRING_RATIO: float = 0.60


@dataclass
class SheetProbe:
    """What we learned about one worksheet."""

    name: str
    sheet_id: str
    part: str
    state: str = "visible"            # visible | hidden | veryHidden
    dimension: str = ""
    row_count: int = 0
    rows_with_values: int = 0
    value_rows_in_scan_window: Set[int] = field(default_factory=set)
    row_count_capped: bool = False
    max_column_index: int = -1
    formula_count: int = 0
    shared_formula_count: int = 0
    header_row: Optional[int] = None
    header_confidence: float = 0.0
    headers: List[str] = field(default_factory=list)
    tables: List[Dict[str, str]] = field(default_factory=list)
    error: Optional[str] = None

    @property
    def is_hidden(self) -> bool:
        return self.state != "visible"

    @property
    def data_row_estimate(self) -> int:
        """Usable record count: rows holding values, below the detected header.

        Uses rows_with_values rather than the physical row count. A sheet padded
        to Excel's row limit by formatting reports 1,048,576 physical rows and
        almost no data; subtracting the header from that number is nonsense.
        """
        basis: int = self.rows_with_values or self.row_count
        if self.header_row is None:
            return max(basis, 0)
        # Subtract only the value-bearing rows at or above the header, so a
        # count of populated rows is never reduced by a physical row index.
        consumed: int = sum(1 for row in self.value_rows_in_scan_window if row <= self.header_row)
        return max(basis - consumed, 0)


@dataclass
class WorkbookProbe:
    """Full structural picture of one workbook."""

    path: str
    parts: List[str] = field(default_factory=list)
    sheets: List[SheetProbe] = field(default_factory=list)
    defined_names: List[Dict[str, str]] = field(default_factory=list)
    connections: List[Dict[str, str]] = field(default_factory=list)
    external_links: List[str] = field(default_factory=list)
    has_vba: bool = False
    vba_modules: List[str] = field(default_factory=list)
    has_power_query: bool = False
    power_query_part: Optional[str] = None
    shared_strings_truncated: bool = False
    pivot_cache_count: int = 0
    errors: List[str] = field(default_factory=list)

    @property
    def visible_sheets(self) -> List[str]:
        return [s.name for s in self.sheets if not s.is_hidden]

    @property
    def hidden_sheets(self) -> List[str]:
        return [s.name for s in self.sheets if s.is_hidden]

    @property
    def total_rows(self) -> int:
        return sum(s.row_count for s in self.sheets)

    @property
    def total_formulas(self) -> int:
        return sum(s.formula_count for s in self.sheets)


def _sheet_entries(zf: zipfile.ZipFile) -> Tuple[List[Tuple[str, str, str, str]], List[Dict[str, str]]]:
    """Read ``xl/workbook.xml``: sheet list and defined names.

    Returns ``([(name, sheet_id, rel_id, state)], [defined_name_dicts])``.
    """
    part: str = "xl/workbook.xml"
    if part not in zf.namelist():
        raise NotAnOoxmlPackage(f"missing {part}; not a spreadsheet package")
    with zf.open(part) as handle:
        root: ET.Element = ET.parse(handle).getroot()

    sheets: List[Tuple[str, str, str, str]] = []
    sheets_node: Optional[ET.Element] = root.find(qn(NS_MAIN, "sheets"))
    if sheets_node is not None:
        for node in sheets_node.findall(qn(NS_MAIN, "sheet")):
            sheets.append(
                (
                    node.get("name", ""),
                    node.get("sheetId", ""),
                    node.get(qn(NS_REL, "id"), ""),
                    node.get("state", "visible"),
                )
            )

    names: List[Dict[str, str]] = []
    names_node: Optional[ET.Element] = root.find(qn(NS_MAIN, "definedNames"))
    if names_node is not None:
        for node in names_node.findall(qn(NS_MAIN, "definedName")):
            names.append(
                {
                    "name": node.get("name", ""),
                    "scope_sheet_index": node.get("localSheetId", ""),
                    "hidden": node.get("hidden", "false"),
                    "refers_to": (node.text or "").strip(),
                }
            )
    return sheets, names


def _choose_header_row(rows: Dict[int, Dict[int, str]]) -> Tuple[Optional[int], float, List[str]]:
    """Pick the header row from the first scanned rows and score the choice.

    Scoring favours rows that are wide, mostly text, and mostly distinct —
    the shape of a real header band. The confidence is returned so the manifest
    can flag workbooks where header detection is shaky and a human must look.
    """
    best_row: Optional[int] = None
    best_score: float = 0.0
    best_values: List[str] = []

    for row_index in sorted(rows):
        cells: Dict[int, str] = rows[row_index]
        values: List[str] = [cells[c] for c in sorted(cells) if str(cells[c]).strip() != ""]
        if len(values) < 2:
            continue
        text_values: List[str] = [v for v in values if not _looks_numeric(v)]
        text_ratio: float = len(text_values) / len(values)
        if text_ratio < MIN_HEADER_STRING_RATIO:
            continue
        distinct_ratio: float = len(set(v.strip().lower() for v in values)) / len(values)
        width_score: float = min(len(values) / 12.0, 1.0)
        score: float = (0.45 * text_ratio) + (0.35 * distinct_ratio) + (0.20 * width_score)
        # Earlier rows win ties; a later row must be meaningfully better.
        if score > best_score + 1e-9:
            best_row, best_score, best_values = row_index, score, values

    return best_row, round(best_score, 3), best_values


def _looks_numeric(value: str) -> bool:
    try:
        float(str(value).replace(",", ""))
        return True
    except (TypeError, ValueError):
        return False


def _probe_sheet(
    zf: zipfile.ZipFile,
    sheet: SheetProbe,
    shared: Sequence[str],
    row_cap: int,
) -> None:
    """Stream one worksheet part, filling counts, headers and table links."""
    if sheet.part not in zf.namelist():
        sheet.error = f"worksheet part not in package: {sheet.part}"
        return

    row_tag: str = qn(NS_MAIN, "row")
    cell_tag: str = qn(NS_MAIN, "c")
    value_tag: str = qn(NS_MAIN, "v")
    inline_tag: str = qn(NS_MAIN, "is")
    text_tag: str = qn(NS_MAIN, "t")
    formula_tag: str = qn(NS_MAIN, "f")
    dimension_tag: str = qn(NS_MAIN, "dimension")

    scanned: Dict[int, Dict[int, str]] = {}
    physical_row: int = 0

    try:
        with zf.open(sheet.part) as handle:
            for _event, element in ET.iterparse(handle, events=("end",)):
                tag: str = element.tag
                if tag == dimension_tag:
                    sheet.dimension = element.get("ref", "")
                    element.clear()
                    continue
                if tag != row_tag:
                    continue

                physical_row += 1
                declared: str = element.get("r", "")
                row_number: int = int(declared) if declared.isdigit() else physical_row
                row_has_value: bool = False

                for cell in element.findall(cell_tag):
                    col: int = column_index(cell_column_letters(cell.get("r")))
                    if col > sheet.max_column_index:
                        sheet.max_column_index = col
                    formula: Optional[ET.Element] = cell.find(formula_tag)
                    if formula is not None:
                        sheet.formula_count += 1
                        if formula.get("t") == "shared":
                            sheet.shared_formula_count += 1
                    if cell.find(value_tag) is not None or cell.find(inline_tag) is not None:
                        row_has_value = True
                    if row_number <= HEADER_SCAN_ROWS:
                        text: str = _cell_text(cell, shared, value_tag, inline_tag, text_tag)
                        if text != "":
                            scanned.setdefault(row_number, {})[col if col >= 0 else 0] = text

                if row_has_value:
                    sheet.rows_with_values += 1
                    if row_number <= HEADER_SCAN_ROWS:
                        sheet.value_rows_in_scan_window.add(row_number)
                element.clear()
                if physical_row >= row_cap:
                    sheet.row_count_capped = True
                    break
    except ET.ParseError as exc:
        sheet.error = f"XML parse failure in {sheet.part}: {exc}"

    sheet.row_count = physical_row
    sheet.header_row, sheet.header_confidence, sheet.headers = _choose_header_row(scanned)


def _cell_text(
    cell: ET.Element,
    shared: Sequence[str],
    value_tag: str,
    inline_tag: str,
    text_tag: str,
) -> str:
    """Resolve a cell's displayed text across the shared/inline/literal encodings."""
    cell_type: str = cell.get("t", "n")
    if cell_type == "s":
        node: Optional[ET.Element] = cell.find(value_tag)
        if node is None or not (node.text or "").strip().isdigit():
            return ""
        index: int = int(node.text.strip())
        return shared[index] if 0 <= index < len(shared) else ""
    if cell_type == "inlineStr":
        node = cell.find(inline_tag)
        return "".join(t.text or "" for t in node.iter(text_tag)) if node is not None else ""
    node = cell.find(value_tag)
    return (node.text or "").strip() if node is not None else ""


def _read_connections(zf: zipfile.ZipFile) -> List[Dict[str, str]]:
    """Parse ``xl/connections.xml`` — the external sources Excel refreshes from."""
    part: str = "xl/connections.xml"
    if part not in zf.namelist():
        return []
    with zf.open(part) as handle:
        root: ET.Element = ET.parse(handle).getroot()
    out: List[Dict[str, str]] = []
    for node in root.findall(qn(NS_MAIN, "connection")):
        record: Dict[str, str] = {
            "name": node.get("name", ""),
            "description": node.get("description", ""),
            "type": node.get("type", ""),
            "refresh_on_load": node.get("refreshOnLoad", "0"),
            "background_refresh": node.get("background", "1"),
            "connection_string": "",
            "command": "",
            "source_file": "",
        }
        for child in node:
            record["connection_string"] = child.get("connection", record["connection_string"])
            record["command"] = child.get("command", record["command"])
            record["source_file"] = child.get("sourceFile", record["source_file"])
        out.append(record)
    return out


def _read_external_links(zf: zipfile.ZipFile, workbook_rels: Dict[str, Relationship]) -> List[str]:
    """Every other workbook this file formula-links to (a hard dependency)."""
    targets: List[str] = []
    for rel in workbook_rels.values():
        if not rel.type_is("/externalLink"):
            continue
        for child_rel in read_rels(zf, rel.resolved).values():
            if child_rel.type_is("/externalLinkPath"):
                targets.append(child_rel.target)
    return targets


def _sheet_tables(zf: zipfile.ZipFile, sheet_part: str) -> List[Dict[str, str]]:
    """Named tables (ListObjects) anchored on a sheet."""
    out: List[Dict[str, str]] = []
    for rel in read_rels(zf, sheet_part).values():
        if not rel.type_is("/table") or rel.resolved not in zf.namelist():
            continue
        with zf.open(rel.resolved) as handle:
            root: ET.Element = ET.parse(handle).getroot()
        columns: List[str] = [
            column.get("name", "")
            for column in root.iter(qn(NS_MAIN, "tableColumn"))
        ]
        out.append(
            {
                "name": root.get("displayName") or root.get("name", ""),
                "ref": root.get("ref", ""),
                "columns": "|".join(columns),
                "column_count": str(len(columns)),
                "part": rel.resolved,
            }
        )
    return out


def probe_workbook(path: str, row_cap: int = 2_000_000) -> WorkbookProbe:
    """Read-only structural probe of an OOXML workbook.

    Never writes, never executes a macro, never opens Excel.
    """
    probe: WorkbookProbe = WorkbookProbe(path=str(path))
    with open_workbook(str(path)) as zf:
        probe.parts = zf.namelist()
        probe.has_vba = "xl/vbaProject.bin" in probe.parts
        probe.pivot_cache_count = sum(
            1 for p in probe.parts if p.startswith("xl/pivotCache/pivotCacheDefinition")
        )

        shared: List[str]
        shared, probe.shared_strings_truncated = iter_shared_strings(zf, SHARED_STRING_CAP)

        workbook_rels: Dict[str, Relationship] = read_rels(zf, "xl/workbook.xml")
        sheet_entries, probe.defined_names = _sheet_entries(zf)

        for name, sheet_id, rel_id, state in sheet_entries:
            rel: Optional[Relationship] = workbook_rels.get(rel_id)
            part: str = rel.resolved if rel is not None else ""
            sheet: SheetProbe = SheetProbe(name=name, sheet_id=sheet_id, part=part, state=state)
            if not part:
                sheet.error = f"sheet '{name}' has unresolvable relationship id '{rel_id}'"
                probe.sheets.append(sheet)
                continue
            _probe_sheet(zf, sheet, shared, row_cap)
            sheet.tables = _sheet_tables(zf, part)
            probe.sheets.append(sheet)

        probe.connections = _read_connections(zf)
        probe.external_links = _read_external_links(zf, workbook_rels)

    for sheet in probe.sheets:
        if sheet.error:
            probe.errors.append(sheet.error)
    return probe
