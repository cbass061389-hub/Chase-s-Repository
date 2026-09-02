"""Reconstruct what a calculation engine computes, from its formulas.

Five sheets in the estate hold most of its logic — Position Engine (~60k
formulas), Allocation Plan (~48k), __CleanBO (~37k), BOMMaster (~35k),
__Alloc Engine (~25k) — and all of them are hidden or veryHidden behind a
password-protected VBA project. The formulas themselves are readable, so the
computation can be recovered even where the macros cannot.

The method is pattern collapse. A column of 5,000 formulas is almost always one
formula repeated with the row number advancing, so normalizing row numbers to
``{r}`` turns 5,000 cells into one pattern. What is then interesting is the
exceptions:

* a column with **one** pattern is a computed field, and the pattern is its
  definition;
* a column with a dominant pattern and a handful of outliers is a hand-edited
  cell in a formula column — nearly always a bug, and invisible in the grid;
* volatile and fragile constructs (INDIRECT, OFFSET, whole-column references)
  explain why a workbook is slow and why it breaks when rows move.
"""

from __future__ import annotations

import re
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Sequence, Set, Tuple
from xml.etree import ElementTree as ET

from ..discovery.ooxml import (
    NS_MAIN,
    NotAnOoxmlPackage,
    Relationship,
    cell_column_letters,
    column_index,
    iter_shared_strings,
    open_workbook,
    qn,
    read_rels,
)

#: Functions whose result depends on something other than their inputs, or that
#: defeat dependency tracking. Each one is a reason a workbook recalculates
#: slowly or silently breaks when rows are inserted.
VOLATILE_FUNCTIONS: Tuple[str, ...] = (
    "INDIRECT", "OFFSET", "TODAY", "NOW", "RAND", "RANDBETWEEN", "CELL", "INFO",
)

#: Lookup styles, tracked because migrating them is the usual modernization step.
LOOKUP_FUNCTIONS: Tuple[str, ...] = (
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "INDEX", "MATCH", "XMATCH",
)

#: Excel error literals. A formula whose *text* contains one is not a formula
#: that failed at calculation time — it is a formula whose reference was
#: destroyed, usually by a delete that moved rows out from under it.
ERROR_LITERALS: Tuple[str, ...] = (
    "#REF!", "#NAME?", "#VALUE!", "#DIV/0!", "#NULL!", "#NUM!",
)

_CELL_REFERENCE: re.Pattern[str] = re.compile(r"(\$?[A-Z]{1,3})\$?(\d+)")
_FUNCTION_CALL: re.Pattern[str] = re.compile(r"\b([A-Z][A-Z0-9._]{1,})\s*\(")
_SHEET_REFERENCE: re.Pattern[str] = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))\s*!")
_WHOLE_COLUMN: re.Pattern[str] = re.compile(r"\$?[A-Z]{1,3}:\$?[A-Z]{1,3}\b")


_LITERAL: re.Pattern[str] = re.compile(r'^\s*(?:"(?:[^"]|"")*"|-?[\d.,]+%?)\s*$')


def _is_literal(pattern: str) -> bool:
    """True for a bare string or numeric constant.

    A status column holding "No completion data (non-BCP)" is a literal despite
    containing parentheses, so the test has to be about the shape of the whole
    expression rather than the characters in it.
    """
    return bool(_LITERAL.match(pattern))


def normalize_formula(formula: str) -> str:
    """Collapse a formula to its pattern by replacing row numbers with ``{r}``.

    ``IF(B5>8,"OVERSTOCK","MONITOR")`` and the same formula on row 6 become one
    pattern, which is what makes 60,000 formulas comprehensible.
    """
    return _CELL_REFERENCE.sub(lambda m: f"{m.group(1)}{{r}}", str(formula).strip())


@dataclass
class ColumnFormulas:
    """Every formula found in one column of one sheet."""

    column_index: int
    column_letter: str
    header: str = ""
    patterns: Counter = field(default_factory=Counter)
    first_row: int = 0
    last_row: int = 0
    broken_reference_cells: int = 0

    @property
    def total(self) -> int:
        return sum(self.patterns.values())

    @property
    def dominant(self) -> Tuple[str, int]:
        return self.patterns.most_common(1)[0] if self.patterns else ("", 0)

    @property
    def distinct_patterns(self) -> int:
        return len(self.patterns)

    @property
    def consistency(self) -> float:
        """Share of cells matching the dominant pattern. 1.0 is a clean column."""
        return (self.dominant[1] / self.total) if self.total else 0.0

    @property
    def outliers(self) -> List[Tuple[str, int]]:
        """Patterns that are not the dominant one — candidate hand-edits."""
        dominant_pattern: str = self.dominant[0]
        return [(p, n) for p, n in self.patterns.most_common() if p != dominant_pattern]

    @property
    def has_broken_references(self) -> bool:
        return self.broken_reference_cells > 0

    @property
    def is_constant_column(self) -> bool:
        """True when every pattern is a literal — no function call, no reference.

        A status column holding four different text literals is legitimately
        varied and must not be reported as an inconsistent formula column.
        """
        return bool(self.patterns) and all(_is_literal(pattern) for pattern in self.patterns)

    @property
    def consistency_label(self) -> str:
        """Consistency to one decimal, never rounded up to a clean 100%."""
        share: float = self.consistency
        if share < 1.0 and share > 0.999:
            return "99.9%"
        return f"{share * 100:.1f}%"

    @property
    def is_suspicious(self) -> bool:
        """A mostly-uniform formula column with exceptions is nearly always a bug.

        Constant columns are excluded: varied text literals are not a defect.
        """
        if self.is_constant_column:
            return False
        return self.total >= 20 and 0.0 < self.consistency < 0.98


@dataclass
class SheetFormulas:
    """The computation one sheet performs."""

    workbook: str
    sheet: str
    state: str = "visible"
    formula_count: int = 0
    columns: List[ColumnFormulas] = field(default_factory=list)
    referenced_sheets: Set[str] = field(default_factory=set)
    functions: Counter = field(default_factory=Counter)
    whole_column_references: int = 0
    broken_reference_cells: int = 0
    error: Optional[str] = None

    @property
    def volatile_usage(self) -> Dict[str, int]:
        return {name: self.functions[name] for name in VOLATILE_FUNCTIONS if self.functions[name]}

    @property
    def lookup_usage(self) -> Dict[str, int]:
        return {name: self.functions[name] for name in LOOKUP_FUNCTIONS if self.functions[name]}

    @property
    def computed_columns(self) -> List[ColumnFormulas]:
        """Columns that are genuinely formula columns, worst-consistency first."""
        return sorted(
            [c for c in self.columns if c.total >= 5],
            key=lambda c: (c.consistency, -c.total),
        )

    @property
    def suspicious_columns(self) -> List[ColumnFormulas]:
        return [c for c in self.columns if c.is_suspicious]

    @property
    def broken_columns(self) -> List[ColumnFormulas]:
        """Columns carrying destroyed references, worst first."""
        return sorted(
            [c for c in self.columns if c.has_broken_references],
            key=lambda c: -c.broken_reference_cells,
        )


def _sheet_targets(zf: zipfile.ZipFile) -> List[Tuple[str, str, str]]:
    """``(sheet_name, state, part)`` for every worksheet."""
    with zf.open("xl/workbook.xml") as handle:
        root: ET.Element = ET.parse(handle).getroot()
    rels: Dict[str, Relationship] = read_rels(zf, "xl/workbook.xml")
    out: List[Tuple[str, str, str]] = []
    node = root.find(qn(NS_MAIN, "sheets"))
    if node is None:
        return out
    for sheet in node.findall(qn(NS_MAIN, "sheet")):
        relationship = rels.get(sheet.get(qn("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id"), ""))
        if relationship is None:
            continue
        out.append((sheet.get("name", ""), sheet.get("state", "visible"), relationship.resolved))
    return out


def map_sheet_formulas(
    zf: zipfile.ZipFile,
    workbook: str,
    sheet_name: str,
    state: str,
    part: str,
    shared_strings: Sequence[str],
    header_row_hint: Optional[int] = None,
) -> SheetFormulas:
    """Stream one sheet and collapse its formulas into per-column patterns."""
    result: SheetFormulas = SheetFormulas(workbook=workbook, sheet=sheet_name, state=state)
    if part not in zf.namelist():
        result.error = f"worksheet part missing: {part}"
        return result

    row_tag: str = qn(NS_MAIN, "row")
    cell_tag: str = qn(NS_MAIN, "c")
    formula_tag: str = qn(NS_MAIN, "f")
    value_tag: str = qn(NS_MAIN, "v")

    by_column: Dict[int, ColumnFormulas] = {}
    shared_masters: Dict[str, str] = {}
    header_candidates: Dict[int, str] = {}
    header_row: int = header_row_hint or 1

    try:
        with zf.open(part) as handle:
            for _event, element in ET.iterparse(handle, events=("end",)):
                if element.tag != row_tag:
                    continue
                declared: str = element.get("r", "")
                row_number: int = int(declared) if declared.isdigit() else 0

                for cell in element.findall(cell_tag):
                    letters: str = cell_column_letters(cell.get("r"))
                    index: int = column_index(letters)

                    # Capture the header band for labelling the columns.
                    if row_number == header_row and cell.get("t") == "s":
                        node = cell.find(value_tag)
                        if node is not None and (node.text or "").strip().isdigit():
                            position: int = int(node.text.strip())
                            if 0 <= position < len(shared_strings):
                                header_candidates[index] = shared_strings[position]

                    formula_node = cell.find(formula_tag)
                    if formula_node is None:
                        continue

                    text: str = (formula_node.text or "").strip()
                    shared_id: Optional[str] = formula_node.get("si")
                    kind: str = formula_node.get("t", "")

                    if kind == "shared":
                        if text and shared_id is not None:
                            shared_masters[shared_id] = text          # master carries the text
                        elif shared_id is not None:
                            text = shared_masters.get(shared_id, "")   # follower reuses it
                    if not text:
                        continue

                    result.formula_count += 1
                    result.functions.update(
                        match.group(1) for match in _FUNCTION_CALL.finditer(text)
                    )
                    for match in _SHEET_REFERENCE.finditer(text):
                        referenced: str = (match.group(1) or match.group(2) or "").strip()
                        if referenced and referenced != sheet_name:
                            result.referenced_sheets.add(referenced)
                    result.whole_column_references += len(_WHOLE_COLUMN.findall(text))
                    is_broken: bool = any(literal in text for literal in ERROR_LITERALS)
                    if is_broken:
                        result.broken_reference_cells += 1

                    column: ColumnFormulas = by_column.setdefault(
                        index, ColumnFormulas(column_index=index, column_letter=letters or "?")
                    )
                    column.patterns.update([normalize_formula(text)])
                    if is_broken:
                        column.broken_reference_cells += 1
                    column.first_row = min(column.first_row or row_number, row_number)
                    column.last_row = max(column.last_row, row_number)

                element.clear()
    except ET.ParseError as exc:
        result.error = f"XML parse failure in {part}: {exc}"

    for index, column in by_column.items():
        column.header = header_candidates.get(index, "")
    result.columns = sorted(by_column.values(), key=lambda c: c.column_index)
    return result


def map_workbook_formulas(
    path: str,
    min_formulas: int = 100,
    sheet_names: Optional[Sequence[str]] = None,
    header_rows: Optional[Dict[str, int]] = None,
) -> List[SheetFormulas]:
    """Map every sheet in a workbook carrying at least *min_formulas* formulas.

    Sheets are mapped in one pass each; a sheet below the threshold is skipped
    before its formulas are collapsed, so a 76-tab workbook stays fast.
    """
    from pathlib import Path

    workbook_label: str = Path(path).name
    wanted: Optional[Set[str]] = set(sheet_names) if sheet_names else None
    results: List[SheetFormulas] = []

    with open_workbook(str(path)) as zf:
        shared_strings, _truncated = iter_shared_strings(zf, 1_000_000)
        for name, state, part in _sheet_targets(zf):
            if wanted is not None and name not in wanted:
                continue
            mapped: SheetFormulas = map_sheet_formulas(
                zf, workbook_label, name, state, part, shared_strings,
                (header_rows or {}).get(name),
            )
            if mapped.formula_count >= min_formulas or (wanted is not None and name in wanted):
                results.append(mapped)

    return sorted(results, key=lambda sheet: -sheet.formula_count)
