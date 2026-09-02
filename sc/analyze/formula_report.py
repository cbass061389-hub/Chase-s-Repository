"""Render the calculation-engine map."""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Sequence

from .formula_map import LOOKUP_FUNCTIONS, VOLATILE_FUNCTIONS, SheetFormulas


def _table(rows: Sequence[Sequence[str]], headers: Sequence[str]) -> List[str]:
    lines: List[str] = ["| " + " | ".join(headers) + " |",
                        "|" + "|".join("---" for _ in headers) + "|"]
    lines.extend("| " + " | ".join(str(c).replace("|", "\\|") for c in row) + " |" for row in rows)
    return lines


def render(sheets: Sequence[SheetFormulas], generated_at: str, min_formulas: int) -> str:
    lines: List[str] = []
    add = lines.append

    suspicious = [(s, c) for s in sheets for c in s.suspicious_columns]
    volatile = [s for s in sheets if s.volatile_usage]
    broken = [(s, c) for s in sheets for c in s.broken_columns]
    broken_cells = sum(c.broken_reference_cells for _s, c in broken)

    add("# Calculation engine map")
    add("")
    add(f"Generated {generated_at}. Sheets with at least {min_formulas:,} formulas, read "
        "directly from the workbook XML — no Excel, no macro execution.")
    add("")

    add("## Executive summary")
    add("")
    add(f"- **{len(sheets)} calculation sheets** hold "
        f"{sum(s.formula_count for s in sheets):,} formulas between them.")
    if broken:
        add(f"- **{broken_cells:,} formula cells carry a destroyed reference** "
            f"(`#REF!` and friends) across {len(broken)} column(s). These are not results that "
            "failed to calculate — the references themselves are gone, so those rows are "
            "producing nothing and have been since whatever delete broke them.")
    add(f"- **{len(suspicious)} column(s) contain a mix of formulas** where one pattern dominates. "
        "In a formula column that is normally a hand-edited cell, and it is invisible in the grid.")
    add(f"- {len(volatile)} sheet(s) use volatile or dependency-breaking functions "
        f"({', '.join(VOLATILE_FUNCTIONS[:4])}...).")
    add("")
    add("Formulas are collapsed by replacing row numbers with `{r}`, so a column of 5,000 "
        "formulas shows as the one pattern it actually is. A column with a single pattern is a "
        "computed field and the pattern is its definition — that is what makes this logic "
        "portable even though the VBA around it is password-protected.")
    add("")

    if broken:
        add("### Destroyed references — fix these first")
        add("")
        rows: List[List[str]] = [
            [
                f"`{sheet.sheet}`",
                f"{column.column_letter} — {column.header or '(no header)'}",
                f"{column.broken_reference_cells:,}",
                f"{column.total:,}",
                f"{column.broken_reference_cells / column.total:.0%}",
                (column.outliers[0][0] if column.outliers else column.dominant[0])[:60],
            ]
            for sheet, column in sorted(broken, key=lambda pair: -pair[1].broken_reference_cells)
        ]
        lines.extend(_table(rows, ["Sheet", "Column", "Broken cells", "Total cells",
                                   "Share", "Broken pattern"]))
        add("")
        add("A `#REF!` inside the formula text means the range it pointed at was deleted. "
            "Every one of these cells is silently contributing nothing to whatever depends on "
            "it, and no total anywhere shows a gap.")
        add("")

    if suspicious:
        add("### Columns that need a look")
        add("")
        suspicious_rows: List[List[str]] = [
            [
                f"`{sheet.sheet}`",
                f"{column.column_letter} — {column.header or '(no header)'}",
                f"{column.total:,}",
                str(column.distinct_patterns),
                column.consistency_label,
                f"{column.outliers[0][1]:,} cell(s) differ" if column.outliers else "",
            ]
            for sheet, column in sorted(suspicious, key=lambda pair: pair[1].consistency)
        ]
        lines.extend(_table(suspicious_rows, ["Sheet", "Column", "Cells", "Patterns",
                                              "Consistency", "Exception"]))
        add("")
        add("Each of these is either a deliberate two-block structure or a formula somebody "
            "overwrote. Both are worth knowing before the logic is ported.")
        add("")

    add("## Sheets, by size")
    add("")
    overview: List[List[str]] = [
        [
            f"`{sheet.sheet}`",
            "" if sheet.state == "visible" else sheet.state,
            f"{sheet.formula_count:,}",
            str(len(sheet.columns)),
            ", ".join(f"{k}x{v:,}" for k, v in sorted(sheet.lookup_usage.items())) or "—",
            ", ".join(f"{k}x{v}" for k, v in sorted(sheet.volatile_usage.items())) or "—",
            f"{sheet.whole_column_references:,}" if sheet.whole_column_references else "—",
            f"{sheet.broken_reference_cells:,}" if sheet.broken_reference_cells else "—",
            ", ".join(sorted(sheet.referenced_sheets)[:5]) or "(self-contained)",
        ]
        for sheet in sheets
    ]
    lines.extend(_table(overview, ["Sheet", "State", "Formulas", "Formula cols", "Lookups",
                                   "Volatile", "Whole-col refs", "Broken refs", "Reads from"]))
    add("")

    add("## Recovered definitions")
    add("")
    add("Per sheet, the formula columns and what each one computes. This is the specification "
        "for porting the logic to the canonical layer.")
    add("")
    for sheet in sheets:
        add(f"### `{sheet.sheet}`"
            + ("" if sheet.state == "visible" else f" [{sheet.state}]")
            + f" — {sheet.formula_count:,} formulas")
        add("")
        if sheet.referenced_sheets:
            add(f"Reads from: {', '.join(f'`{name}`' for name in sorted(sheet.referenced_sheets))}")
            add("")
        if sheet.error:
            add(f"**Probe error:** {sheet.error}")
            add("")
            continue

        for column in sorted(sheet.computed_columns, key=lambda c: c.column_index):
            pattern, count = column.dominant
            label: str = column.header or "(no header)"
            add(f"**{column.column_letter} — {label}** · {column.total:,} cells · "
                f"rows {column.first_row}-{column.last_row} · "
                f"{column.distinct_patterns} pattern(s) · {column.consistency_label} consistent")
            add("")
            add("```excel")
            add(pattern)
            add("```")
            if column.outliers:
                add("")
                add(f"{len(column.outliers)} other pattern(s) in this column, "
                    f"largest with {column.outliers[0][1]:,} cell(s):")
                add("")
                add("```excel")
                for other_pattern, other_count in column.outliers[:2]:
                    add(f"{other_pattern}   /* x{other_count:,} */")
                add("```")
            add("")
    return "\n".join(lines)


def write_reports(
    sheets: Sequence[SheetFormulas], out_dir: Path, generated_at: str, min_formulas: int
) -> List[Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    markdown: Path = out_dir / "FORMULA_MAP.md"
    markdown.write_text(render(sheets, generated_at, min_formulas), encoding="utf-8")

    payload: Dict[str, Any] = {
        "generated_at": generated_at,
        "min_formulas": min_formulas,
        "sheets": [
            {
                "workbook": sheet.workbook,
                "sheet": sheet.sheet,
                "state": sheet.state,
                "formula_count": sheet.formula_count,
                "referenced_sheets": sorted(sheet.referenced_sheets),
                "lookup_usage": sheet.lookup_usage,
                "volatile_usage": sheet.volatile_usage,
                "whole_column_references": sheet.whole_column_references,
                "broken_reference_cells": sheet.broken_reference_cells,
                "error": sheet.error,
                "columns": [
                    {
                        "column": column.column_letter,
                        "header": column.header,
                        "cells": column.total,
                        "first_row": column.first_row,
                        "last_row": column.last_row,
                        "distinct_patterns": column.distinct_patterns,
                        "consistency": round(column.consistency, 4),
                        "suspicious": column.is_suspicious,
                        "broken_reference_cells": column.broken_reference_cells,
                        "dominant_pattern": column.dominant[0],
                        "other_patterns": [
                            {"pattern": p, "cells": n} for p, n in column.outliers[:5]
                        ],
                    }
                    for column in sheet.columns
                ],
            }
            for sheet in sheets
        ],
    }
    json_path: Path = out_dir / "formula_map.json"
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return [markdown, json_path]
