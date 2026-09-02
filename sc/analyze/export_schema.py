"""Recover each upstream's schema from the M that reads it.

Every query's ``Table.TransformColumnTypes`` step enumerates the columns of the
export and their types. That makes the export's schema recoverable from the
committed M source — no credentials, no fetch, no Excel. It is the contract the
extraction layer validates against, and it is how the readers were written
without ever holding the data.

Where two queries declare the same export, the declarations are merged and any
disagreement on a column's type is reported rather than silently resolved.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

from .m_ast import QueryProfile

#: M type -> the pandas dtype the reader should use.
M_TYPE_TO_PANDAS: Dict[str, str] = {
    "Int64.Type": "Int64",
    "type number": "Float64",
    "type text": "string",
    "type date": "datetime64[ns]",
    "type datetime": "datetime64[ns]",
    "type datetimezone": "datetime64[ns]",
    "type time": "string",
    "type logical": "boolean",
    "type any": "string",
    "Currency.Type": "Float64",
    "Percentage.Type": "Float64",
}


def pandas_dtype_for(m_type: str) -> str:
    """Map an M type to a pandas dtype, defaulting to string.

    Defaulting to string is deliberate: a wrong numeric coercion silently
    changes values, while a string that should have been numeric fails loudly
    at the first arithmetic.
    """
    cleaned: str = re.sub(r"\s+", " ", str(m_type)).strip().rstrip(",")
    cleaned = re.sub(r"^type nullable ", "type ", cleaned)
    return M_TYPE_TO_PANDAS.get(cleaned, "string")


@dataclass
class ColumnSpec:
    name: str
    m_type: str
    pandas_dtype: str
    declared_by: List[str] = field(default_factory=list)
    conflicting_types: List[str] = field(default_factory=list)

    @property
    def has_conflict(self) -> bool:
        return bool(self.conflicting_types)


@dataclass
class ExportSpec:
    """One upstream and the schema its readers declare."""

    key: str
    label: str
    kind: str
    columns: List[ColumnSpec] = field(default_factory=list)
    declared_by: List[str] = field(default_factory=list)
    partial: bool = False

    @property
    def column_names(self) -> List[str]:
        return [column.name for column in self.columns]

    @property
    def conflicts(self) -> List[ColumnSpec]:
        return [column for column in self.columns if column.has_conflict]

    @property
    def slug(self) -> str:
        """Stable identifier usable as a filename and a table name."""
        media = re.search(r"media:(\d+)", self.key)
        if media is not None:
            return f"netsuite_{media.group(1)}"
        return re.sub(r"[^a-z0-9]+", "_", self.label.lower()).strip("_")[:60] or "unknown"


def build_export_specs(
    profiles: Sequence[QueryProfile],
    kind_by_key: Optional[Dict[str, str]] = None,
    min_columns_for_complete: int = 5,
) -> List[ExportSpec]:
    """Merge every query's declared schema into one spec per upstream."""
    kinds: Dict[str, str] = kind_by_key or {}
    specs: Dict[str, ExportSpec] = {}

    for profile in profiles:
        if not profile.upstream_key or not profile.column_types:
            continue
        spec: ExportSpec = specs.setdefault(
            profile.upstream_key,
            ExportSpec(
                key=profile.upstream_key,
                label=profile.upstream_label,
                kind=kinds.get(profile.upstream_key, "unknown"),
            ),
        )
        source_label: str = f"{profile.workbook}!{profile.query}"
        if source_label not in spec.declared_by:
            spec.declared_by.append(source_label)

        existing: Dict[str, ColumnSpec] = {column.name: column for column in spec.columns}
        for name, m_type in profile.column_types:
            normalized: str = re.sub(r"\s+", " ", m_type).strip().rstrip(",")
            if name not in existing:
                column = ColumnSpec(
                    name=name, m_type=normalized, pandas_dtype=pandas_dtype_for(normalized),
                    declared_by=[source_label],
                )
                spec.columns.append(column)
                existing[name] = column
                continue
            column = existing[name]
            if source_label not in column.declared_by:
                column.declared_by.append(source_label)
            if normalized != column.m_type and normalized not in column.conflicting_types:
                column.conflicting_types.append(normalized)

    for spec in specs.values():
        spec.partial = len(spec.columns) < min_columns_for_complete

    return sorted(specs.values(), key=lambda spec: (-len(spec.columns), spec.label))
