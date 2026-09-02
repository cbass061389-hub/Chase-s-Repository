"""Persist and load the recovered export schemas.

The schema file is generated from the committed M source by
``python -m sc.cli schemas``. It is checked in, so the readers have a contract
to validate against even on a machine with no access to the exports, and a
change to an export's shape shows up as a diff.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence

import yaml


@dataclass
class SourceColumn:
    name: str
    dtype: str
    m_type: str = ""
    type_conflict: List[str] = field(default_factory=list)


@dataclass
class SourceSchema:
    """One upstream export and the columns its readers expect."""

    slug: str
    key: str
    label: str
    kind: str
    columns: List[SourceColumn] = field(default_factory=list)
    declared_by: List[str] = field(default_factory=list)
    partial: bool = False
    entity: str = ""
    notes: str = ""

    @property
    def column_names(self) -> List[str]:
        return [column.name for column in self.columns]

    @property
    def dtypes(self) -> Dict[str, str]:
        return {column.name: column.dtype for column in self.columns}


def dump_schemas(schemas: Sequence[SourceSchema], path: Path) -> Path:
    """Write the schema file. Deterministic ordering so diffs stay readable."""
    payload: Dict[str, Any] = {
        "generated_from": "committed Power Query M source (queries/*/_Section1.m)",
        "note": (
            "Recovered from each query's Table.TransformColumnTypes step. Regenerate with "
            "`python -m sc.cli schemas`. Edit `entity` and `notes` by hand; they are preserved "
            "across regeneration."
        ),
        "sources": [
            {
                "slug": schema.slug,
                "key": schema.key,
                "label": schema.label,
                "kind": schema.kind,
                "entity": schema.entity,
                "notes": schema.notes,
                "partial": schema.partial,
                "declared_by": schema.declared_by,
                "columns": [
                    {
                        "name": column.name,
                        "dtype": column.dtype,
                        "m_type": column.m_type,
                        **({"type_conflict": column.type_conflict} if column.type_conflict else {}),
                    }
                    for column in schema.columns
                ],
            }
            for schema in schemas
        ],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(payload, sort_keys=False, width=100), encoding="utf-8")
    return path


def load_schemas(path: Path) -> List[SourceSchema]:
    """Read the schema file. Raises if it is missing — never guesses a schema."""
    if not path.is_file():
        raise FileNotFoundError(
            f"{path} not found. Run `python -m sc.cli schemas` to generate it from the "
            "committed M source."
        )
    raw: Any = yaml.safe_load(path.read_text(encoding="utf-8"))
    if not isinstance(raw, dict) or "sources" not in raw:
        raise ValueError(f"{path} is not a schema file (no 'sources' key)")

    schemas: List[SourceSchema] = []
    for entry in raw["sources"]:
        schemas.append(
            SourceSchema(
                slug=str(entry["slug"]),
                key=str(entry.get("key", "")),
                label=str(entry.get("label", "")),
                kind=str(entry.get("kind", "unknown")),
                entity=str(entry.get("entity") or ""),
                notes=str(entry.get("notes") or ""),
                partial=bool(entry.get("partial", False)),
                declared_by=[str(d) for d in entry.get("declared_by") or []],
                columns=[
                    SourceColumn(
                        name=str(column["name"]),
                        dtype=str(column.get("dtype", "string")),
                        m_type=str(column.get("m_type", "")),
                        type_conflict=[str(t) for t in column.get("type_conflict") or []],
                    )
                    for column in entry.get("columns") or []
                ],
            )
        )
    return schemas


def preserved_annotations(path: Path) -> Dict[str, Dict[str, str]]:
    """Hand-written ``entity``/``notes`` from an existing schema file, by slug.

    Regeneration must not discard the mapping decisions a human recorded, so
    those fields are carried across.
    """
    if not path.is_file():
        return {}
    try:
        raw: Any = yaml.safe_load(path.read_text(encoding="utf-8"))
    except yaml.YAMLError:
        return {}
    if not isinstance(raw, dict):
        return {}
    return {
        str(entry["slug"]): {
            "entity": str(entry.get("entity") or ""),
            "notes": str(entry.get("notes") or ""),
        }
        for entry in raw.get("sources") or []
        if isinstance(entry, dict) and entry.get("slug")
    }
