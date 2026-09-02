"""Validation gates. A blocking failure stops a publish.

The gates encode what discovery proved can go wrong here: keys that do not
normalize, a grain that is not unique, a location nobody mapped, quantities
that parsed to nothing. Each produces an exception row carrying both the
offending value and its lineage, per rule 4 of the canonical model — conflicts
are data, not something to smooth over.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional, Sequence

import pandas as pd

SEVERITY_BLOCKING: str = "blocking"
SEVERITY_WARNING: str = "warning"


@dataclass
class GateResult:
    """One gate's verdict on one dataset."""

    gate: str
    entity: str
    severity: str
    passed: bool
    detail: str
    offending_rows: int = 0
    sample: List[Dict[str, Any]] = field(default_factory=list)

    @property
    def blocks_publish(self) -> bool:
        return not self.passed and self.severity == SEVERITY_BLOCKING


def _sample(frame: pd.DataFrame, columns: Sequence[str], limit: int = 5) -> List[Dict[str, Any]]:
    usable = [column for column in columns if column in frame.columns]
    if not usable or frame.empty:
        return []
    return frame[usable].head(limit).astype(object).where(frame[usable].notna(), None).to_dict("records")


def gate_min_rows(frame: pd.DataFrame, entity: str, minimum: int) -> GateResult:
    passed: bool = len(frame) >= minimum
    return GateResult(
        gate="min_rows", entity=entity, severity=SEVERITY_BLOCKING, passed=passed,
        detail=f"{len(frame):,} rows (minimum {minimum:,})"
               + ("" if passed else " — an empty extract usually means the export moved, not that supply is empty"),
        offending_rows=0 if passed else 1,
    )


def gate_no_null_keys(frame: pd.DataFrame, entity: str, key_column: str, max_ratio: float) -> GateResult:
    if key_column not in frame.columns:
        return GateResult(
            gate="no_null_keys", entity=entity, severity=SEVERITY_BLOCKING, passed=False,
            detail=f"key column '{key_column}' is absent from the frame",
        )
    nulls: "pd.Series[bool]" = frame[key_column].isna()
    count: int = int(nulls.sum())
    ratio: float = (count / len(frame)) if len(frame) else 0.0
    passed: bool = ratio <= max_ratio
    return GateResult(
        gate="no_null_keys", entity=entity, severity=SEVERITY_BLOCKING, passed=passed,
        detail=f"{count:,} of {len(frame):,} rows ({ratio:.2%}) have no {key_column} "
               f"(allowed {max_ratio:.2%})",
        offending_rows=count,
        sample=_sample(frame[nulls], [c for c in frame.columns[:6]]),
    )


def gate_unique_grain(frame: pd.DataFrame, entity: str, grain: Sequence[str]) -> GateResult:
    missing = [column for column in grain if column not in frame.columns]
    if missing:
        return GateResult(
            gate="unique_grain", entity=entity, severity=SEVERITY_BLOCKING, passed=False,
            detail=f"grain columns absent from the frame: {missing}",
        )
    duplicated: "pd.Series[bool]" = frame.duplicated(subset=list(grain), keep=False)
    count: int = int(duplicated.sum())
    return GateResult(
        gate="unique_grain", entity=entity, severity=SEVERITY_BLOCKING, passed=count == 0,
        detail=f"grain {' x '.join(grain)}: {count:,} row(s) share a key"
               + ("" if count == 0 else " — the declared grain is wrong, or the source double-counts"),
        offending_rows=count,
        sample=_sample(frame[duplicated].sort_values(list(grain)), list(grain) + ["qty_on_hand"]),
    )


def gate_known_locations(
    frame: pd.DataFrame, entity: str, location_column: str, attribute_column: str, max_ratio: float
) -> GateResult:
    if location_column not in frame.columns or attribute_column not in frame.columns:
        return GateResult(
            gate="known_locations", entity=entity, severity=SEVERITY_WARNING, passed=True,
            detail="location columns not present; gate not applicable",
        )
    unmapped: "pd.Series[bool]" = frame[attribute_column].isna() & frame[location_column].notna()
    count: int = int(unmapped.sum())
    ratio: float = (count / len(frame)) if len(frame) else 0.0
    names = sorted(frame.loc[unmapped, location_column].dropna().unique().tolist())
    return GateResult(
        gate="known_locations", entity=entity, severity=SEVERITY_BLOCKING, passed=ratio <= max_ratio,
        detail=f"{count:,} row(s) ({ratio:.2%}) sit at a location missing from "
               f"canonical.locations: {names[:8]} (allowed {max_ratio:.2%})",
        offending_rows=count,
        sample=_sample(frame[unmapped], [location_column, "sku", "qty_on_hand"]),
    )


def gate_non_negative(frame: pd.DataFrame, entity: str, column: str) -> GateResult:
    if column not in frame.columns:
        return GateResult(
            gate="non_negative", entity=entity, severity=SEVERITY_WARNING, passed=True,
            detail=f"{column} not present; gate not applicable",
        )
    negative: "pd.Series[bool]" = frame[column].fillna(0) < 0
    count: int = int(negative.sum())
    return GateResult(
        gate="non_negative", entity=entity, severity=SEVERITY_WARNING, passed=count == 0,
        detail=f"{count:,} row(s) have {column} < 0"
               + ("" if count == 0 else " — legitimate for in-transit netting, wrong for on-hand"),
        offending_rows=count,
        sample=_sample(frame[negative], ["sku", "location_id", column]),
    )


def to_exception_rows(results: Sequence[GateResult], source_slug: str) -> pd.DataFrame:
    """Failed gates as exception rows, ready for the exceptions table."""
    detected: str = datetime.now(timezone.utc).isoformat(timespec="seconds")
    records: List[Dict[str, Any]] = [
        {
            "exception_id": f"{source_slug}:{result.entity}:{result.gate}",
            "exception_type": result.gate,
            "severity": result.severity,
            "entity": result.entity,
            "source_slug": source_slug,
            "message": result.detail,
            "offending_rows": result.offending_rows,
            "sample": str(result.sample) if result.sample else "",
            "detected_at": detected,
        }
        for result in results
        if not result.passed
    ]
    return pd.DataFrame.from_records(records) if records else pd.DataFrame(
        columns=["exception_id", "exception_type", "severity", "entity", "source_slug",
                 "message", "offending_rows", "sample", "detected_at"]
    )
