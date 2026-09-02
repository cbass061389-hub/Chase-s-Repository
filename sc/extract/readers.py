"""Parse a fetched export into a validated frame.

Three rules, all of them consequences of what discovery found:

* **Quoting is always honoured.** Two queries read the same export with
  ``QuoteStyle.None``, which ignores CSV quoting and shifts every value after an
  embedded comma. Nothing here can be configured into that mistake.
* **Headers are validated against the recovered schema, by name.** A missing or
  renamed column fails loudly with the column named, rather than shifting data
  into the wrong field.
* **Dtypes are explicit.** Everything is read as text first, then coerced per
  the declared type, so pandas never infers a type and quietly changes a value.
"""

from __future__ import annotations

import io
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import pandas as pd

from ..configuration import CanonicalRules
from .fetch import Fetched
from .normalize import coerce_numeric
from .schema_store import SourceSchema


class ReadError(RuntimeError):
    """Raised when an export cannot be read into its declared schema."""


@dataclass
class ReadReport:
    """What happened during a read. Attached to the frame's lineage."""

    slug: str
    origin: str
    rows: int = 0
    columns: int = 0
    encoding: str = ""
    missing_columns: List[str] = field(default_factory=list)
    unexpected_columns: List[str] = field(default_factory=list)
    coercion_failures: Dict[str, int] = field(default_factory=dict)

    @property
    def ok(self) -> bool:
        return not self.missing_columns


def _decode(payload: bytes, preferred: str) -> tuple[str, str]:
    """Decode bytes to text, reporting which encoding worked."""
    for encoding in (preferred, "utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return payload.decode(encoding), encoding
        except (UnicodeDecodeError, LookupError):
            continue
    raise ReadError("no supported encoding decoded the export")


def read_export(
    fetched: Fetched,
    schema: SourceSchema,
    rules: CanonicalRules,
    strict_columns: bool = True,
) -> tuple[pd.DataFrame, ReadReport]:
    """Read one export into a frame typed by its declared schema."""
    text, encoding = _decode(fetched.payload, str(rules.csv.get("encoding", "utf-8-sig")))
    report: ReadReport = ReadReport(slug=schema.slug, origin=fetched.origin, encoding=encoding)

    try:
        # dtype=str across the board: coercion is explicit below, never inferred.
        # quoting defaults to QUOTE_MINIMAL, which is the point — the estate's
        # QuoteStyle.None variants are what corrupted rows.
        frame: pd.DataFrame = pd.read_csv(
            io.StringIO(text),
            dtype=str,
            keep_default_na=False,
            na_values=[""],
            skip_blank_lines=True,
        )
    except (pd.errors.ParserError, ValueError) as exc:
        raise ReadError(f"{schema.slug}: CSV parse failed from {fetched.origin} ({exc})") from exc

    frame.columns = [str(column).strip() for column in frame.columns]
    declared: List[str] = schema.column_names
    present: set[str] = set(frame.columns)

    report.missing_columns = [name for name in declared if name not in present]
    report.unexpected_columns = [name for name in frame.columns if name not in set(declared)]
    report.rows = int(len(frame))
    report.columns = int(frame.shape[1])

    if report.missing_columns and strict_columns:
        raise ReadError(
            f"{schema.slug}: export at {fetched.origin} is missing declared column(s) "
            f"{report.missing_columns}. Either the export changed shape or the schema is "
            f"stale — regenerate with `python -m sc.cli schemas`. Columns present: "
            f"{sorted(present)[:15]}"
        )

    thousands: str = str(rules.csv.get("thousands_separator", ","))
    for column in schema.columns:
        if column.name not in frame.columns:
            continue
        series: "pd.Series[Any]" = frame[column.name]
        if column.dtype in ("Int64", "Float64"):
            coerced = coerce_numeric(series, thousands)
            report.coercion_failures[column.name] = int(
                (coerced.isna() & series.notna()).sum()
            )
            frame[column.name] = (
                coerced.round().astype("Int64") if column.dtype == "Int64" else coerced.astype("Float64")
            )
        elif column.dtype.startswith("datetime"):
            parsed = pd.to_datetime(series, errors="coerce", format="mixed")
            report.coercion_failures[column.name] = int((parsed.isna() & series.notna()).sum())
            frame[column.name] = parsed
        elif column.dtype == "boolean":
            lowered = series.astype("string").str.strip().str.lower()
            frame[column.name] = lowered.map(
                {"true": True, "t": True, "yes": True, "y": True, "1": True,
                 "false": False, "f": False, "no": False, "n": False, "0": False}
            ).astype("boolean")
        else:
            frame[column.name] = series.astype("string").str.strip()

    report.coercion_failures = {k: v for k, v in report.coercion_failures.items() if v}
    return frame, report
