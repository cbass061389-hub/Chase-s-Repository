"""Write the canonical layer: parquet for the app, SQLite for ad-hoc queries.

Both are regenerated from scratch each run and neither is hand-edited, so the
warehouse is a pure function of the exports plus `config.yaml`. Re-running never
leaves half-written state: each table is written to a temporary file and moved
into place.
"""

from __future__ import annotations

import shutil
import sqlite3
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence

import pandas as pd

SQLITE_FILENAME: str = "sc.sqlite"


@dataclass
class WriteReport:
    table: str
    rows: int
    parquet_path: str
    columns: int = 0


def attach_lineage(
    frame: pd.DataFrame,
    *,
    source_slug: str,
    source_origin: str,
    extracted_at: datetime,
    data_as_of: Optional[datetime] = None,
) -> pd.DataFrame:
    """Stamp lineage onto every row.

    ``data_as_of`` is the vintage of the data, which is not the same as when the
    reader ran. Where the source does not state its own as-of date it stays
    null, and the vintage layer shows it as unknown rather than implying the
    file's timestamp is the data's age.
    """
    stamped: pd.DataFrame = frame.copy()
    stamped["source_slug"] = source_slug
    stamped["source_origin"] = source_origin
    stamped["extracted_at"] = pd.Timestamp(extracted_at)
    stamped["data_as_of"] = pd.Timestamp(data_as_of) if data_as_of is not None else pd.NaT
    return stamped


def write_table(frame: pd.DataFrame, table: str, warehouse_dir: Path) -> WriteReport:
    """Write one table to parquet atomically."""
    warehouse_dir.mkdir(parents=True, exist_ok=True)
    target: Path = warehouse_dir / f"{table}.parquet"
    staging: Path = warehouse_dir / f".{table}.parquet.tmp"
    frame.to_parquet(staging, index=False)
    shutil.move(str(staging), str(target))
    return WriteReport(table=table, rows=int(len(frame)), parquet_path=str(target),
                       columns=int(frame.shape[1]))


def write_sqlite(tables: Dict[str, pd.DataFrame], warehouse_dir: Path) -> Path:
    """Mirror every table into SQLite so a number can be checked without Python."""
    warehouse_dir.mkdir(parents=True, exist_ok=True)
    target: Path = warehouse_dir / SQLITE_FILENAME
    staging: Path = warehouse_dir / f".{SQLITE_FILENAME}.tmp"
    if staging.exists():
        staging.unlink()

    with sqlite3.connect(staging) as connection:
        for name, frame in tables.items():
            # SQLite has no native datetime; ISO text keeps it sortable and readable.
            flattened: pd.DataFrame = frame.copy()
            for column in flattened.columns:
                if pd.api.types.is_datetime64_any_dtype(flattened[column]):
                    flattened[column] = flattened[column].dt.strftime("%Y-%m-%d %H:%M:%S")
            flattened.to_sql(name, connection, if_exists="replace", index=False)
        connection.commit()
    shutil.move(str(staging), str(target))
    return target
