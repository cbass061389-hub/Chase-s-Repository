"""CSV probe for the NetSuite File Cabinet exports.

These land as flat exports and are the true upstream for open PO, SO and item
data, so they need the same treatment as a workbook: real header names, a real
row count, and the encoding/delimiter actually on disk rather than assumed.
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

# Ordered by likelihood for NetSuite exports: UTF-8 with BOM is the default.
_ENCODINGS: tuple[str, ...] = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
_DELIMITERS: str = ",;\t|"


@dataclass
class CsvProbe:
    path: str
    encoding: str = ""
    delimiter: str = ""
    headers: List[str] = field(default_factory=list)
    row_count: int = 0
    row_count_capped: bool = False
    ragged_rows: int = 0
    error: Optional[str] = None

    @property
    def column_count(self) -> int:
        return len(self.headers)


def probe_csv(path: Path | str, sniff_bytes: int, row_cap: int = 2_000_000) -> CsvProbe:
    """Detect encoding, delimiter and headers, then count rows in one pass."""
    target: Path = Path(path)
    probe: CsvProbe = CsvProbe(path=str(target))

    sample: bytes = b""
    try:
        with target.open("rb") as handle:
            sample = handle.read(sniff_bytes)
    except OSError as exc:
        probe.error = f"unreadable: {exc}"
        return probe

    if not sample.strip():
        probe.error = "file is empty"
        return probe

    text: str = ""
    for encoding in _ENCODINGS:
        try:
            text = sample.decode(encoding)
            probe.encoding = encoding
            break
        except UnicodeDecodeError:
            continue
    if not probe.encoding:
        probe.error = f"no encoding in {_ENCODINGS} decoded the first {sniff_bytes} bytes"
        return probe

    try:
        dialect: type[csv.Dialect] | csv.Dialect = csv.Sniffer().sniff(text, delimiters=_DELIMITERS)
        probe.delimiter = dialect.delimiter
    except csv.Error:
        # Fall back to the delimiter with the highest count on the first line.
        first_line: str = text.splitlines()[0] if text.splitlines() else ""
        probe.delimiter = max(_DELIMITERS, key=first_line.count) if first_line else ","

    try:
        with target.open("r", encoding=probe.encoding, newline="") as handle:
            reader = csv.reader(handle, delimiter=probe.delimiter)
            probe.headers = [h.strip() for h in next(reader, [])]
            expected: int = len(probe.headers)
            for row in reader:
                probe.row_count += 1
                if expected and len(row) != expected:
                    probe.ragged_rows += 1
                if probe.row_count >= row_cap:
                    probe.row_count_capped = True
                    break
    except (OSError, csv.Error, UnicodeDecodeError) as exc:
        probe.error = f"row scan failed after {probe.row_count} rows: {exc}"

    return probe
