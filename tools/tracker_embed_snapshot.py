"""Re-embed CM_Open_Order_Tracker.xlsx as the built-in snapshot of the HTML viewer.

Usage:
    python tools/tracker_embed_snapshot.py <tracker.html> <tracker.xlsx> <out.html> [label]

The viewer keeps the workbook as a base64 string on one line:
    <script>window.SNAPSHOT_NAME="...";window.SNAPSHOT_B64="...";</script>
This replaces that line so the page opens on the latest data without a manual load.
"""
from __future__ import annotations

import base64
import datetime as dt
import re
import sys

PAT = re.compile(r'<script>window\.SNAPSHOT_NAME="[^"]*";window\.SNAPSHOT_B64="[^"]*";</script>')


def embed(html_in: str, xlsx: str, html_out: str, label: str | None = None) -> None:
    html = open(html_in, encoding="utf-8").read()
    b64 = base64.b64encode(open(xlsx, "rb").read()).decode("ascii")
    label = label or f"CM_Open_Order_Tracker.xlsx (snapshot {dt.date.today():%-m/%-d/%Y})"
    new = f'<script>window.SNAPSHOT_NAME="{label}";window.SNAPSHOT_B64="{b64}";</script>'
    html, n = PAT.subn(lambda _: new, html, count=1)
    if n != 1:
        raise SystemExit("snapshot <script> line not found in HTML")
    open(html_out, "w", encoding="utf-8").write(html)
    print(f"embedded {xlsx} ({len(b64)//1024} KB base64) as '{label}' -> {html_out}")


if __name__ == "__main__":
    embed(*sys.argv[1:5])
