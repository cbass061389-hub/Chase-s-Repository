# `discovery/` — generated, not hand-written

`python -m sc.cli discover` writes two files here:

- **`manifest.json`** — the full machine-readable record of the estate. Every
  later phase reads this, not the workbooks directly.
- **`DISCOVERY.md`** — the report to read. Conclusion first, then the ranked
  source table, duplicate-truth pairs, dependency map and probe failures.

Both are empty in a fresh clone because they describe **your** OneDrive tree,
and a sweep has to run where the files are. Run it and commit the results — the
manifest is the version-controlled record of what the estate looked like on a
given day, and a diff between runs shows what moved.

See [`../docs/PHASE0_RUNBOOK.md`](../docs/PHASE0_RUNBOOK.md).
