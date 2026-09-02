# `discovery/` — generated, not hand-written

`python -m sc.cli discover` writes two files here:

- **`manifest.json`** — the full machine-readable record of the estate. Every
  later phase reads this, not the workbooks directly.
- **`DISCOVERY.md`** — the report to read. Conclusion first, then the refresh
  dependency graph, the sheet-level entity map, duplicate-truth pairs and probe
  failures.

Both are committed so the manifest is a version-controlled record of what the
estate looked like on a given day, and a diff between runs shows what moved.

**The committed copies cover only three uploaded workbooks and carry temporary
paths — see [`PROVENANCE.md`](PROVENANCE.md).** Re-run on your own machine to
replace them with canonical paths and full coverage.

See [`../docs/PHASE0_RUNBOOK.md`](../docs/PHASE0_RUNBOOK.md).
