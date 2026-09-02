# Phase 0 runbook — run this on your Windows machine

Phase 0 could not run in the session that built it: the engine was written in a
Linux container with no access to `C:\Users\CharlesBass\OneDrive - Predator Group\`.
Everything here is machine-independent and tested; the sweep itself has to run
where the files are.

## One-time setup

```powershell
git clone <this repo>
cd Chase-s-Repository
python -m pip install PyYAML
```

That is the whole install. The discovery engine is standard library plus PyYAML
on purpose — it opens workbooks as zip archives rather than through Excel or
openpyxl, so it does not need either and does not choke on the large sheets.

## Run the sweep

```powershell
python -m sc.cli discover
```

Roots come from `sc/config.yaml`. `${USERPROFILE}` is expanded, and
`auto_detect_onedrive: true` also picks up any `OneDrive*` folder under your
profile, so a renamed tenant folder does not silently produce an empty manifest.

Add a prototype or repo folder without editing config:

```powershell
python -m sc.cli discover --root "C:\Users\CharlesBass\OneDrive - Predator Group" --root "D:\Prototypes"
```

Other flags: `-v` (debug logging), `--no-queries` (skip writing `queries/*.m`),
`--json-only` (manifest without the markdown report).

## What it writes

| Path | What it is |
|---|---|
| `discovery/manifest.json` | full machine-readable record; the contract for every later phase |
| `discovery/DISCOVERY.md` | the report to read — conclusion first, then the ranked table |
| `queries/<workbook>__<id>/*.m` | every Power Query decoded to text, one file per query, plus `_Section1.m` |
| `runs/<timestamp>/discover.log` | structured log |
| `runs/<timestamp>/manifest.json` | per-source row counts and timings for that run |

**Nothing under a configured root is written to.** Workbooks are opened
read-only as zip archives, macros are never executed, and Excel is never
launched. Writes go only to `discovery/`, `queries/` and `runs/`.

## Exit codes

| Code | Meaning | What to do |
|---|---|---|
| 0 | clean sweep | read `discovery/DISCOVERY.md` |
| 2 | blocked | a root was unreachable, nothing was found, or a file could not be probed — the manifest is not a full picture; fix and re-run |
| 1 | error | config is broken; the log names the key |

Exit 2 is deliberately loud. A discovery run that quietly finds nothing is worse
than one that fails, because it looks like a clean estate.

## Verify the engine before you trust it

```powershell
python -m unittest discover -s tests -t .
```

50 tests, no network, no fixtures on disk — they build real OOXML and BIFF12
packages in a temp directory and probe them. If these pass, the parsing paths
are sound.

## What Phase 0 cannot tell you

- **Whether a number is right.** It reports grain, keys and lineage; it does not
  read values or reconcile them.
- **What a formula means.** Formula counts are recorded as an is-it-a-source
  signal. The logic itself is not parsed.
- **What VBA does.** Component names only, from the plain-text `PROJECT` stream.
  A password-protected project cannot be read at all — that shows up as a
  critical-risk finding, because it is one.
- **`data_as_of` per source.** Nothing in a file reliably states its own data
  vintage. Until you tell me where each source records it, the vintage layer
  falls back to file modified time, which is not the same thing.

## Then

1. Read `discovery/DISCOVERY.md`, top section first.
2. Answer the seven questions at the end of `SCHEMA.md`.
3. Send me the manifest. Phase 2 readers get written against confirmed sources
   only.
