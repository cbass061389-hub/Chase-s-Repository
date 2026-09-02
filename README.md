# Supply Chain Reference & Update Engine — Predator Group

**Phase 0 is built and runnable. Phases 2-6 are deliberately gated.**

Phase 0 sweeps the workbook estate read-only and produces the manifest plus the
canonical model proposal. The build spec's own rules of engagement say stop
there and agree on the model before extraction code exists, so that is where
this stops.

Run it: **[`docs/PHASE0_RUNBOOK.md`](docs/PHASE0_RUNBOOK.md)**
Review it: **[`SCHEMA.md`](SCHEMA.md)** — the findings and the eight questions are the ask.

```powershell
python -m pip install -r requirements.txt
python -m sc.cli discover      # sweep the estate, extract the query stack
python -m sc.cli schemas       # recover each export's schema from the M
python -m sc.cli reconcile     # where the forked query logic diverges
python -m sc.cli formulas      # what the calculation engines compute
python -m sc.cli extract       # land the canonical layer
```

> **Read `SCHEMA.md` §5 first.** `Position Engine` column P — Monthly Run Rate,
> the denominator of months-of-supply — is `#REF!` in all 2,213 rows, and
> `__CleanBO` loses 31% of its back-order rows the same way. Both sheets are
> hidden, so nothing surfaces it.

## What Phase 0 gives you

| | |
|---|---|
| **Every source found and described** | path, size, vintage, sheets (visible and hidden), row counts, named tables, defined ranges, connections |
| **The Power Query stack as text** | DataMashup blob decoded, `Section1.m` split into one `.m` file per query under `queries/`, diffed against the previous run so a query edit that breaks the pipeline shows up in git |
| **Source vs downstream copy** | inferred from formula density, query sources and external links — not from the filename |
| **Grain and keys per source** | validated header detection (no assumption that row 1 is the header), then supply-chain key detection: SKU, PO, container, invoice, POL/POD, location, UOM |
| **What breaks if a path moves** | every external location, ranked — a query reading from `Downloads` scores worse than one reading from SharePoint |
| **Duplicate truth** | two workbooks holding the same columns with different row counts, matched on header signature so it catches pairs whose names look nothing alike |
| **Expected assets confirmed, not assumed** | each workbook from the spec is verified present; 17 seeds, and misses are named with the patterns searched |
| **VBA inventory** | component names, never executed; a password-protected project is reported as a blocker |

## Design decisions worth knowing

**Zip-level parsing, not Excel or openpyxl.** Workbooks are opened as zip
archives and their XML streamed. The estate has multi-hundred-megabyte sheets
that defeat DOM parsers, the tool has to run without installs, and nothing that
never opens Excel can accidentally save over a source.

**`.xlsb` is parsed properly.** The Overstock workbook is `.xlsb`, so a BIFF12
record reader is included rather than leaving a named source as a blind spot.

**Header detection is scored, not assumed.** Real workbooks have title blocks
above the header band. The probe scans the first 25 rows, picks the header by
width, text ratio and distinctness, and reports its confidence — low confidence
is itself a risk finding.

**Failures are loud.** An unreachable root or an unprobeable file exits non-zero.
A discovery run that quietly finds nothing looks like a clean estate, which is
the worst possible outcome.

**Missing is missing.** Nothing is interpolated, zero-filled or carried forward,
here or in the proposed model.

## Layout

```
sc/
  config.yaml          every path, threshold, keyword and seed — no literals in code
  configuration.py     strict typed loader; a bad key fails immediately
  models.py            Phase 1 canonical entities (proposed)
  cli.py               `discover` (built) and `run` (gated)
  discovery/
    scan.py            orchestrator: sweep -> probe -> classify -> manifest
    ooxml.py           zip/relationship/shared-string plumbing
    workbook_probe.py  sheets, rows, headers, tables, connections, links
    xlsb.py            BIFF12 reader
    datamashup.py      DataMashup -> Section1.m -> queries + source paths
    vba.py             VBA component inventory (never executed)
    csv_probe.py       encoding, delimiter, headers, ragged rows
    classify.py        domain, role, grain, risk, duplicate truth
    report.py          manifest.json + DISCOVERY.md
  analyze/             query reconciliation, schema recovery, formula mapping
  extract/             canonical readers, normalization, gates, warehouse
tests/                 122 tests, stdlib unittest, real packages built in temp dirs
docs/PHASE0_RUNBOOK.md
SCHEMA.md              canonical model proposal
discovery/             generated: manifest.json + DISCOVERY.md (committed)
queries/               generated: extracted M source (committed)
```

## What the analysis found

Run against the three real workbooks, the tools found these without anyone
reading a formula by hand:

| | |
|---|---|
| **A circular refresh dependency** | Meeting and HIE each refresh from the other, so refresh order decides the numbers |
| **One export, three transformations** | 5 of 8 NetSuite exports feed 2-3 workbooks, each with its own logic. The reports don't disagree about data — they disagree about logic |
| **`"Available"` means two things** | `List.Sum([On Hand])` in two queries, `List.Sum([Available])` in a third, all emitted under the same column name |
| **A join key built four ways** | Only one of four SKU derivations trimmed whitespace, so the same item produced two key strings |
| **`QuoteStyle.None`** | Two queries read a quoted CSV with quoting disabled, which shifts every value after an embedded comma |
| **2,213 + 11,636 broken formula cells** | `Position Engine!P` is 100% `#REF!`; `__CleanBO` A-D are 31% |
| **A superseded file still in use** | The meeting workbook's `HIEInv` reads an older HIE workbook than its three sibling queries |

`discovery/RECONCILIATION.md` and `discovery/FORMULA_MAP.md` have the detail.
Each finding names the two queries, the values, and the consequence.

## Honest status

**Built and tested:** the discovery engine, the query reconciler, export-schema
recovery, the formula mapper, the canonical extraction layer, and the config
layer that drives all of it. 122 tests.

**Tested against what:** real OOXML and BIFF12 packages assembled byte-for-byte
in the test suite — a genuine base64 DataMashup blob wrapping a real nested zip,
real BIFF12 record framing, a real `PROJECT` stream. **Not** against your actual
workbooks, which were not reachable from the environment this was written in.
Expect layout quirks in the real files — merged header bands, stray totals rows,
columns renamed since a query was written. The probes report low confidence and
structural errors rather than guessing, so those surface as findings on the
first run instead of as wrong numbers later.

**Built for 3 of 8 exports:** `inventory_onhand`, `open_po` and `allocation`
have canonical builders, gates and tests. The other five are declared with the
reason they are not modelled yet — each names the blocking decision in
`SCHEMA.md` rather than guessing a grain. `sc/extract/entities.py::UNMODELLED`
is the list.

**Not built, on purpose:** Phase 3 refresh orchestration, `SC_Reference.xlsx`,
the HTML engine. `python -m sc.cli run` exits 2 and says why.

**Where I disagree with the spec:** Phase 3 proposes an Excel COM driver to
refresh Power Query. It works, and it is the wrong place to end up — a headless
scheduled task that depends on a GUI Excel instance is the most fragile
component in any stack that has one, and it will be the thing that breaks at
6am. The discovery output already ranks every query by fragility for exactly
this reason. My recommendation is to treat COM as a bridge, port the fragile
queries to Python in `transform/`, and reduce Power Query to plain extraction.
Full argument once the manifest shows how many queries are actually load-bearing.
