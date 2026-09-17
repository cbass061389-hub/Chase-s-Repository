# REVO Production System

VBA modules for `Revo_Production.xlsm` — the release flow, coded quality
capture, a quality dashboard, measured cart velocity, and an operating
scorecard that says what to do rather than just what happened.

**Start here: [`docs/INSTALL.md`](docs/INSTALL.md).** Ten minutes.

---

## Why this exists

The release UserForm stopped appearing. The cause was not the form — it was
that the release macro wrote status text (`RELEASED`, `PARTIAL`, `NO DATA`)
into column W on REVO Floor, which is an Excel *cell-checkbox* column. A
checkbox cell holding text stops being a checkbox, permanently. Cart 34 still
has 48 units on it and can never be released again without hand-editing.

Three more defects sat behind that one, all in
[`docs/FINDINGS.md`](docs/FINDINGS.md).

Fixing it properly meant separating status from trigger, which meant touching
the write path, which was the moment to build the rest.

## What you get

| Module | Does |
|---|---|
| `REVO_Core` | SKU attribute parsing, safe conversion, header discovery, app state, robust statistics, audit trail |
| `REVO_Release` | Patched release flow. Never writes text to the checkbox column, never skips silently, cross-checks against Release Log |
| `REVO_Quality` | Quality Log event model, sheet-driven taxonomy, classifier, history backfill |
| `REVO_QualityDash` | Pareto by defect / root cause / op / family, movers, data hygiene |
| `REVO_CartVelocity` | Measures cart movement from Floor Log snapshots, rejects outliers, applies capacity within guardrails |
| `REVO_Scorecard` | Week status, the constraint, overtime / extra day / headcount, ranked focus list |
| `REVO_FormBuilder` | Builds `frmReleaseDetails` from code so the form is version-controlled |
| `REVO_Install` | One-shot install, daily driver, and the self test |

## Entry points

| Macro | When |
|---|---|
| `REVO_Install` | Once, first |
| `REVO_ReleaseCarts` | The REVO Floor button (re-pointed automatically) |
| `REVO_DailyUpdate` | Each plan update — velocity, quality, scorecard |
| `REVO_ReconcileFloor` | When the board and Release Log disagree |
| `REVO_ResetCart` | Free a cart stranded by the old macro |
| `REVO_RunSelfTest` | Any time you want the assumptions checked |
| `REVO_BuildReleaseForm` | After editing the form's code or the taxonomy |

## Design positions worth knowing

**The checkbox column is a trigger, not a status field.** Status lives in its
own `Release Status` column. The checkbox is reset to `FALSE` after a pass and
stays clickable forever.

**Release Log is the record of what shipped.** Where REVO Floor disagrees with
it, the log wins and the operator is told. Column Y was previously both a
formula and a VBA write target — two owners, so neither was right.

**Nothing is skipped silently.** Every row the release macro passes over is
reported with a reason. Silent skips are what made this look broken.

**Measured beats assumed, but only with enough samples.** Cart velocity applies
a measured rate to Cart Tracker only when it has `MIN_SAMPLE` clean intervals
*and* the rate is within `TOLERANCE` of the value already there. Outside that,
the human number stands and the proposal is written beside it. Every figure on
the scorecard says whether it is MEASURED or ASSUMED.

**Outliers are removed by median absolute deviation, not standard deviation.**
A cart held for nine days is exactly the observation that inflates a standard
deviation enough to hide itself. Carts on an open hold are dropped outright —
they are stopped, not slow.

**The taxonomy is data, not code.** Defect types, locations, root causes, ops
and corrective actions live on the `Quality_Codes` sheet. The form reads them
every time it opens. Adding a defect type is a sheet edit.

**Night shift cannot release.** The scorecard separates throughput hours from
release hours, so "add night overtime" is never recommended for a release-side
gap.

## Checking the source

```
python3 revo/tools/check_vba.py revo/vba
```

Five classes of compile error that do not need Excel to find: the
24-continuation statement limit, reserved-word collisions (VBA identifiers are
case-insensitive, so `dO` is `Do` and `eNum` is `Enum`), block balance, error
handlers that read `Err` after something cleared it or fall off without
`Resume`, and early-bound UserForm references in standard modules. Run it before
handing modules to anyone.

## Two things to know before you trust it

**This code has never been executed.** It was written without Excel in the
loop. `REVO_SelfTest` exists for exactly that reason: it walks every assumption
the code makes about your sheets and reports what it actually found. Run it
first and read it.

**Cart velocity has almost no data yet.** Floor Log holds 2 snapshots, so one
interval. The engine will measure, report `HELD — 1 clean sample, need 8`, and
leave your capacity inputs alone. That is the guardrail working, not a failure.
Run `LOG FLOOR` daily and it starts applying on its own after roughly two
weeks.
