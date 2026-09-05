# REVO_Production.xlsm — lean-out and rebuild

Takes the 13 MB / 65-sheet `Revo_Production.xlsm` to a 29-sheet
`REVO_Operations.xlsm` with the data, history, macro buttons and VBA project
intact, plus a quality-tracking layer and a weekly-update email builder.

**Nothing here modifies the original.** The driver copies it, hashes it before
and after the copy, and works only on the copy.

---

## Run it

On the Windows machine with desktop Excel:

```powershell
cd revo\driver
.\Invoke-RevoMigration.ps1 -Source 'C:\REVO\Revo_Production.xlsm' -OutDir 'C:\REVO\out'
```

Phases run in order and each takes its own backup first, so any one can be
rolled back without unwinding the others:

| Phase | What it does |
|---|---|
| `baseline` | Copy, per-sheet used-range and shape inventory, no mutation |
| `lean` | Import `modConfig` + `modMigrate`, run `MigrateAll` |
| `layers` | Import the other three modules, build quality sheets, input sheet and `frmDisposition` |
| `verify` | Re-open, confirm the 29 sheets, size and log, no mutation |

Run one at a time with `-Phase lean`. Output lands in `-OutDir`:
`REVO_Operations.xlsm`, `REVO_Operations.before-<phase>.xlsm`,
`inventory-<phase>.csv`, `migration-<timestamp>.log`, and
`REVO_Migration_Log.txt` beside the workbook — that last one is the audit trail
the VBA itself writes.

**Before the `layers` phase**, once, by hand:
File → Options → Trust Center → Trust Center Settings → Macro Settings →
**Trust access to the VBA project object model**. The driver checks the registry
and stops with instructions if it is off.

---

## Phase 1 findings

The pre-flight scan ran against the live file. Three things in the original
brief were wrong.

**The sheet split is clean.** Every `<f>` formula cell on all 29 retained sheets
was scanned against all 36 cut sheet names, quoted and unquoted. Zero
references. All 36 cut names and all 29 keep names exist, 36 + 29 = 65, no
orphans. `MigrateAll` re-proves this after the fact with `AuditRefErrors`.

**Correction — Release Log has no formulas.** Not 1,048,552 rows of live
formulas. It has **one** formula in the whole sheet. The 96 MB is 1,048,552 row
elements each carrying a single `FALSE` in column I ("Recalled") — a checkbox
column dragged to the bottom of the grid. Real data ends at **row 765**. There
was never a recalc problem here, only a size problem.

**Correction — STOCK ROOM COUNT is 1,736 cells and 1,048,564 empty hidden rows.**
51 MB of `<row hidden="1"/>`.

**Correction — 12 chart parts, not 3.** 22 sheets carry drawings or controls,
which matches. Command Center and Plan Config carry their macro buttons as
legacy VML form controls rather than DrawingML — invisible to most tooling, and
the reason the no-Python-writes rule is not negotiable.

**New — Power Query.** Nine connections plus a Data Model. Eight load into cut
sheets. Per Chase's call these are deleted outright, connection *and* query
definition. `Query - REVO Sales Forecast` is explicitly protected: it lands on
`REVO_SALES_FRCST`, which is retained.

**New — `Weekly_Update_Input` did not exist and nothing built it.**
`modWeeklyUpdate` called `RequireSheet` on it and would have failed on the first
line of every entry point. `BuildWeeklyInputSheet` now creates it.

---

## Release Log semantics — verified, not assumed

```
A Date | B SKU | C Cart | D Qty Released | E Work Order
F Qty Rework | G Qty B Grade | H Qty Reject
I Recalled | J Cart Qty Released | K Cart Remaining
```

Columns A–H match what `modWeeklyUpdate` assumed, so no constants moved. But
**D is net**. On every row where J is populated, `D + F + G + H = J` exactly
(row 758: 94 + 0 + 2 + 5 = 101; row 762: 52 + 0 + 5 + 3 = 60). So:

- Total processed off the cart = `D + F + G + H`
- A-Grade released to inventory = **`D`**

The first draft computed A-Grade as `Released − B-Grade`, subtracting the
B-Grade quantity a second time and understating A-Grade by exactly that amount
in the executive email. Fixed.

Reject rate now divides by total processed, which is what the number means on
the floor. `{REJECTRATE_REL}` exposes the old released-only basis if the
historical trend has to stay comparable — **see the open questions.**

J is only populated on recent rows (11 of 764), so it is not usable as a
historical total. The metrics recompute it from D+F+G+H instead.

---

## What changed in the five modules

None of them had ever been executed. These are the defects found and fixed.

### modConfig
- Added `SH_WEEKLYINPUT` and `SH_COMMAND`; the module's own rule is that every
  sheet name lives here, and `Weekly_Update_Input` was private to another module.
- Added `AddSheetAfter()`. Three build routines passed `GetSheet(...)` straight
  into `Worksheets.Add(After:=)`, which raises 1004 whenever the anchor sheet
  does not exist yet — which is exactly the state on a first run.
- Added `FastOn` / `FastOff`, which **save and restore** calculation, screen,
  events and alerts. The original hard-set `xlCalculationAutomatic` on exit from
  every routine, forcing a full recalc between each migration phase.
- Added `SilentMode`, `Say()`, `Confirm()`, `SetSilent()`. A `MsgBox` in an
  unattended COM run hangs Excel forever with no window to click and the driver
  times out holding a file lock.
- Added `LogLine()` → `REVO_Migration_Log.txt`.

### modConfig — one blocking collision found

`Update_Analysis` exports `Public Function SheetExists` with a different
signature. Two public procedures of the same name in different standard modules
make every unqualified call an **"Ambiguous name detected"** compile error, so
importing `modConfig` as written would have failed the whole `layers` phase.
modConfig's version is renamed `HasSheet`. A full public-name scan across the
new modules (62 exported names) against every existing module confirms this was
the only collision.

### modMigrate
- **`ResetAllUsedRanges` would have deleted macro buttons.** It deleted every row
  below and column right of the last used cell. On REVO Floor, Command Center,
  Plan Config and Rework Tracker the buttons are anchored below or right of the
  data. `ShapeFloor()` now computes the lowest row and rightmost column occupied
  by any shape and refuses to cut above it.
- Added `SnapshotShapes` / `VerifyShapes` — records every shape on every retained
  sheet before the deletes and proves afterwards that none were lost. This is the
  acceptance criterion, now evidenced rather than eyeballed.
- Added `DeleteCutQueries` — the 8 orphaned connections *and* their `WorkbookQuery`
  definitions. Deleting only the connection leaves the query in the Queries pane.
- Added `AuditRefErrors` — post-delete `#REF!` sweep across every retained sheet.
- Added `CleanBrokenNames` — 19 defined names referenced cut sheets
  (`_FilterDatabase`, `ExternalData_n`, a `Print_Area`). Workbook-scoped ones
  survive the sheet delete as `#REF!` and are swept.
- `TruncateReleaseLog` now uses `End(xlUp)` with the 1,048,576-row array scan as a
  fallback rather than as the default.
- `MigrateAll` holds one calc-manual window across all phases instead of thrashing
  between them, and never saves on failure.

### modQuality
- `Worksheets.Add(After:=GetSheet(...))` in three places → `AddSheetAfter`.
- **A blank `Active` cell silently hid a defect code from the operator's combo
  box.** `Empty <> False` evaluates to `False` in VBA, so any code added without
  explicitly ticking Active disappeared from the form with no error. Blank now
  means active; only an explicit FALSE deactivates.
- `BuildAnalysisShell` wiped the user's From/To dates on every rebuild. Preserved.
- Dictionary accumulation now uses `.Exists` / `.Add` instead of reading a
  missing key to create it.
- Added TRUE/FALSE validation on `Active`, and list validation on Cause Category
  and Typical Op Caused, so master data cannot drift into free text.
- `Rows(1).AutoFilter` guarded — it toggles the filter *off* if one already exists.

### modWeeklyUpdate
- Added `BuildWeeklyInputSheet` — the sheet the module depends on and nothing
  created. Builds the email block, the metric block, the narrative grid seeded
  with all nine sections in order, and a token cheat-sheet.
- Fixed the A-Grade double-subtraction and the reject-rate denominator (above).
- Added `{PROCESSED}` and `{REJECTRATE_REL}` tokens.
- `RefreshWeeklyMetrics` now raises if Release Log is missing instead of
  silently reporting zeros as if the week had no production.
- `BuildWeeklyUpdate` refuses to build while the To line is still a placeholder.
- Signature moved out of the code and onto the input sheet.
- `Substitute` renamed `ApplyTokens` — shadowing a worksheet-function name in a
  module this size is a trap for whoever maintains it next.
- **Still never calls `.Send`.** Verified: `.Display` only, no `.Send` anywhere
  in the module or the generated form code.

### modFormBuilder
- Cosmetic UserForm properties (Width/Height/BackColor) are set through the
  `Properties` collection, which throws on some Excel builds. Each is now
  individually guarded so a failed BackColor no longer aborts the whole form.
- **The three combo boxes were free-text.** Op codes have to match
  `IC / ASSYM / 10 / … / FQC` exactly or the escape analysis buckets typos as
  separate operations and the whole where-caused-vs-where-caught table degrades.
  All three are now `fmStyleDropDownList` with `MatchRequired`.
- `cboDefect.Value` etc. are `& ""` guarded — a Null from an untouched combo
  otherwise raises on `CStr`.
- Added `ShowDispositionTest`, the Phase 3 smoke test: primes the form with a
  dummy cart, shows it, and reports what came back plus how many rows landed in
  `Reject_Log_V2`.

The injected form code was rendered out of the builder's string concatenation
and checked independently: 209 lines, all blocks balanced, every referenced
control is one the builder creates, all four external calls resolve, no `.Send`.

---

## Open questions for Chase

1. **Reject rate denominator.** Now `Reject / total processed`. The old code used
   `Reject / released`, which is what the emails going out have been quoting. On
   the last 764 rows that is 1,219 / 61,281 = 1.99% versus 1,219 / 56,136 = 2.17%.
   Switching bases breaks comparability with the trend already in front of the
   exec team. `{REJECTRATE_REL}` gives you the old basis if you want it — say
   which one the Summary paragraph should use.
2. **The nine recipients.** `Weekly_Update_Input!C9` ships with a placeholder and
   `BuildWeeklyUpdate` refuses to run until it is replaced. I do not have the
   addresses.
3. **`Production Table` vs `Production Table Backup`.** See Phase 4 below — cutting
   Production Table silently kills the Production Analysis builders. My
   recommendation is to retain Production Table and cut only the 10 MB Backup.
   The Backup is 41,120 rows of genuine archive data, not formatting bloat like
   the other two big sheets, so confirm you have it elsewhere before it goes.
4. **`frmDisposition` is not wired to anything.** It builds, opens, validates and
   writes to `Reject_Log_V2`, but no existing macro calls `Prime`. Hooking it into
   the live release flow means editing `REVO_Ops` — that is Phase 4, and I would
   rather see which release entry point you actually use than guess.
5. **Weeks remaining** (`{WEEKS}`) is a manual cell. There is no source in the
   workbook I can derive it from. Say where it comes from and I will calculate it.

---

## Phase 4 — measured, and far smaller than the brief assumed

The brief's "103 shared procedures" is a call-graph reachability count, not an
edit surface. The actual surface was measured by scanning both modules for every
hard-coded cut-sheet name:

| Module | Refs | Procedures with use-sites |
|---|---|---|
| `Tube_Dash` | 10 | deleted wholesale — 2,251 lines, no edits needed |
| `REVO_Ops` | 16 | 7 procedures, 14 use-sites |
| `Update_Analysis` | 6 | 10 procedures, 21 use-sites |

**35 use-sites across 17 procedures.** And the code is better than expected:
every single reference is already a named constant except one literal at
`Update_Analysis` L3394 (`ProdTable_NormalizeDates_FromCutoff`). `Update_Analysis`
also already guards with `SheetExists(...) Then Exit Sub` at the top of most
builders, so it degrades rather than crashes.

Several of these procedures are not "guard the reference" work — they are dead by
definition, because their whole job is to build a sheet that is being deleted:
`BuildBOMHierarchyView` (both modules), `Build_EST_REVO_INV`,
`BuildFlattenedBOMMatrix`, `BuildComponentHealth`,
`BuildComponentWeeklyInventoryAudit`, `BuildRevoProductionAllocator`. Those get
retired, not patched.

Realistically this is one working session against a live Excel, not the long tail
the brief describes — but it still needs the `lean` phase finished first.

### The one that needs your decision

**`Production Analysis` is retained. It is built from `Production Table`, which
is cut.** The formula dependency scan came back clean because this dependency
lives in VBA, not in formulas:

| Procedure | Reads | Writes |
|---|---|---|
| `BuildProdAnalysisBVP` | Production Table | Production Analysis |
| `BuildProdAnalysisWVP` | Production Table | Production Analysis |
| `Build8WeekView` | Production Table | Production Analysis |
| `AppendProductionTable` | Production Table + Backup | Production Table |
| `CleanProductionTableDuplicates` | Production Table | Production Table |
| `DedupHorizon_ProdTable` | Production Table | Production Table |

Because `Update_Analysis` already guards with `SheetExists`, deleting Production
Table does not throw — those builders just `Exit Sub`. **Production Analysis
silently stops updating.** Stale numbers on a retained reporting sheet with no
error is a worse failure than a crash.

Three ways out, in the order I would take them:

1. **Retain `Production Table`** (0.08 MB — it is the *Backup* at 10 MB that
   carries the weight) and cut only `Production Table Backup`. Keeps the whole
   Production Analysis chain working, costs almost nothing in file size.
   `AppendProductionTable` needs its backup-write path guarded, and that is a
   ten-line change.
2. Cut both and accept that Production Analysis is retired.
3. Cut both and rebuild the three analysis builders against Release Log instead.
   Most work, and Release Log may not carry the same forecast columns —
   I would not start here.

**Recommendation: option 1.** It is the only one that preserves a retained
sheet's function, and the size argument does not apply — the 10 MB is entirely
in the Backup.

---

## VBA inventory

Confirmed at 12,264 lines across 78 components — `REVO_Ops` 3,897,
`Update_Analysis` 3,457, `Tube_Dash` 2,251, `modRejectDashboard_OneStop` 922,
`EmailKarimReport` 386, three userforms, and 65 eight-line sheet stubs.
