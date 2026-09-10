# Phase 10 — Windows run evidence

**What this file is.** A record of what happened on the target machine, run by
run, kept because a run that produced no number still produced information and
because three of the first three runs failed for three different reasons.

**What it is not.** It is not a baseline, not an acceptance record and not a
Gate-B result. Nothing here may be quoted as performance evidence.

The target machine, for every run below: Windows 11 Pro · Excel 16.0 build
23026, 64-bit · Ryzen 7 7800X3D · 31.2 GB RAM · repository under
`C:\Users\pcd\OneDrive\Desktop\PCCM-GateB\mine`.

---

## Run 1 — PERF-SMALL — ABORTED IN SETUP

**Harness commit:** `4e1f109`

Stage A 351/351. Stage-B bootstrap PASS (68.440 s). Aborted during the
environment capture, before any timed operation.

```
System.Management.Automation.PropertyNotFoundException
The property 'Value' cannot be found on this object.
FullyQualifiedErrorId: PropertyNotFoundStrict
```

**Cause.** `(Get-Item Env:OneDriveCommercial -ErrorAction SilentlyContinue).Value`.
The variable does not exist on a machine with a consumer OneDrive and no work
account, so the pipeline emitted nothing, the parenthesised expression was
`$null`, and `$null.Value` under `Set-StrictMode -Version 2.0` is a terminating
error rather than a quiet `$null`.

**Corrected in `6e87fda`** — normalise with `@()`, validate the property, then
read it. The same round added the stage cursor, so a later failure names the
stage it was in, and the baseline-status machinery, so an aborted run cannot
look like evidence.

**Status: 0 valid warm medians. NOT a baseline.**

---

## Run 2 — PERF-SMALL — ABORTED IN SETUP

**Harness commit:** `6e87fda`

Stage A 351/351. Stage-B bootstrap PASS (68.350 s). Passed the release-identity
preflight, the bootstrap, the workbook open, and aborted in the environment
inventory.

```
stage      : setup
doing      : capturing the environment inventory
exception  : System.Management.Automation.PropertyNotFoundException
message    : The property 'builder_version' cannot be found on this object.
at line    : 724
statement  : $record.Add('builder_version', [string]$Manifest.builder_version)
```

**Cause.** `stage_b_manifest.json` is a projection of the model side of the
specification. It has never carried `builder_version` and cannot: Phase 10 Step 3
settled that the model version and the builder version are independent
authorities. The harness asked the wrong artifact.

**Corrected in `4ad035a`** — `BUILDER_VERSION` is projected into
`build/phase10_benchmark_plan.json` as `release_identity`, with the name of the
authority beside each of the three values, and the runner refuses a plan that
cannot identify its release before anything is built.

**Status: 0 valid warm medians. NOT a baseline.**

---

## Run 3 — PERF-SMALL — ABORTED BUILDING THE SCENARIO

**Harness commit:** `4ad035a`

Stage A 351/351. Stage-B bootstrap PASS. Passed the release-identity preflight,
the bootstrap, the workbook open and the environment inventory. Aborted while
constructing the scenario through the accepted production endpoints.

```
stage     : scenario
doing     : building the PERF-SMALL fixture through the accepted production endpoints
scenario  : PERF-SMALL
exception : System.Runtime.InteropServices.COMException
message   : Table features aren't available because the sheet is protected.
at line   : 313
valid warm medians : 0 of 11 planned runs
```

Session shut down cleanly.

**Status: 0 of 11 valid warm medians. NOT a baseline, NOT a partial baseline,
NOT a performance sample.**

### What Run 3 established, and it is not only about the harness

Line 313 is `$victim.Delete()` in `Remove-TableRow` — the harness deleting a
`ListRow` from `tblFXRates` through COM. So the immediate call was the
harness's.

But the run also established something the harness cannot be blamed for. Before
it died, the fixture had already written four Setup scalars to locked cells on
protected sheets, and those writes SUCCEEDED. So on this machine, with this
build:

| Operation on a protected sheet | Result |
|---|---|
| code writes a VALUE to a locked cell | **works** — `UserInterfaceOnly:=True` is honoured |
| code performs a LISTOBJECT STRUCTURAL operation | **refused** |

Those are different capabilities, and the delivered workbook needs both. The
audit of `pccm/src/vba` — pinned in
`tests/test_phase10_table_structure_under_protection.py` — finds fifteen
ListObject structural operations across five modules, reached by seven of the
user commands, two of them unconditionally:

- `PCCM_ApplyTimeline` adds year `ListColumns` to three grids. A fresh workbook
  has none, so this fires for **any** timeline.
- `PCCM_Calculate` resizes the five `_Calc` tables to the model on **every**
  run.

Those are the first two buttons a user presses. `modWorkbook.RestoreTable`,
which is `PCCM_Calculate`'s transactional rollback, is structural too — so a
command that failed part-way under protection could not be undone either.

**This is therefore not concluded as a harness defect.** It is very probably a
production protection defect, and the correction it implies edits accepted
modules and contradicts an accepted contract sentence. That correction is not
being made on a reading of the documentation:
`bootstrap/windows/phase10_protection_probe.ps1` asks Excel the remaining
question directly — whether the VBA caller is refused exactly as the COM caller
was — in one command and about a minute.

---

## Protection probe Run 1 — INCONCLUSIVE

**Probe commit:** `24015ef`

Stage A 351/351. Stage-B bootstrap PASS. The probe then raised before it asked
its question.

```
System.Management.Automation.PropertyNotFoundException
The property 'Count' cannot be found on this object.
FullyQualifiedErrorId: PropertyNotFoundStrict
Verdict emitted: INCONCLUSIVE - the probe session raised before it finished
```

Shutdown clean.

**Status: INCONCLUSIVE. It does not prove production is blocked. It does not
prove production is fine. It is not a baseline, not Gate-B, not acceptance
evidence.**

**Cause.** Six statements in that draft read `.Count`. Five were wrapped in
`@()`, which guarantees an array before anything is read from it. One was not:

```powershell
for ($index = 1; $index -le $sheets.Count; $index++)
```

`$sheets` was whatever `$Workbook.Worksheets` handed back, and under
`Set-StrictMode -Version 2.0` a member that is not there is terminating. It is
the same class as Benchmark Run 1 — a member read off a value whose shape was
not guaranteed by construction — in a file written before that lesson was turned
into a control.

**Corrected in this round.** The shape is gone rather than guarded: no COM
collection is indexed by position or asked for its `.Count` anywhere in the
probe. Collections are enumerated through `Measure-ProbeCollection`, an absent
collection is a refusal rather than a zero, and the probe carries a stage cursor
so the next failure names the stage, the endpoint, the line and the statement.

The same round tightened the probe's evidence, because Run 1 also showed the
first draft would have accepted too little: an endpoint that announced success
without reshaping any table would have counted as proof that the structural
operation is permitted. **FINE** now requires success **and** an observed shape
change **and** protection in force on both sides of every command.

---

## Protection probe Run 2 — INCONCLUSIVE

**Probe commit:** `ad84ea6`

Stage A 351/351. Stage-B bootstrap PASS. The workbook opened with **14 of 14
sheets protected**, structure protected, and the protection owner reporting
`True` — the first direct confirmation that the protection lifecycle survives a
reopen. The probe reached the locked-cell control and then failed.

**Status: INCONCLUSIVE. It does not prove production is blocked. It does not
prove production is fine. `PCCM_ApplyTimeline` was never invoked.**

### Two defects, one run

**1. The control contradicted itself.** The transcript printed, in order:

```
a cell VALUE write on a protected sheet SUCCEEDED
a cell VALUE write on a protected sheet was REFUSED: System.NullReferenceException
UserInterfaceOnly permits code VALUE writes: True
```

The write and the **restore** shared one `try`. The write succeeded and printed;
the restore threw; the `catch` printed a refusal over the top; `$controlWorked`
had already been set `$true`, so the closing summary said `True`.

The restore threw because the original value was **blank** — and the accepted
`Set-NamedValue` has a `ClearContents` branch for exactly that case which this
probe's copy had dropped. Worse, the target was `inpDiscountRate`, which is an
**editable input** with `Locked = False`: it was never a locked-cell control at
all, so it could not have proved anything about `UserInterfaceOnly` even had it
worked.

**2. `DISP_E_BADINDEX` at `$ws = $sheets.Item($SheetName)`.** Not a missing
sheet: all five watched tab names and all five table names exist in the built
workbook, and a control now proves that against the artifacts without Windows.
Every shape read re-acquired `$Workbook.Worksheets` and released it again —
dozens of times per run against one underlying collection — and the lookup
eventually ran against a collection that had been released out from under it.

**Corrected in this round.** Worksheets and ListObjects are resolved **once**
and held for the session, with the tab name, CodeName and table name recorded as
evidence; an unresolvable sheet or table throws rather than falling back to an
index. The control now targets a cell it **proves** is `Locked = True` on a
protected sheet, keeps the write and the cleanup in separate `try` blocks,
verifies the restored value, and returns exactly one word — `SUCCEEDED`,
`REFUSED` or `INCONCLUSIVE`. A cleanup failure is never rewritten as a refusal
of the write, and an untrustworthy control stops the probe before the production
question is asked.

The endpoint path gained an `$invoked` flag set in the instant before
`Application.Run` and nowhere else, so a probe-side failure can never be read as
a production result — the log now distinguishes **PREPARING TO TEST** from
**ENDPOINT INVOKED**.

---
