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

## Protection probe Run 3 — INCONCLUSIVE

**Probe commit:** `a0a4dc5`

Stage A 351/351. Stage-B bootstrap PASS. Workbook opened with 14 of 14 sheets
protected, structure protected, owner reporting `True`. The probe then failed
resolving its watched targets.

```
stage     : resolve
doing     : resolving every watched worksheet and table
endpoint  : (none)
exception : System.Management.Automation.RuntimeException
message   : the workbook has no worksheet named
            'Cost Lines Risk Register Cost Profiling Risk Profiling Inflation'
            (asked for by the manifest entry
            cost_lines risk_register cost_profiling risk_profiling inflation)
underlying: System.Runtime.InteropServices.COMException: Invalid index. (DISP_E_BADINDEX)
at line   : 311
```

**Status: INCONCLUSIVE. `PCCM_ApplyTimeline` was NOT invoked** — the failure was
in target resolution, before any endpoint. It does not prove BLOCKED and does
not prove FINE.

### Root cause: five records became one

`Get-ProbeWatchedTables` ended with `return ,@($watched)`. The unary comma makes
the function emit **one** pipeline item that *is* the array — and the caller
wrote `@(Get-ProbeWatchedTables ...)`, which **collects pipeline items rather
than flattening nested arrays**. The five records arrived double-wrapped.
`foreach ($entry in @($Watched))` then bound `$entry` to the inner array,
`$entry.Sheet` became **member enumeration** over five records, and `[string]`
joined the result with spaces.

The coercion is what hid it: `[string]` turned a structural error into a
plausible-looking worksheet name, and the only symptom was `DISP_E_BADINDEX`.

**This also re-explains Run 2.** That run died on the same lookup with the same
HRESULT, and the earlier diagnosis blamed COM release churn. The joined name was
already the cause; Run 2's message simply did not carry the name. The Run-2
"resolve once" change was a genuine robustness improvement but did not address
this, which is why Run 3 failed at the same wall.

### The second instance would have been worse than an abort

`Get-ProbeShapeDelta` used the same idiom with the same `@()` at its call site.
Double-wrapped, an **empty** change list arrives as a one-element array, so
`StructuralEffect` would have been true for every endpoint — and the probe could
have reached **PRODUCTION IS FINE** having observed no shape change at all. A
wrong conclusion is more dangerous than a failed run.

### Corrected in this round

Both producers emit their records normally; the callers' `@()` keeps a zero- or
one-record result an array. The collection boundary is now proved **before Excel
is started**: five records (counted from the manifest, never a literal), each
with scalar `Key`/`Sheet`/`Table` validated **before** any `[string]` cast,
unique keys and unique sheet/table pairs, and a record that is itself an array
refused by name. Each target prints in its own block with tab, CodeName and
table on separate lines, with the record count shown.

The locked-cell control gained a fourth state. Run 3 printed
`UserInterfaceOnly permits code VALUE writes: False` after failing before the
control ran — `False` reads as "blocking was observed", which nothing had
tested. The states are now `SUCCEEDED` / `REFUSED` / `INCONCLUSIVE` /
`NOT ATTEMPTED`, there is no boolean projection at all, and an untested
capability prints **NOT TESTED**. This does not touch the verdict logic.

---

## Stage-B verification attempts 1 and 2 — PROBE NOT STARTED

**These are not probe runs.** The protection probe did not execute, produced no
verdict, and produced no evidence about whether protection blocks production
table operations. They are prerequisite failures, recorded so the count of probe
attempts stays honest: the probe still stands at Run 3, INCONCLUSIVE.

They are **not** a baseline, **not** Gate-B evidence and **not** acceptance
evidence.

### What both attempts reported

The Stage-B **build** completed on both:

```
  [PASS] Apply the locked worksheet CodeNames      14 sheets
  [PASS] Import every manifest-declared VBA module 32 modules
  [PASS] Write the ThisWorkbook document module    ThisWorkbook: Workbook_Open
  [PASS] Create the Phase-4 command buttons        11 buttons
  [PASS] Apply passwordless protection             14 sheet(s); structure=True; UserInterfaceOnly=True
  [PASS] Save the Stage-B workbook
  [PASS] Build instance closed naturally
```

The Stage-B **verification** then failed, identically, both times:

```
  [FAIL] Verify the reopened .xlsm
         Dashboard: System.Runtime.InteropServices.COMException:
         Call was rejected by callee. (0x80010001 RPC_E_CALL_REJECTED)
```

### The exact statement, established from source rather than the HRESULT

`build_stage_b.ps1` section 8 formats a per-sheet problem as
`("{0}: {1}" -f $sheet.name, (Format-Err $_))` and a per-button problem as
`("{0}: {1}" -f $button.shape_name, ...)`. `Dashboard` is `sheets[0]` of
`stage_b_manifest.json` and matches **no** button `shape_name` — every button is
`btnPCCM*`. So the failing statement is the **first iteration of the CodeName
loop**: `$worksheets2.Item($sheet.name)` or the `[string]$ws.CodeName` that
follows it.

The sequence that reached it, and what each step proves:

| # | Call | Outcome |
|---|---|---|
| 1 | `New-Object -ComObject Excel.Application` | succeeded |
| 2 | `$excel2.Visible/DisplayAlerts/AskToUpdateLinks = $false` | succeeded |
| 3 | `$excel2.Workbooks` | succeeded |
| 4 | `$workbooks2.Open($stageBPath)` | **returned** — a throw here would have been reported by the outer catch, with no `Dashboard:` prefix |
| 5 | `[int]$wb2.FileFormat` | succeeded — no problem was recorded for it |
| 6 | `$wb2.Worksheets` | succeeded |
| 7 | `$worksheets2.Item('Dashboard')` / `.CodeName` | **RPC_E_CALL_REJECTED** |
| 8+ | sheets 2–14, 32 modules, 11 buttons | **all verified** — `$problems` joined to exactly one entry |

### Classification: A, with a structural gap that is B-flavoured

**D — workbook or Dashboard defect: ruled out.** A missing sheet raises
`DISP_E_BADINDEX`, a wrong CodeName is reported as a mismatch, and neither is an
RPC rejection. The other thirteen sheets, all thirty-two modules and all eleven
buttons verified from the same instance moments later.

**C — production `Workbook_Open` defect: not supported.** `Workbook_Open` runs
inside `Workbooks.Open`, which returned. Everything it is responsible for —
protection across fourteen sheets and workbook structure — had already been
applied by the build and verified there. Nothing it produced failed to verify.
The reopened instance *is* the first place this bootstrap ever executes VBA, so
the correlation is real, but a correlation is not the defect and no evidence
here makes it one.

**B — verifier COM-lifecycle defect: ruled out as the cause.** The failing read
is the **first** iteration, before any `Release-Transient` in that loop has run,
so release churn cannot explain it. (This is the diagnosis that was wrong in
probe Run 2, and the sequence refutes it here rather than repeating it.)

**A — transient COM busy: what the evidence supports.** RPC_E_CALL_REJECTED is
an OLE message-filter result. Its contract is that the call was **not
delivered**: it did not run and no state changed. One call was refused; the
thirty-odd that followed were answered.

What is *not* claimed: **why** Excel was busy at call 7 and not at calls 5 and 6.
That would need instrumentation this batch has not run, and neither the fix nor
the classification depends on it.

The B-flavoured gap that made a refused call fatal: **no Windows harness in this
repository had any retry or readiness handling at any COM boundary.** A single
message-filter rejection anywhere in verification became a hard FAIL of a step
that had verified everything else correctly.

### What was changed, and what deliberately was not

`com_lifecycle.ps1` gained `Invoke-ComRetryRead`: it reads **one member of one
COM object** — a property, or `.Item(key)` — and reissues the call **only** when
the error is RPC_E_CALL_REJECTED or RPC_E_SERVERCALL_RETRYLATER, the two results
whose contract is that nothing ran. At most 12 attempts, each delay capped at
2000 ms, total wait capped at 15000 ms. On exhaustion the **original** exception
is rethrown and the step fails exactly as it did before. There is no scriptblock
parameter, so no write, `Open`, `SaveAs`, `Import`, `Protect` or `Run` can be
expressed through it at all.

`build_stage_b.ps1` section 8 now performs every read through it. The build
block is deliberately **not** wrapped: its calls write, and a write Excel may or
may not have accepted is not something to reissue on a guess.

Not done, and why:

* **No readiness barrier.** The evidence refutes it: calls 5 and 6 were answered,
  so a barrier before the loop would have passed and call 7 would still have been
  refused. A barrier here would imply a guarantee it does not provide.
* **No blind sleep**, anywhere — a control refuses `Start-Sleep` in the bootstrap.
* **No production change.** `modProtection`, `modAppState`, the structural
  commands and `phase10_protection_probe.ps1` are untouched. The probe question
  is still open and still frozen.
* **Nothing suppressed.** Events stay enabled, `Workbook_Open` still runs,
  protection still applies. The verification instance opens the workbook exactly
  as it will open for a person.

`Add-Step 'Transient COM rejections'` reports the outcome on **every** run —
including `0 ms waited` when nothing was refused — so a retried run and an
untroubled one can never be mistaken for each other.

### Still owed

A clean Windows Stage-B run. Until one exists, this correction is **unverified
on Windows**: no Windows execution was performed in this batch.

---

## Protection probe Run 4 — INCONCLUSIVE, but not empty

Stage A 351/351. Stage-B bootstrap PASS. Stage-B reopened verification PASS —
14 CodeNames, 32 modules, 11 buttons, **0 transient COM rejections**, clean
shutdown. The bounded read retry added for the previous batch was not needed on
this run and reported so; that is the outcome it is designed to make visible.

The probe then got further than any run before it.

### CONFIRMED: UserInterfaceOnly permits code VALUE writes to a locked cell

The workbook opened with 14/14 worksheets protected and workbook structure
protected, and `modProtection` reported protection applied. The locked-cell
control ran to completion for the first time:

| Fact | Observed |
|---|---|
| Target | `Cost Lines!tblCostLines` header cell |
| Precondition | `Locked=True`, worksheet protected |
| Temporary value write | **SUCCEEDED** |
| Value readback | **SUCCEEDED** |
| Original restored | exactly |
| Restoration readback | **SUCCEEDED** |
| Protection after | still applied |

**This is valid Windows evidence and is retained.** It establishes that
`UserInterfaceOnly=True` does permit code-driven VALUE writes to a locked cell
while worksheet protection is active.

It says nothing whatever about ListObject structural operations. Permission to
write a value and permission to add a column are different permissions, and the
probe exists to separate them.

### UNRESOLVED: ListObject structural-operation behaviour

`PCCM_ApplyTimeline` was **NOT INVOKED**. `$invoked` is set only in the instant
before `Application.Run`, and the run never reached it.

### FINAL VERDICT: INCONCLUSIVE

### What failed, and it was the probe's own instrumentation

```
stage     : endpoint
doing     : setting the timeline inputs
endpoint  : PCCM_ApplyTimeline
exception : System.InvalidCastException
message   : Unable to cast object of type 'System.Double' to type 'System.String'.
at line   : 175
statement : $range.Value2 = $Value
```

`Set-ProbeNamedValue`, called three times with `[double]` values for
`inpBaseYear` / `inpProjectStartYear` / `inpDurationYears` — three single,
unlocked, `0`-formatted cells on `Setup` (C12, C11, C10).

### Root cause: a diagnosis this repository already owns

**PowerShell binds a COM property setter per call site.** A single polymorphic
`$x.Value2 = $Value` line cannot carry more than one CLR type. This is not
inferred from the exception text — Phase-5 Runtime Run 4 hit the identical
exception with the identical type pair at `phase5_gate_b_scenarios.ps1:922`, and
the settlement is recorded there and in `Set-Phase5TypedCell`: one COM
assignment site per type, each with its own cast, "a no-op that exists only to
give the branch its own bound call site".

The probe had **reimplemented** the accepted `Set-NamedValue` and dropped both
things that make it work:

| | accepted `Set-NamedValue` | probe's copy |
|---|---|---|
| numeric write | `$rng.Value2 = [double]$Value` | `$range.Value2 = $Value` |
| null/blank | `$rng.ClearContents()` | *(absent)* |

The missing `ClearContents` branch was already on the record — it is what made
Run 2 print SUCCEEDED and REFUSED from one try block. The same reimplementation
carried a second defect nobody had looked for.

### What is NOT the fix

Casting the value to `[string]` would satisfy the COM binder and then be refused
by production: `modTimeline.ReadTriple` gates every element of the triple
through `TryReadDouble` and then `d = Int(d)`. A text `"2026"` produces a
**REFUSED** endpoint — and a refusal has no way to distinguish itself from
protection blocking the work. Stringifying would have manufactured the very
verdict the probe exists to test.

### Corrected in this round

* The probe now carries a **verbatim copy** of the accepted `Set-NamedValue`,
  proved verbatim by control — the same arrangement `phase10_benchmark.ps1` uses.
* Cell writes that are not named ranges go through `Set-ProbeCellExact`: one
  assignment site per CLR type, an unsupported type refused **by name**.
* The locked-cell control's restore was the **same class of latent defect** — a
  polymorphic write-back of a captured value, correct today only because the
  header happens to be text. It now dispatches on the captured type.
* The restoration check no longer compares `[string]` forms.
  `Test-ProbeExactValue` establishes CLR type identity first, so a `Double`
  restored as the text `'2026'` is caught.
* Every timeline input is **read back and type-checked before** the endpoint is
  touched. A setter or readback failure stops the run as probe instrumentation
  failure; the endpoint is not entered.
* The stage wording now separates `SETTING ENDPOINT PRECONDITIONS` and
  `VERIFYING ENDPOINT PRECONDITIONS` from `ENDPOINT INVOKED`, both carrying
  `production NOT invoked`, so no future transcript can imply production ran
  while fixture preparation was still going.

### Still owed

The first structural production endpoint has still never been invoked. Nothing
here changes production, and no Windows execution was performed in this batch.

---

## Protection probe Run 5 — BLOCKED BY PROTECTION

Stage A 351/351. Stage-B bootstrap PASS. Stage-B reopened verification PASS —
14 CodeNames, 32 modules, 11 buttons, 0 transient COM rejections, clean
shutdown. The probe reached the real production endpoint for the first time.

### Original assumption — kept, not rewritten

The accepted Phase-10 contract (**`6ab8f6a`**) held that `UserInterfaceOnly:=True`
would let every accepted command go on writing exactly as it did before
protection existed, so no command would ever need to unprotect anything. That
was a reasonable reading of what `UserInterfaceOnly` means, it was recorded
honestly, and it turned out to be wrong about one specific capability. **It
remains historical evidence and is not being rewritten.**

### Windows evidence

Workbook state before invocation: 14/14 worksheets protected, workbook structure
protected = True, `modProtection` reports protection applied,
`UserInterfaceOnly=True`.

**The locked-cell control — CONFIRMED.** Target `Cost Lines!tblCostLines` header,
`Locked=True`, sheet protected. Code-driven temporary VALUE write succeeded;
readback succeeded; original exact value restored; restoration verified;
protection remained applied.

**The endpoint — REFUSED.**

```
PCCM_ApplyTimeline    endpoint invoked: TRUE    outcome: REFUSED

FAIL|Error 1004: Table features aren't available because the sheet is
protected.|The applied timeline, both profiling grids and the inflation grid
have been restored to their state from before this operation, including row
count, column count, number formats and column widths. No partial change
remains.

protection before : 14/14 sheets protected; structure=True
protection after  : 14/14 sheets protected; structure=True
structural effect : NONE
```

The failing path needs `ListColumns.Add` on both profiling grids and the
inflation grid, which a fresh workbook always requires.

### What this establishes

`UserInterfaceOnly=True` separates two capabilities:

| Capability | Verdict |
|---|---|
| ordinary code-driven cell VALUE writes | **PERMITTED** |
| Excel ListObject STRUCTURAL mutation while the worksheet is protected | **NOT PERMITTED** |

The transactional rollback also worked exactly as designed: the command refused,
restored all three grids, and left protection intact. That is production
behaving correctly under a capability it did not have.

### Reconciliation

Retain protection at rest and `UserInterfaceOnly` for value writes. Permit only
the sole protection owner to open a temporary, depth-safe **worksheet**-protection
window for declared structural commands, restoring the full accepted protection
state before the operation returns.

The obsolete absolute rule

> Do NOT call `Unprotect`

is replaced by

> Only `modProtection` may temporarily release worksheet protection inside the
> contracted structural-operation envelope, and it must restore the full accepted
> protection state before the operation returns.

**Workbook structure protection is not released.** The 1004 named the *sheet*,
and workbook-structure protection governs adding, deleting and renaming
worksheets rather than table columns. Releasing it would be wider than any
evidence asks for. If a future Windows run proves a specific need, that is its
own reconciliation.

### The verdict wording this run also corrected

Run 5 summarised its own result as "2 of 4 production endpoints did not
succeed". The second was `PCCM_Calculate` refusing because the applied timeline
was still pending — a business prerequisite that would refuse on a completely
unprotected workbook too. **BLOCKED now requires evidence attributable to the
protected structural operation**: Excel's own "table features aren't available
because the sheet is protected", from an endpoint that was actually invoked. A
validation refusal, a missing input, stale state or an unmet structural
prerequisite cannot produce BLOCKED, and endpoints refused for other reasons are
named as *not counted*.

This run stays **BLOCKED** because `PCCM_ApplyTimeline` supplies exactly that
evidence.

### Still owed

The fix is **unverified on Windows**. No Windows execution was performed in the
reconciliation batch.

---
