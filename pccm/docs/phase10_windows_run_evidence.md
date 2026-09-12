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

## Protection probe Run 6 — the reconciliation works; verdict still INCONCLUSIVE

First Windows run against `0946cf6`, the structural-window reconciliation.
Git fast-forwarded `58b2394 → 0946cf6`.

**Ordering caveat, recorded because it matters.** The first Stage-A command used
`python3`, which does not exist on this Windows host, and failed. The probe was
then run **before** Stage A was rebuilt with `python`. The successful fresh
Stage-A build (`351 passed, 0 failed`) came **after** this probe. This run is
therefore **not** ideal acceptance evidence and is not treated as such.

### What the run established

| Observation | Result |
|---|---|
| Workbook opened | 14/14 sheets protected, structure = True |
| UserInterfaceOnly locked-cell VALUE write control | SUCCEEDED |
| `PCCM_ApplyTimeline` | **invoked and SUCCEEDED** |
| — structural effect | `tblCostProfiling` 25×2 → 25×5; `tblRiskProfiling` 25×2 → 25×5; `tblInflation` 10×1 → 10×4 |
| Protection after the endpoint | 14/14 sheets protected, structure = **True** |
| `PCCM_AddCostLine`, `PCCM_AddRisk` | both SUCCEEDED, protection fully applied |
| — watched shape change | **none** |
| `PCCM_Calculate` | invoked, **REFUSED** — `FAIL\|Calculate\|Discount Rate: the value is blank. A blank is not zero.` |
| — protection during/after | fully applied; no 1004, no protection refusal |
| Structural initialisation A–E | **PROVEN** |
| Final verdict | **INCONCLUSIVE** — 1 of 4 endpoints did not succeed |

### The open question this settles

The reconciliation batch released **worksheet** protection only and deliberately
left **workbook-structure** protection applied, on the reading that Excel's 1004
named the *sheet*. That was an inference, and it is now **runtime evidence**:

> `ListColumns.Add` on `tblCostProfiling`, `tblRiskProfiling` and `tblInflation`
> succeeded with `ThisWorkbook.ProtectStructure = True` throughout.

**The privilege envelope is not widened.** Workbook-structure protection stays
applied inside the structural window, and a control refuses a window that
releases it.

### Why the verdict is still INCONCLUSIVE, and why that is correct

`PCCM_Calculate` refused for a **non-protection** reason. The corrected verdict
logic did exactly what the previous batch built it to do: it did **not** call
that a protection block, and named it as refused-for-other-reasons. Protection
never entered the picture.

### What Run 6 did NOT establish

* **`ListRows.Add` capacity expansion.** Add Cost Line and Add Risk succeeded
  with protection intact, and correctly produced no watched shape change: Stage A
  reserves 25 register rows, so an Add writes an id into a reserved row and
  `ListRows.Add` never fires. **Endpoint functionality under protection was
  observed; capacity expansion was not.** This probe does not claim otherwise.
* **The `_Calc` resize path.** `PCCM_Calculate` never reached it.

### Corrected in this round (source only — no Windows)

**The refusal was production being right.** `modCalcResolve.ResolveAppliedTimeline`
requires `NM_INPUT_DISCOUNT_RATE` through `NumericNamedCell`, which refuses a
blank by design and refuses a numeric-looking *string* too, because
`IsRealNumber` tests the VarType rather than parsing. So the probe now supplies
it as a real `Double` through the accepted `Set-NamedValue`, at the value the
accepted Phase-7 fixture already declares (`discount_rate = 0.05`). Nothing is
hard-coded around the validation.

**Calculate now runs immediately after ApplyTimeline, before the Add commands.**
`AddDriver` writes a permanent ID, and `ReadRegister` reads every row whose id
column is non-blank — so after an Add the register holds one identified driver
with every other field empty, and Calculate refuses on it. Filling those fields
would mean the probe manufacturing a cost line. An empty driver set is valid in
the contract's own words, and with zero drivers no currency and no inflation
profile is *referenced*, so FX and the deliberately-blank inflation grid are
never resolved. The applied timeline and the discount rate are the whole
prerequisite — and this is the user's own first Calculate.

**Calculate is no longer judged by its announcement.** The probe said the `_Calc`
tables "are not watched above, so judge this one by its announcement", which
contradicts its own central argument. The five `_Calc` tables are now read before
and after, from `calc.sheet` and `calc.tables[*].table_name` in the Gate-B
inspection. The check is **predictive, not merely a difference**:
`calc_years` and `calc_annual` carry the row rule *"one row per applied project
year"*, and Stage A builds every `_Calc` table with one body row — so a
3-year timeline must leave them at exactly 3 rows, which only `ResizeBody` can
produce. A Calculate that announces success without that shape yields
`CONTRADICTED` and **cannot reach FINE**. A Calculate that *refused* is required
to change nothing and is not failed for it.

### Still owed

A Windows run in which Stage A is rebuilt **first**, and in which
`PCCM_Calculate` reaches its `_Calc` resize. The reconciliation itself remains
**unverified for the Calculate path**.

---

## Protection probe Run 7 — PRODUCTION IS FINE UNDER PROTECTION (for the ADD direction)

First run with Stage A rebuilt **before** the probe, against `5b14a81`.

```
Stage A: 351 passed, 0 failed
```

### What Run 7 observed

| Observation | Result |
|---|---|
| Workbook opened | 14/14 sheets protected, `ProtectStructure = True` |
| `PCCM_ApplyTimeline` | **SUCCEEDED** |
| — structural effect | `tblCostProfiling` 25×2 → 25×5; `tblRiskProfiling` 25×2 → 25×5; `tblInflation` 10×1 → 10×4 |
| `PCCM_Calculate` | **SUCCEEDED** |
| — `_Calc` structural evidence | **OBSERVED** — `tblCalcYears` 1×3 → 3×3; `tblCalcAnnual` 1×8 → 3×8 |
| `PCCM_AddCostLine`, `PCCM_AddRisk` | SUCCEEDED; correctly no shape change (reserved capacity remained) |
| Protection after every endpoint | fully restored |
| Probe verdict | `PRODUCTION IS FINE UNDER PROTECTION` |

This is strong evidence and it is retained exactly as recorded.

### What Run 7 proved — and what it did NOT

**Proved, under the production structural window:**

* `ListColumns.Add` — three grids gained project-year columns.
* `ListRows.Add` / `ResizeBody` growth — the per-project-year `_Calc` tables grew
  from one body row to three.
* Protection restored after every endpoint, with workbook structure protection
  never released.

**Not proved:**

* **`ListRows.Delete`.** Benchmark Run 3 died specifically on a
  `ListRow.Delete()` with *"Table features aren't available because the sheet is
  protected"*. Every round in Run 7 moved in the **ADD** direction. No successful
  add says anything about a delete.
* `ListColumns.Delete` — likewise never exercised.

**Therefore the broad conclusion that the Benchmark Run 3 defect was a "HARNESS
defect only" remains PENDING exact delete-path runtime evidence.** Run 7's
`PRODUCTION IS FINE UNDER PROTECTION` was reported honestly against the criteria
that existed at the time; those criteria did not yet require the delete
direction. This is recorded as a coverage gap in the criteria, not as an error in
the run.

### Added in this round (source only — no Windows)

A deterministic **shrink round** now follows the successful 3-year round:

1. `inpDurationYears` is set to a genuine `Double` **1** through the same
   accepted `Set-NamedValue`, read back and type-checked exactly as every other
   input. `1` is inside the accepted bound (`modTimeline.ReadTriple` requires
   `1 ≤ d ≤ LIMIT_MAX_YEAR_COLUMNS`), so this is a valid timeline, not a value
   chosen to force a failure.
2. `PCCM_ApplyTimeline` is invoked again. Every grid that **gained** columns in
   the growth round must **lose** them — expectation derived from what was
   actually observed growing, so a partial shrink is refused. This drives
   `modProfiling.SetYearColumns` and `modInflation.SetYearColumns` down their
   `ListColumns(...).Delete` loops.
3. `PCCM_Calculate` is invoked again through the same round helper. The per-year
   `_Calc` tables must end at the shrunk duration **and must have lost rows to
   get there** — accepting 3 → 3 would call an unchanged table delete evidence.
   This is `ResizeBody`'s `ListRows(...).Delete` loop: the call Benchmark Run 3
   died on.

Protection is now checked **before and after every endpoint**, with
`ProtectStructure` read as a boolean rather than inside a sentence.

`PRODUCTION IS FINE UNDER PROTECTION` now requires **all four** classes:

```
LISTCOLUMN ADD    : OBSERVED
LISTCOLUMN DELETE : OBSERVED
LISTROW GROWTH    : OBSERVED
LISTROW DELETE    : OBSERVED
```

If the second `ApplyTimeline` or the second `Calculate` refuses, errors,
announces success without the contracted shape, or leaves protection wrong, the
verdict is **not** FINE.

### Still owed

**The delete path is unproved on Windows.** Nothing in this round is runtime
evidence; production VBA is byte-identical to `0946cf6`.

---

## Protection probe Run 8 — PROBE REGRESSION, INCONCLUSIVE, no endpoint reached

Windows executed against `c8e2d02`. **This is not a production failure.** The run
ended before any production endpoint was exercised.

### What succeeded

```
Stage A (rebuilt BEFORE the probe): 351 passed, 0 failed
```

Stage-B bootstrap completed: 14 worksheets, 32 VBA modules, protection applied,
workbook `structure=True`, `UserInterfaceOnly=True`, persistence verification
passed, both Excel instances shut down naturally, transient COM releases clean.

### What failed

```
stage     : protection
doing     : reading the protection state as the workbook opened
exception : System.Management.Automation.PropertyNotFoundException
message   : The property 'ProtectContents' cannot be found on this object.
            Verify that the property exists.
statement : if ($sheet.ProtectContents) {
                $protectedNames += [string]$sheet.Name
            }
at line   : 637
```

```
CONTROL : NOT ATTEMPTED
          UserInterfaceOnly code-value-write capability: NOT TESTED
VERDICT : INCONCLUSIVE
          the probe itself failed in stage protection while reading the
          protection state as the workbook opened
```

Shutdown was clean.

### What this run establishes, and what it does not

* **No production defect is established.** No endpoint was invoked.
* **No delete-path evidence was obtained.** Nothing here bears on
  `ListColumns.Delete` or `ListRows.Delete`.
* **Run 7 remains valid** for exactly what it already proved.
* **The protection architecture is untouched** and remains supported by Run 7.
* Runtime Protection Reconciliation remains **OPEN**.

### Root cause — and what is proved versus inferred

**PROVED FROM SOURCE.** `Get-ProbeProtectionState` is **byte-identical** between
`5b14a81` (Run 7, succeeded) and `c8e2d02` (Run 8, failed):
`git diff 5b14a81 c8e2d02` over the probe touches only `Get-ProbeCalcShapes`'
release label and the new shrink-round helpers. Nothing that executes before it
changed either — Run 8 failed on the **first** call, three statements after
`Workbooks.Open`. `$sheet` is bound in exactly three places in the file, the
other two are function-local to helpers that run later, and PowerShell function
locals do not leak. **The delete-path work did not introduce this.** It is a
latent defect in the reader that Run 8 exposed.

**INFERRED, AND LABELLED AS SUCH.** `PropertyNotFoundException` is PowerShell's,
not Excel's. A COM object for which PowerShell could not obtain type information
exposes **no properties at all**, and every access on it reports exactly this
message. Excel had just run `Workbook_Open` → `modProtection.ProtectionApply`
across 14 sheets, and the accepted Stage-B verification already **proved on this
machine** that the first inbound calls into a freshly-opened instance can be
refused (`RPC_E_CALL_REJECTED`). A refused `IDispatch::GetTypeInfo` is the shape
that produces this exception. That last step cannot be proved from Linux, so the
correction is built to settle it either way rather than repeat an undiagnosable
failure.

### Corrected in this round (source only — no Windows)

1. **Every enumerated item is proved to be a Worksheet before any property is
   read**, by a membership test on `.PSObject.Properties` — which is always safe,
   unlike the dereference that crashed. A non-Worksheet now fails **explicitly**
   with its .NET type, its PSTypeNames, whether it is a COM object, how many
   properties it exposes, the item index and the stage. One run will diagnose it.
2. **The reads go through the accepted `Invoke-ComRetryRead` boundary** in
   `com_lifecycle.ps1` — the same one Stage-B verification uses. It reissues only
   the two OLE message-filter HRESULTs, whose contract is that the call never
   ran. If a refused call is what broke Run 8, that is now survived; if it is
   not, the helper rethrows untouched.

Nothing is weakened: `ProtectContents` is still read per worksheet,
`ProtectStructure` is still read per workbook, both from the real Excel
properties, and a workbook that enumerates no worksheets is still refused.
`PropertyNotFoundException` is still fatal — it is never caught, never tested
for, and never defaulted.

### The delete-path design is unchanged

The shrink round added at `c8e2d02` stands. The next run must still exercise
timeline growth → `ListColumns.Add`, Calculate growth → `ListRows.Add`, duration
3 → 1, timeline shrink → `ListColumns.Delete`, Calculate shrink →
`ListRows.Delete`, and `PRODUCTION IS FINE UNDER PROTECTION` still requires all
four evidence classes.

### Still owed

Everything Run 8 was meant to obtain. Production VBA is byte-identical to
`c8e2d02`, `5b14a81` and `0946cf6`.

---

## Protection probe Run 9 — all four classes OBSERVED; predicate defect, not a production defect

Windows executed against `04fcf82`. **The runtime evidence is valid and the
acceptance predicate was defective.** Those are two different things and are
recorded as two different things.

### Runtime production evidence — VALID

| Class | Evidence |
|---|---|
| **LISTCOLUMN ADD** | first `PCCM_ApplyTimeline`: `cost_profiling` 25×2 → 25×5, `risk_profiling` 25×2 → 25×5, `inflation` 10×1 → 10×4 |
| **LISTROW GROWTH** | first `PCCM_Calculate`: `tblCalcYears` 1×3 → 3×3, `tblCalcAnnual` 1×8 → 3×8 |
| **LISTCOLUMN DELETE** | shrink `PCCM_ApplyTimeline`, duration 3 → 1: `cost_profiling` 25×5 → 25×3, `risk_profiling` 25×5 → 25×3, `inflation` 10×4 → 10×2 |
| **LISTROW DELETE** | second `PCCM_Calculate`: `tblCalcYears` 3×3 → 1×3, `tblCalcAnnual` 3×8 → 1×8 |

Protection was **14 of 14 sheets protected with `ProtectStructure = True` before
and after every invoked endpoint**. No protection 1004 occurred. The locked-cell
`UserInterfaceOnly` control succeeded. Shutdown and COM release were clean.

**This directly exercises the same `ListRows.Delete` class that Benchmark Run 3
failed on.** The substantive runtime question is settled: the accepted production
structural window supports the required Add **and** Delete operations while
workbook structure remains protected.

### Acceptance / reporting predicate — DEFECTIVE

The same report printed, and this is recorded exactly as it was printed:

```
C. NOT MET  PCCM_ApplyTimeline SUCCEEDED, created year columns, and left protection applied
STRUCTURAL INITIALISATION: NOT PROVEN - C not met.
...
LISTCOLUMN ADD    : OBSERVED
LISTCOLUMN DELETE : OBSERVED
LISTROW GROWTH    : OBSERVED
LISTROW DELETE    : OBSERVED
PRODUCTION IS FINE UNDER PROTECTION
```

**Criterion C did print NOT MET, and this record does not pretend otherwise.**

### Root cause — proved from source

Criterion C selected its round by **endpoint name**:

```powershell
$timelineOutcome = @($outcomes | Where-Object { $_.Endpoint -eq 'PCCM_ApplyTimeline' })
Met = ((@($timelineOutcome).Count -eq 1) -and ...)
```

The `-eq 1` was a correctness guard written when exactly one `ApplyTimeline`
existed. The shrink round added a **second** outcome carrying the same endpoint
name, the count became **2**, and C went false on a run whose growth apply had
succeeded, grown all three grids and left protection applied.

**It was never reading post-shrink workbook state.** `$outcomes` is append-only
and each record is an immutable snapshot taken around its own endpoint. The
*selector* simply stopped being unique. The predicate was also computed **after**
the verdict, so it could not gate it — which is why NOT PROVEN and FINE could
appear in one report.

### Corrected in this round (source only — no Windows)

* Every criterion binds to **the recorded round itself** (`$timeline`,
  `$addCost`, `$addRisk`) rather than filtering the outcome list by a name that
  is no longer unique. A third round cannot silently falsify a criterion again.
* Criterion C is explicitly *the FIRST (growth) `PCCM_ApplyTimeline`*, and now
  checks protection on **both** sides including the structure flag.
* The shrink round keeps its own separate representation as **LISTCOLUMN
  DELETE**; neither side is weakened by the other.
* The criteria are computed **before** the verdict and **gate** it:
  `PRODUCTION IS FINE UNDER PROTECTION` is unreachable unless A–E are all MET,
  all four classes are OBSERVED, and every protection-before/after snapshot is
  valid. A run can no longer print NOT PROVEN and FINE together.
* One computation, one report — the printed criteria are the ones the verdict
  used.

The locked-cell control now reaches the verdict through criterion B, **by
declaration**. It can only ever *block* FINE; it can never grant one and can
never produce BLOCKED.

### Production conclusion

**The protection architecture is not reopened.** Run 9's runtime evidence
supports the accepted contract, production VBA is byte-identical to `04fcf82`,
`c8e2d02`, `5b14a81` and `0946cf6`, and no privilege-envelope expansion was made:
workbook structure protection is still never released.

---

## Protection probe Run 10 — CLOSURE ARTIFACT — RECONCILIATION ACCEPTED AND CLOSED

Windows executed against **`0119bee`**. This is the self-consistent closure
artifact the predicate reconciliation was built to produce.

```
Stage A: 351 passed, 0 failed
Workbook opened: 14/14 worksheets protected, ProtectStructure = True
Locked-cell UserInterfaceOnly control: SUCCEEDED

Growth PCCM_ApplyTimeline : SUCCEEDED   ListColumns.Add on all three grids
Growth PCCM_Calculate     : SUCCEEDED   tblCalcYears 1 -> 3, tblCalcAnnual 1 -> 3
Shrink PCCM_ApplyTimeline : SUCCEEDED   duration 3 -> 1, all three grids lost columns
Shrink PCCM_Calculate     : SUCCEEDED   tblCalcYears 3 -> 1, tblCalcAnnual 3 -> 1

Protection fully restored after every endpoint.

A. MET   B. MET   C. MET   D. MET   E. MET
STRUCTURAL INITIALISATION: PROVEN

LISTCOLUMN ADD    : OBSERVED
LISTCOLUMN DELETE : OBSERVED
LISTROW GROWTH    : OBSERVED
LISTROW DELETE    : OBSERVED

PRODUCTION IS FINE UNDER PROTECTION
```

Shutdown and COM release were clean.

### RUNTIME PROTECTION RECONCILIATION — ACCEPTED AND CLOSED

1. The accepted protection architecture is **settled**.
2. Worksheet protection may be temporarily released **only** inside the
   contracted production structural-operation window.
3. Workbook **structure** protection remains applied throughout.
4. Benchmark Run 3's `ListRow.Delete` failure is classified as a **HARNESS
   defect**, not a production defect — the failing call was the harness's own
   COM `$victim.Delete()` in `Remove-TableRow`, not a production endpoint.
5. This architecture is **not to be reopened or redesigned** without genuinely
   new runtime evidence.
6. **No further Windows protection probe is required.**

**This closes ONLY the Runtime Protection Reconciliation.** It does not close
Phase 10, does not establish a benchmark baseline, and is not final project
acceptance.

### What this closure does NOT resolve — the harness consequence

Classification 4 is precise, and it has a consequence that is still open.

`phase10_benchmark.ps1:313` still performs `$victim.Delete()` — a direct COM
`ListRow.Delete` on `tblFXRates`, reached from `Set-Phase5Fixture` →
`Reset-Phase5FxTable` → `Remove-TableRow` while building the PERF-SMALL fixture.

**The production structural window does not help it.** That window is opened by
`modAppState.BeginStructuralOperation` from inside production VBA; a PowerShell
COM caller never enters it. Every accepted Windows harness was written against a
Stage-B workbook that had no protection — protection is a Phase-10 addition
applied by `build_stage_b.ps1` step 7 and re-applied by `Workbook_Open`.

So a Benchmark Run 4 against the current tree **would abort at the same line with
the same 1004**. That is recorded here as the open item it is, not as a surprise
for the next run to rediscover.

---
## Benchmark fixture under protection — SETTLED IN SOURCE — NO WINDOWS

Settles the open item recorded immediately above. **No Windows was executed for
this round**; the section above stays exactly as it was written, because it was
true of the tree it described.

### The failing path, traced

```
Set-Phase5Fixture                    phase5_gate_b_scenarios.ps1:1834
  Invoke-Phase5FixtureSteps                                  :1869
    step C  Reset-Phase5FxTable                              :1651
              Remove-TableRow  ->  $victim.Delete()   phase10_benchmark.ps1:313
```

`tblFXRates` is built by the input contract as `data_rows: 12` with one seeded
row, so `Get-TableRowCount` returns **12** and the accepted reset's
`for ($row = $rows; $row -gt 1; $row--)` loop issued **eleven** `ListRow.Delete`
calls. The first one is the 1004. Every row it was deleting was already blank.

### What the fixture actually requires afterwards

Two things, and only two:

1. row 1 **is** the captured locked seed — value for value and type for type;
2. every row below row 1 carries **no currency and no rate**.

### Physical deletion is not required — proved in production's source

`modCalcResolve` is the only production consumer of `tblFXRates`.

| Could depend on physical deletion | Does it? | Where |
|---|---|---|
| `ListRows.Count` | **No** | never read |
| `DataBodyRange` dimensions | **No** | never read as a value |
| physical table row count | **No** | `BodyRowCount` is a loop bound in `MatchingFxRows` only |
| row position / index | **No** | `firstMatch` only fetches that row's own rate |
| presence/absence of blank body rows | **No** | `RawCellText` exits `False` on `IsEmpty`, so a blank row is never a match candidate |
| formulas / validation on retained rows | **Yes — and deletion DESTROYED them** | the reserved rows carry `lstCurrencies` and decimal-greater-than-zero validation |
| structural fingerprint / shape invariant | **No** | the invariant is "the reporting currency appears exactly once", which blank rows cannot affect |

A blanked row is therefore indistinguishable from an absent one to every
production reader. It is also the shape production already runs against: Stage A
delivers eleven blank FX rows, and every accepted Phase-4 through Phase-9 run
resolved FX over exactly that.

**The deletion was the operation doing damage.** It removed validated rows from a
table the input contract declares as `data_rows: 12`, quietly shrinking the
delivered shape. Blanking leaves the contract's validation where the contract
put it.

### Every structural fixture site found, and its disposition

| # | Site | Call | Fired in Run 3? | Disposition |
|---|---|---|---|---|
| 1 | `Reset-Phase5FxTable`, the `rows -lt 1` guard | `ListRows.Add` | No — the table has 12 rows | gone; the override refuses with a named diagnosis |
| 2 | `Reset-Phase5FxTable`, the delete loop | `ListRow.Delete` ×11 | **Yes — this is the abort** | replaced by `ClearContents` below the seed |
| 3 | `Invoke-Phase5FixtureSteps` step C, the FX append | `ListRows.Add` ×1 | No — Run 3 died first | **would have been the next abort**; now a reserved-row search |
| 4 | `Set-Phase5InflationProfileMaster`, the capacity guard | `ListRows.Add` | No — 2 profiles ≤ `data_rows: 10` | latent; now refuses with a named diagnosis instead of a bare 1004 |

`Clear-Phase5UserRows` also calls `Remove-TableRow`, and the `fx_remove` arm of
`Invoke-Phase5Mutation` does too. Neither is reachable at runtime: the first has
no caller anywhere in the tree, and the second belongs to the Gate-B scenarios
the benchmark never executes.

### The correction's own first attempt was wrong, and the audit caught it

`Remove-TableRow` was deleted outright. That passed every text control — nothing
in the benchmark calls it — and
`tests/powershell_command_resolution_audit.ps1` refused it:

```
UNRESOLVED Remove-TableRow
```

The audit is an AST closure over the runner **and every file it dot-sources**, so
the two callers above are exactly the reason the name must still resolve. This is
the defect class that ended Phase-9 Windows run 1 — `The term 'Write-RowObject'
is not recognized`, raised from inside a dot-sourced file — and deleting the
definition recreated it.

So the name resolves and the **capability** is gone: `Remove-TableRow` now throws,
naming the row, the table and why a COM caller cannot delete it. It fetches no
COM object at all, so it cannot be mistaken for a working helper, and there is no
empty body for a later edit to fill back in.

The same audit reported the override itself:

```
DUPLICATE reset-phase5fxtable defined at phase10_benchmark.ps1, phase5_gate_b_scenarios.ps1
```

That rule — *a name defined twice is a finding, because the reader cannot tell
which one runs* — was true of every runner in this tree until one needed to change
a helper without editing the accepted file that defines it. It is not loosened.
The audit gained a `-DeclaredOverride` parameter, and the declaration is
**checked**: the name must be defined exactly twice, once in the runner and once
in a dot-sourced file, and the runner's definition must sit **after** the
dot-source that loads the other — which is the only thing that makes it the
definition that wins. A declared name that overrides nothing is itself a finding,
and every caller that does not pass the parameter keeps the original behaviour
exactly. Nothing checked load order before; this is strictly stronger than the
rule it relaxes.

No `ListColumns.Add`, `ListColumns.Delete` or `Resize` appears anywhere in the
benchmark's fixture or setup code.

### Why the correction is in `phase10_benchmark.ps1` only

`Reset-Phase5FxTable` lives in `phase5_gate_b_scenarios.ps1`, an accepted harness
whose bytes are not changed for the sake of a measurement. That file **uses**
`Remove-TableRow`, `Add-BlankTableRow`, `Set-TableCell` and `Get-TableBody` and
**defines** none of them — every standalone runner carries its own copies — and
PowerShell resolves a function name at call time. The benchmark's definitions
follow its dot-source, so they are the ones the accepted fixture tree reaches
**in the benchmark process only**. No other runner changes and no accepted
evidence is disturbed.

### The protection boundary is untouched

No `ProtectionRelease`, no maintenance window, no `Unprotect` from PowerShell, no
change to `Workbook_Open`, and no production VBA change. `git diff --name-only`
against `0119bee` over `src`, `spec` and `builder` is empty.

### What is inferred rather than observed, stated plainly

`ClearContents` on a locked cell of a protected sheet has **not yet been observed
on Windows**. It is in the same capability class as the value write that Run 3
and Run 10 both proved `UserInterfaceOnly` permits, and it is not a ListObject
structural operation, which is the one capability proved refused — but that is an
inference from the proved split, not an observation.

The run does not have to be trusted on it. The reset reads every blanked cell
back and **requires `$null`**, and its refusal names protection as a candidate
cause. The PERF-SMALL baseline run is therefore the first observation either way:
it clears and proceeds, or it stops with a sentence saying which cell refused.

## Benchmark Run 4 — PERF-SMALL — ABORTED BUILDING THE SCENARIO

**Harness commit:** `7077608`

Stage A 351 passed, 0 failed. Stage-B bootstrap succeeded and protection was
applied: 14 sheets, `ProtectStructure = True`, `UserInterfaceOnly = True`. The
run passed the preflight, the bootstrap, the workbook open and the environment
inventory, then aborted while constructing the fixture — before any timed run.

```
stage     : scenario
doing     : building the PERF-SMALL fixture through the accepted production endpoints
scenario  : PERF-SMALL
exception : System.Runtime.InteropServices.COMException
message   : The cell or chart you're trying to change is on a protected sheet.
            To make a change, unprotect the sheet.
statement : $null = $cell.ClearContents()

BASELINE STATUS    : ABORTED BEFORE A COMPLETE BASELINE
valid warm medians : 0 of 11 planned run(s)
```

Shutdown and COM release were clean.

**Status: 0 of 11 valid warm medians. NOT a baseline, NOT a partial baseline,
NOT a performance sample.**

### What this run established

**External COM `ClearContents` under worksheet protection is RUNTIME-REFUTED.**
The settlement at `7077608` recorded honestly that this had not been observed and
that it was an inference from the proved capability split. The inference was
wrong, the reset's own read-back is what turned it into a named statement rather
than a silent stall, and the correction it was part of — reset by content, no
structural deletion — is unaffected and stays.

**No production defect is established.** The failing call was the harness's, from
an out-of-process PowerShell COM client. Production VBA was not reached.

**It does not contradict the Runtime Protection Reconciliation.** That evidence
was VBA executing *inside* the workbook, where `UserInterfaceOnly:=True` means
what it says. This is an external automation client, and the capability split for
such a client is finer than the one recorded for VBA:

| External COM caller, protected sheet | Result | Evidence |
|---|---|---|
| `Range.Value2 = <value>` on a locked cell | **permitted** | Benchmark Run 3 — four Setup scalars written and read back |
| `Range.ClearContents()` | **refused** | Benchmark Run 4, above |
| `ListRow.Delete` / `ListRows.Add` | **refused** | Benchmark Run 3; probe Runs 5–10 |

Those three lines are now all runtime facts. None of them is inferred.

---

## Benchmark fixture maintenance window — SETTLED IN SOURCE — NO WINDOWS

Settles Benchmark Run 4. **No Windows was executed for this round.**

### The call path

```
phase10_benchmark.ps1   Open-BenchmarkFixtureWindow
                          $excel.Run('P10FW_Begin')
phase10_fixture_window.bas   P10FW_Begin
                               modProtection.ProtectionBeginStructural(detail)
   ... Set-Phase5Fixture ...
phase10_benchmark.ps1   Close-BenchmarkFixtureWindow
                          $excel.Run('P10FW_End')
phase10_fixture_window.bas   P10FW_End
                               modProtection.ProtectionEndStructural(detail)
                          Assert-BenchmarkProtectionApplied
                            $excel.Run('P10FW_State')
```

### Why a shim, and why it is not a new production API

`modProtection`'s two mutators are `Public Function` in a standard module and are
reachable by name, but both take `ByRef detail As String`. **Every** procedure any
accepted harness in this tree has ever invoked through `Application.Run` —
production `PCCM_*` and Gate-B `GBD_*` alike — takes `ByVal` parameters or none.
There is no precedent here for marshalling a `ByRef` out-parameter across that
boundary, and the failure mode would be a lost or mangled diagnostic on the one
call whose failure must abort the run.

`bootstrap/windows/phase10_fixture_window.bas` owns the `String` inside VBA and
returns a `String`, which is the shape that is proven. It is **harness-owned**: it
lives beside `phase5_gate_b_diagnostics.bas`, which Gate B has imported into the
disposable workbook by this same mechanism since Phase 5; it is never declared in
`stage_b_manifest.json`, and the runner refuses to import it if it ever is. It
holds no protection policy — no `.Protect`, no `.Unprotect`, no branch of its own
on protection state. It forwards, and it reports.

It adds **no new machine dependency**: `build_stage_b.ps1`, which this runner
already invokes on every run, reaches `$wb.VBProject` to build the workbook at
all.

### Workbook structure protection is never released

`ProtectionRelease` — the maintenance path — is the **only** procedure in
production that calls `ThisWorkbook.Unprotect`, and a control pins that count at
one. Neither the runner nor the shim names it in code. `ProtectionBeginStructural`
releases worksheet protection only, and the runner refuses outright if
`ProtectStructure` is not `True` on either side of the window.

### The writes that need the window, and the ones that do not

| Fixture step | Write | Needs the window |
|---|---|---|
| A | `PCCM_DeleteCostLineById` / `PCCM_DeleteRiskById` | production endpoints — they open their own |
| A, B | `Set-NamedValue` — counters, timeline, discount rate | **no** — `Value2 =`, permitted (Run 3) |
| C | `Set-TableCell -Value $null` → `ClearContents` on `tblFXRates` | **YES — this is the Run 4 abort** |
| C | `Set-Phase5TypedCell` — the locked FX seed row | `Value2 =`, but inside the same span |
| C | `Set-TableCell` — the appended FX row | `Value2 =`, but inside the same span |
| D | `Set-TableCell -Value $null` → `ClearContents` on `tblInflationProfiles` | **YES — the next abort** |
| E, F | `PCCM_ApplyTimeline`, `PCCM_AddCostLine`, `PCCM_AddRisk` | production endpoints, nested inside |
| G | `Write-Phase5InflationRates`, `Write-Phase5Weights` | `Value2 =`, inside the same span |
| — | the random seed, after the fixture | **no — deliberately left outside** |

The window spans `Set-Phase5Fixture` and nothing else. It cannot be narrowed
further without editing `phase5_gate_b_scenarios.ps1`, whose steps interleave
`ClearContents` with production endpoint calls, and that file is not edited.

### Setup only, and provably outside every measurement

The window opens immediately before `Set-Phase5Fixture` and closes immediately
after it, inside the fixture stopwatch — whose figure the runner already reports
as setup and uses in no measurement. It is closed, and the closure verified,
before the run loop exists. Controls read the ORDER off positions in the source
rather than off a comment promising it.

The close sits in a `finally`, so a fixture that raises still closes the window;
if the close then fails too, its failure is the one reported, because an
unprotected workbook is the worse fact.

### What is verified, and the one thing that cannot be

Before opening and after closing: the workbook reports itself protected
(`modProtection.ProtectionIsApplied` — production's own single fact), every
worksheet is protected, `ProtectStructure` is `True`, the structural depth is `0`,
and the worksheet count matches `stage_b_manifest.json`'s protection projection —
read from the projection, never from the literal 14.

The state is read **inside VBA** and returned as one string. Probe Run 8 failed
reading `Worksheet.ProtectContents` across COM: PowerShell had no type information
for the object and StrictMode turned a missing member into a terminating
`PropertyNotFoundException`. In VBA the property binds at compile time and that
class is gone.

**`UserInterfaceOnly` cannot be read back.** Excel exposes no property for it — it
is a write-only argument to `Worksheet.Protect`. It is therefore established by
CONSTRUCTION, not by inspection: the outermost close routes through
`modProtection.ProtectionApply`, the same single function `Workbook_Open` uses,
which unprotects and re-protects every sheet with `UserInterfaceOnly:=True`, and
`ProtectionEndStructural` then requires `ProtectionIsApplied` before reporting
success. That is the accepted architecture's own definition of restored, and this
record says plainly that it is a construction argument rather than a read-back.

### Any failure aborts before timing

Every refusal is a `throw`, and a throw inside the measurement session sets
`$abandoned`, which forces `$runComplete` false, which makes the status ABORTED
and the exit code 1. No benchmark result can be recorded from a workbook whose
protection was not restored and verified.

### Nothing about what is measured changed

Scenario identity, driver counts, project years, the iteration matrix, cold/warm
counts, the timing boundaries and the baseline policy are untouched. The
content-based FX reset stays and the structural `ListRow.Delete` stays gone — a
mutation that restores it under the open window, where it would now succeed, is
caught.

## Fixture window: the partial-open gap — SETTLED IN SOURCE — NO WINDOWS

Found while proving the exception-safety property that was asked for before the
PERF-SMALL run, and closed in the same round. **No Windows was executed.**

### What was asked, and what held

*If `Open-BenchmarkFixtureWindow` succeeds and `Set-Phase5Fixture` throws at any
point, is `P10FW_End` guaranteed to be attempted before the error reaches the
outer abort path?* **Yes**, and it was already true: the close is in a `finally`
inner to the session `try`, so PowerShell unwinds it first. Shutdown is never
relied on — and could not help anyway, since it closes the workbook without
saving rather than restoring anything.

### The adjacent gap that did not hold

`Open-BenchmarkFixtureWindow` could run `P10FW_Begin` successfully and then throw
in its OWN post-open checks — before the caller had entered its `try`/`finally`.
Nothing closed what had been opened. `modProtection` names that outcome the worst
one there is, and a disposable workbook does not make it acceptable.

Three routes existed: the depth read-back not returning 1, the protection-state
read refusing or raising, and the `ProtectStructure` assertion failing.

### The contract now

```
Begin refuses          -> nothing was opened, nothing is owed, no End is called
Begin succeeds, then
  a post-open check
  fails                -> exactly ONE compensating End, then the original failure
                          is rethrown WITH whether the rollback took
Open returns           -> the window is open, verified, depth exactly 1, and this
                          function has closed nothing
Fixture succeeds/fails -> the caller's finally closes exactly once
```

A `catch`, not a `finally`: a `finally` would run on the success path too and
close a window the caller still expects to hold, and one that threw would REPLACE
the original exception with the rollback's. `return` inside a `try` does not
enter its `catch`, so there is no path on which both this function and the caller
close the same window. The depth cannot be decremented twice.

`Invoke-BenchmarkWindowRollback` cannot raise — it is called from a catch that is
about to rethrow. It returns a sentence instead, and distinguishes a REFUSAL from
a RAISE, because "the workbook says it could not restore protection" and "the
call never arrived" are different facts about the machine. Nothing is swallowed:
every outcome becomes text in the exception that is thrown.

### Proved by execution, not by reading

`tests/phase10_fixture_window_flow.ps1` lifts the real functions — and the
caller's own `try`/`finally` — out of `phase10_benchmark.ps1` by AST and by
anchored text, and runs them against a fake that records every
`Application.Run`. Excel is never started. "Exactly one compensating close" is a
COUNT of what PowerShell does; asserting it from source would be asserting a
belief about the language.

| Scenario | macros | `P10FW_End` | timed |
|---|---|---|---|
| Begin refuses | State, Begin | **0** | 0 |
| post-open depth ≠ 1 | State, Begin, State, End | **1** | 0 |
| post-open state read refuses | State, Begin, State, End | **1** | 0 |
| post-open state read raises | State, Begin, State, End | **1** | 0 |
| post-open `ProtectStructure` false | State, Begin, State, End | **1** | 0 |
| post-open fails, rollback refuses | State, Begin, State, End | **1** | 0 |
| post-open fails, rollback raises | State, Begin, State, End | **1** | 0 |
| open ok, fixture ok | State, Begin, State, End, State | **1** | **1** |
| open ok, fixture throws | State, Begin, State, End, State | **1** | 0 |
| open ok, close refuses | State, Begin, State, End | **1** | 0 |
| open ok, close leaves depth 1 | State, Begin, State, End, State | **1** | 0 |

Timed work follows exactly one row. On the two rollback-failure rows the thrown
message carries both the original failure and `IT REFUSED` / `IT RAISED`.

### Terminology, corrected

An earlier return described one of these routes as "structure released". That was
`ThisWorkbook.ProtectStructure` becoming False — **workbook structure**
protection, not worksheet protection. The window releases worksheets and only
worksheets.

**No path can make `ProtectStructure` False.** Production contains exactly one
`ThisWorkbook.Unprotect`, at `modProtection.bas:245`, inside `ProtectionRelease`
— the maintenance path, which neither the runner nor the shim calls.
`ProtectionBeginStructural` contains none. The assertion is belt-and-braces
against a future production change, not a live route. **No new defect.**

## Benchmark Run 5 — PERF-SMALL — REACHED THE TIMED SECTION — ABORTED

**Harness commit:** `ce5951f`

Stage A 351 passed, 0 failed. Stage-B bootstrap succeeded. **The fixture
maintenance window worked**, and this is the run that proves it:

```
PROTECTION
  as opened         : applied=True|depth=0|structure=True|sheets=14|protected=14
  after the fixture : applied=True|depth=0|structure=True|sheets=14|protected=14
                      [window closed and verified before any timed run]

Cost Lines in book  : 12   (plan: 12)
Risks in book       : 8    (plan: 8)
project years       : 10
```

The window opened, the fixture was built through it, it closed, protection was
restored and depth was verified `0` — all before the first timed operation.
**Workbook structure protection stayed applied throughout.** The fixture-window
architecture is RUNTIME PROVEN and is not to be reopened.

The run then reached the timed section for the first time in the project and lost
it to two harness shape defects.

```
Calculate
  cold  0.707 s   INVALID: System.Object[]
  warm  0.412 s   INVALID: System.Object[]
  warm  0.444 s   INVALID: System.Object[]
  warm  0.464 s   INVALID: System.Object[]
  WARM MEDIAN : NOT COMPUTED

Workbook recalculation
  cold  0.290 s   INVALID: System.Object[]
  warm  0.299 s   INVALID: System.Object[]
  warm  0.297 s   INVALID: System.Object[]
  warm  0.311 s   INVALID: System.Object[]
  WARM MEDIAN : NOT COMPUTED

Run Simulation  @ 10000 iterations
  stage     : measurement
  doing     : setting the iteration control to 10000 and re-establishing the
              deterministic basis
  exception : Cannot convert the "System.Int32[]" value of type "System.Int32[]"
              to type "System.Double".

BASELINE STATUS    : ABORTED BEFORE A COMPLETE BASELINE
valid warm medians : 0 of 11 planned run(s)
```

Shutdown and COM release were clean.

**Status: 0 of 11 valid warm medians. NOT a baseline, NOT a partial baseline, NOT
a performance sample.** Calculate and the workbook recalculation DID execute —
those eight elapsed times are real — but no sample was accepted, so none of them
is a measurement of record.

**No production defect is established.** Both defects are PowerShell shape
defects in the harness. Production VBA was invoked and answered correctly
throughout.

---

## The two shape defects — SETTLED IN SOURCE — NO WINDOWS

**No Windows was executed for this round.** The two defects are independent and
have different root causes; neither is a symptom of the other.

### Defect 1 — `INVALID: System.Object[]`: a double-wrapped problem list

`Test-BenchmarkSample` ends `return ,@($problems)`. The leading comma hands back
**one object that IS the array** — which is what stops an EMPTY result being
enumerated into nothing by the pipeline. The caller then wrote
`$problems = @(Test-BenchmarkSample ...)`.

**`@()` collects pipeline items; it does NOT flatten a nested array.** So it
collected that one object into a new one-element array whose single element was
the real list:

```
caller @() : Count=1  Valid=False  join=System.Object[]
assigned   : Count=0  Valid=True   join=
```

Therefore `$problems.Count` was **1 whatever the sample found**,
`Valid = ($problems.Count -eq 0)` was **always False**, and
`($execution.Problems) -join '; '` rendered the inner array as its type name.
Calculate and the recalculation had no problems at all — the validator found
nothing and the shape said otherwise.

This is the same defect class the project recorded at Gate-B Run 3: *`@()`
collects pipeline items; it does not flatten nested arrays.*

**The same pairing existed at one other site**, latent:
`$roots = @(Get-BenchmarkOneDriveRoots)`. Double-wrapped, every OneDrive root
became one nested array, so a workbook under OneDrive was recorded as a LOCAL
location and the roots field held a list containing a list. Corrected with it.
`New-BenchmarkWeights` returns `,@(...)` too and was already assigned directly —
correct, and left alone.

**The correction is the shape, not the strictness.** The caller assigns; and
because a contract spread over two places is one an edit can half-keep — it was
half-kept twice — `Assert-BenchmarkProblemList` now refuses a nested or
non-string element AT THE CALL, by name. It **throws** rather than flattening: a
malformed problem list is a defect in the harness, not a fact about the workbook,
so recording it as an invalid sample would be the harness marking its own bug as
the model's.

### Defect 2 — `System.Int32[]` in the iteration control: the loop variable was the parameter

The run loop used `$iterations`. **PowerShell variable names are
case-insensitive, so that IS the script parameter `[int[]]$Iterations`** — and a
typed parameter keeps its type constraint for the whole life of the variable, so
every assignment to it is coerced back to `[int[]]`:

```
assigned 10000 -> type System.Int32[]  value [10000]
[string] gives : 10000
[double] gives : Cannot convert the "System.Int32[]" value of type
                 "System.Int32[]" to type "System.Double".
```

That is why the banner read `Run Simulation @ 10000 iterations` and looked
perfectly right — `[string]` of a one-element array is the element — and only the
`[double]` cast at the iteration-control write raised.

**The plan carries a scalar.** `runs[].iterations` is `10000`, not `[10000]`, and
`$declaredIterations` (the whole list) was never involved. Nothing needed
selecting out of an array, so no `[0]` was added: the shape defect was the NAME.
The loop local is now `$runIterations`, which no parameter can constrain.

**Generalised, not just fixed.** A control now reads the param block out of the
runner and refuses any assignment anywhere to a name the block declares with a
type. The three `[string]` path parameters are defaulted in place on purpose and
are named in that control rather than excluded by a pattern.

### Proved by execution

`tests/phase10_run_shape_flow.ps1` lifts the runner's **own param block** and the
loop's **own selection lines** by AST, composes them into a script, and runs it
once per planned PERF-SMALL run — because the param block is the defect's other
half, and a probe that declared its own `[int[]]$Iterations` would be testing its
own guess. It also lifts `Test-BenchmarkSample` and `Assert-BenchmarkProblemList`
and drives them over scripted evidence. Excel is never started.

| planned run | CLR type | value | `[double]` |
|---|---|---|---|
| calculate | `<null>` | n/a | n/a |
| recalculation | `<null>` | n/a | n/a |
| simulation / sensitivity / annual × 10000 | `System.Int32` | 10000 | 10000 |
| simulation / sensitivity / annual × 50000 | `System.Int32` | 50000 | 50000 |
| simulation / sensitivity / annual × 100000 | `System.Int32` | 100000 | 100000 |
| with `-Iterations 10000` actually supplied | `System.Int32` | 10000 | 10000 |

Nine iteration-dependent runs, each exactly one integer; two carrying none;
eleven in total.

| sample case | count | valid | shape guard |
|---|---|---|---|
| clean command | 0 | **True** | accepted |
| clean recalculation | 0 | **True** | accepted |
| endpoint refusal | 1 | False | accepted |
| malformed announcement | 1 | False | accepted |
| missing published iterations | 1 | False | accepted |
| wrong published iterations | 1 | False | accepted |
| correct published iterations | 0 | **True** | accepted |
| annual state not CURRENT | 1 | False | accepted |
| three problems at once | 3 | False | accepted |
| **the defect, recreated** | 1 | False | **REFUSED by name** |

Validity stayed strict: every refusal still refuses. The two rows that are now
`valid=True` are the state that was previously impossible.

### What did not change

Scenario identity, driver counts, project years, the iteration matrix, cold = 1,
warm = 3, eleven planned runs, the median gate (a median exists only when EVERY
warm sample was valid, and the count comes from the plan), the timing boundaries,
and the baseline policy. The fixture window and its shim are byte-identical to
`ce5951f` — the run that proved them — and a control compares each of its six
functions against that commit.

## PERF-LARGE attempt 1 — OPERATOR-ABORTED IN FIXTURE CONSTRUCTION

**Harness commit:** `f3b3a33`

Stage A 351 passed, 0 failed. Stage-B PASS — 14 sheets, 32 VBA modules, 11
buttons, protection applied, reopened verification PASS, clean COM lifecycle.
Protection at benchmark open:
`OK|applied=True|depth=0|structure=True|sheets=14|protected=14`.

Environment: Windows 11 Pro, Excel 64-bit, AMD Ryzen 7 7800X3D, 31.2 GB RAM,
workbook on local temp storage, repository under OneDrive.

The run reached `BUILDING PERF-LARGE` and stayed there for **more than four
hours** without reaching the first timed production operation. The operator
aborted it — the run is **operator-aborted**. Shutdown was clean: `Workbook.Close = True`,
`Application.Quit = True`, natural PID exit `True`, emergency required `False`,
all transient COM releases clean.

```
PERF-LARGE WINDOWS RUN: OPERATOR-ABORTED DURING FIXTURE CONSTRUCTION AFTER >4 HOURS
0 timed operations completed
0 warm medians
```

**NOT a baseline. NOT a partial baseline. NOT performance evidence for the PCCM
runtime.** No Calculate, Simulation, Sensitivity or Annual timing began.

### What it is evidence of, and what it is not

It is evidence about the **benchmark fixture-construction method**. The original
endpoint-by-endpoint method is therefore **operationally impractical** for the
Large scenario.

It **does NOT mean** any of the following, and none may be inferred from it:

* that a Large calculation takes >4 hours;
* that Simulation takes >4 hours;
* that the Large scenario is unsupported.

**No production defect is established.** Every production command the fixture
invoked answered correctly; there were simply three hundred of them.

---

## The fixture cost, counted

Derived from source, not estimated. Two costs, and the smaller one is the
harness's.

### The harness's cross-process COM calls

Counted from the helpers' own bodies in `phase10_benchmark.ps1` —
`Set-TableCell` is 7 calls, `Get-TableBody` is `9 + 2·rows·cols`,
`Get-TableRowCount` is a full `Get-TableBody`.

| | SMALL | MEDIUM | LARGE |
|---|---|---|---|
| step F, the Adds and their two verifications each | 24,816 | 157,560 | **1,123,080** |
| step G, rates and profiling weights | 3,007 | 23,817 | 110,627 |
| everything else | 7,638 | 9,898 | 19,018 |
| **total** | **35,461** | **191,275** | **1,252,725** |

Step F is 70% of SMALL and **89.7%** of LARGE, because
`Invoke-Phase5AddDriverAndRequireSuccess` reads the whole register **twice** per
Add — once through `Get-IdColumnValues` and once for the permanent identifier —
so the scans alone are O(N²·cols).

### The production cost, which is the dominant one

Each `PCCM_AddCostLine` / `PCCM_AddRisk` is a full structural operation
(`modDrivers.RunDriverOperation`):

| inside one Add | scope |
|---|---|
| `modProtection.ProtectionBeginStructural` | unprotect 14 sheets, then a second pass over all 14 to **prove** the release |
| `modWorkbook.SnapshotTable` ×2 | every cell of the register **and** of the profiling grid, for the rollback |
| `modProfiling.SyncRows` | every existing weight into a Dictionary, then the grid rewritten |
| `modStructuralCheck.ValidateStructure` | the register and the grid again |
| `modAppState.FinishOperation` | re-apply protection to 14 sheets and verify, then `RecalculateStructuralState` |
| `modAppState.RecalculateStructuralState` | `.Calculate` on Setup, Cost Profiling, Risk Profiling, Inflation |

That is **O(register rows × project years) per Add**, so N Adds is
**O(N²·years)**:

| | Adds | in-VBA cell visits | worksheet recalcs | sheet protect ops |
|---|---|---|---|---|
| SMALL | 20 | 46,580 | 80 | 1,400 |
| MEDIUM | 100 | 543,760 | 400 | 7,000 |
| **LARGE** | **300** | **5,816,730** | **1,200** | **21,000** |

LARGE is **10.7×** MEDIUM on in-VBA cell visits and each of its 1,200
recalculations covers a far larger grid. MEDIUM's fixture was tolerable; ×10.7 on
a superlinear term is the four hours.

**All of that is production behaving correctly.** One user adding one cost line
*should* snapshot for rollback, re-sync the grid, validate and re-protect.
Production is not changed to make a benchmark convenient.

---

## The bulk fixture

### What must come from production, and does

| product | producer |
|---|---|
| the year columns on all three grids | `modProfiling.SetYearColumns`, `modInflation.SetYearColumns` |
| the profiling **rows** and their permanent-ID keying | `modProfiling.SyncRows`, `modInflation.SyncProfileRows` |
| the applied-timeline defined names | `PCCM_ApplyTimeline` |
| structural validation | `modStructuralCheck.ValidateStructure`, `PCCM_StructuralReport` |
| **every calculation, simulation, sensitivity and annual result** | the timed endpoints, untouched |

The first four all arrive from **ONE real `PCCM_ApplyTimeline`**, which does
exactly them.

### What the builder writes

Only what a user types — the register business columns, the FX rates, the Config
profile names, the profiling weights, the Setup scalars — plus the two identity
artifacts production would have issued: the permanent IDs and the two counters.
`modDrivers.AllocateId` increments the counter, persists it, and formats prefix +
the sequence zero-padded to the declared width, so N adds always yield
`CL-001..CL-00N` with the counter left at N. The prefix and pad width are read
from the manifest's counter projection, never from a literal.

Those identity artifacts are the one thing here that is not a plain user input,
and they are exactly what the equivalence gate exists to prove.

### It publishes nothing

No write to `_Calc`, no write to `_SimData`, no fingerprint, no state label, no
result of any kind. Controls ban each by name, and the only production operation
the builder invokes is `PCCM_ApplyTimeline`.

### The expected reduction

| | endpoint path | bulk path |
|---|---|---|
| production structural operations | **301** (300 Adds + 1 Apply) | **1** (one Apply) |
| in-VBA cell visits from those | **5,816,730** | the one Apply's own sync |
| worksheet recalculations | **1,200** | **4** |
| sheet protect/unprotect ops | **21,000** | **~70** |
| register body COM writes | 3,420 individual | **2** rectangular |
| profiling weight COM writes | 12,000 individual | **2** rectangular |
| harness COM calls, total | **1,252,725** | **~2,000** |

The superlinear term is removed outright: there is no longer a per-driver
production operation, so nothing is O(N²·years) any more.

### Protection

Unchanged. The builder runs inside the **same** fixture maintenance window the
endpoint path uses — no new release, no `Unprotect` of its own, no second
authority. `ListRows.Add` to grow a register is a structural operation and is
legal only inside that window, which is where it runs. The window's six functions
and the shim are byte-identical to `ce5951f`, the run that proved them.

### Timing semantics

Unchanged. The builder sits between the fixture stopwatch's start and stop, so
its cost is reported as the setup it is; the run loop is downstream of the window
closing; cold is 1 and warm is 3; a median still requires every warm sample
valid. The iteration matrix is untouched — Large stays capped at 50,000 with
eight planned runs.

---

## The fixture equivalence gate

**RUN AND ACCEPTED at `8caffb0`** — see *Equivalence run 13 — ACCEPTED* at the end
of this record. **EVERY FIELD FAMILY MATCHED** and `CALCEQUIV|match`; no field
differed. The Bulk fixture is **AUTHORISED for performance benchmark
construction** and the gate is **CLOSED / ACCEPTED**. The paragraphs below
describe the gate as first specified (two Stage-A bundles); its single-baseline
form and runs 1–12 are recorded in their own sections and are not rewritten.

`tests/phase10_fixture_equivalence.ps1` builds PERF-SMALL **both ways**, in two
Excel sessions over two disposable copies of the same Stage-A build, captures a
full state snapshot from each and compares them field for field, then runs the
real `PCCM_Calculate` on both and compares production's own status and
fingerprint. It lifts the builder out of the shipping runner by AST, so it cannot
pass against a restatement of itself, and it asserts nothing — it prints tagged
lines and `test_phase10_benchmark_harness.py` decides.

Field families compared: cost-line identifiers and order; risk identifiers and
order; both counters; both register bodies in full, column by column; the FX
table; the Config profile master; the applied timeline defined names; the
generated year headers on all three grids; the Cost Profiling and Risk Profiling
weight bodies; the inflation grid; `nmStructuralState`; `PCCM_StructuralReport`;
`PCCM_CurrentInputFingerprint`; `PCCM_CurrentSimulationRequestFingerprint`;
`PCCM_ModelCheckCalculationState`. Then `PCCM_Calculate` on both, comparing
`PCCM_CalculationStatus` and `PCCM_CalculationFingerprint`.

**Until this gate has run and every family has matched, `-FixtureMode Bulk` must
not be used for a baseline.** The default stays `Endpoints`, which is what the
accepted SMALL and MEDIUM baselines were built by, and the artifact records
`fixture_mode` so no baseline can be compared across methods unseen.

### What is still unproven on Windows, stated plainly

Two things, and both are what the gate is for:

1. **`ListRows.Add` from a COM caller inside the open fixture window.** Run 5
   proved production's own structural operations work inside production's window;
   the harness doing it inside the harness's window is the same capability class
   but has not been observed. If it is refused, the gate reports it as a raised
   Bulk pass rather than a silent wrong answer.
2. **A rectangular `Range.Value2 = object[,]` block write** into a ListObject body
   on a protected sheet inside the window. Same class as the value writes Run 3
   proved permitted, but as a block rather than a cell.

**Both were proven on Windows at `8caffb0`:** the Bulk pass completed, grew both
registers by `ListRows.Add`, wrote all four blocks, and reached the same state and
fingerprint as the Endpoints pass.

## Equivalence run 1 — INVALID — NO COMPARISON WAS EXECUTED

**Harness commit:** `99cb472`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed.

Both passes failed **before Excel was started**:

```
stage_b_manifest.json not found at <Endpoints temp>\stage_b_manifest.json
stage_b_manifest.json not found at <Bulk temp>\stage_b_manifest.json
```

No Endpoints fixture was built. No Bulk fixture was built. No Excel comparison
occurred. No `PCCM_Calculate` comparison occurred.

```
Equivalence run 1: INVALID / NOT EXECUTED
```

**No semantic equivalence claim may be made from it.** Bulk remains not approved
for formal baseline use; PERF-LARGE must not be rerun; the endpoint fixture
remains the only proven fixture path.

### The final line was misleading, and so was the per-pass line

The gate printed `EQUIV|<no comparison>|differ|one of the two passes did not
build`. That **does NOT mean the fixtures differ** — it means the harness never
reached comparison. It also printed `PASS|Endpoints|RAISED` and
`PASS|Bulk|RAISED`: a pass that raised before producing a workbook is not a PASS.
Both are corrected below.

---

## The workdir contract, and why the manifest was absent

`build_stage_b.ps1` resolves exactly **three** things against the supplied
`-BuildDir`:

| artifact | where | line |
|---|---|---|
| `stage_b_manifest.json` | `$BuildDir` | 86; the throw at 93–94 is the message above |
| `$manifest.stage_a_filename` → `PCCM_stageA.xlsx` | `$BuildDir` | 98, 100 |
| the **generated** VBA directory, `Split-Path -Leaf $manifest.vba.generated_dir` → `vba` | `$BuildDir` | 125 |

and **two** deliberately against the repository, not the BuildDir:

| artifact | where | line |
|---|---|---|
| `$manifest.vba.source_dir` → `src/vba` (the 29 version-controlled modules) | `$pccmRoot` | 124 |
| `$manifest.vba.document_module.file` → `src/vba/ThisWorkbook.vba` | `$srcDir` | 274 |

That split is the bootstrap's own documented rule: source modules are shared
input; generated projections must come from the build being assembled.
`$manifest.stage_b_filename` is the **output** and must not pre-exist unless
`-Force`.

Nothing else is read from the BuildDir — `grep -c inspection build_stage_b.ps1`
is **0**, so the two Gate-B inspection projections the benchmark also copies are
not part of this contract. The gate reads them from the repository build
directory, which is correct.

**The gate copied two of the three.** It copied the Stage-A workbook and the
`vba` directory and not the manifest, so the bootstrap threw on its first read —
and its message, written for an operator who had never built Stage A, told this
operator to do the thing they had just done.

`git show 99cb472` on the gate shows the two `Copy-Item` calls; the benchmark at
the same revision makes **five**, and the manifest is one of the three the
bootstrap actually needs.

---

## The correction

### One pristine bundle per pass, derived not guessed

`Get-BundleArtifacts` returns the three BuildDir artifacts, taking the generated
directory's name from `$Manifest.vba.generated_dir` rather than naming `vba`.
`New-EquivalenceBundle` creates a directory per pass under the work root, copies
each artifact, hashes **every file that arrived** with SHA-256, and refuses if a
Stage-B workbook is present before the bootstrap runs.

### Starting-state identity, proved before Excel

`Test-BundleIdentity` compares the two digest maps key by key in both directions.
The driver builds **both** bundles first, prints one `BUNDLE|<mode>|<path>|<sha>`
line per file, and only then — if the maps agree — runs the passes. If they
disagree it prints `BUNDLE|differ|…` and **nothing is built**.

Digest keys are separator-normalised, so `vba/modConstants.bas` is the key on any
host.

### The vocabulary

| line | meaning |
|---|---|
| `PASS\|<mode>\|COMPLETED\|…` | this pass built a fixture **and** produced a calculation fingerprint |
| `FAIL\|<mode>\|SETUP\|…` | the bundle could not be prepared |
| `FAIL\|<mode>\|BOOTSTRAP\|…` | Stage-B produced no workbook |
| `FAIL\|<mode>\|RAISED\|…` | anything after Excel started |
| `EQUIV\|<family>\|match\|differ` | a comparison that really ran |
| `EQUIV\|<not evaluated>\|invalid\|…` | no comparison happened, and why |
| `CALC\|…`, `CALCEQUIV\|…` | only from the branch where **both** passes completed |

`PASS` cannot coexist with a failure: completion now requires both a state
snapshot and a non-empty fingerprint. `differ` is reserved for a real comparison.
`CALC` and `CALCEQUIV` are emitted only inside the both-completed branch, so a
setup failure cannot print a placeholder verdict on a calculation that never ran.

### Two defects the Linux harness caught before Windows could

1. **`New-EquivalenceBundle` both wrote and returned.** A function that emits to
   the output stream *and* returns a value has its return polluted: `$bundle`
   came back as an array of report lines with the object at the end, and the very
   first `$bundle.Root` failed. The same "assign, never emit" rule this project
   recorded for COM collections. The caller prints from `Digests` now.
2. **The relative digest key trimmed only a backslash**, leaving a leading
   separator on any host whose separator is not one. Both are trimmed and the
   survivor normalised.

`tests/phase10_bundle_flow.ps1` lifts the three bundle functions by AST and drives
them against a fake build directory: the derived list, five files per bundle,
isolated roots, identity, identity proved non-vacuous by editing one artifact, a
refusal for each missing artifact, and a stale Stage-B workbook proved not to
travel. Excel is never started.

### What did not change

The snapshot function is **byte-identical** to `99cb472` — every field family the
gate was built to compare is still compared. Production, the bulk builder, the
timed path, the timing semantics and the iteration matrix are untouched, and
`-FixtureMode` still defaults to `Endpoints`.

## Equivalence run 2 — INVALID — THE BULK FIXTURE DID NOT BUILD

**Harness commit:** `d1af4e1`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed.

### What run 2 proved

The bundle correction worked. Both disposable starting bundles held the same five
files and every SHA-256 matched:

```
BUNDLE|identical|5 artifact(s)
  stage_b_manifest.json
  PCCM_stageA.xlsx
  vba/modCalcContract.bas
  vba/modConstants.bas
  vba/modSimContract.bas
```

The Endpoints pass completed:

```
PASS|Endpoints|COMPLETED|fixture built and PCCM_Calculate ran
```

Its Stage-B bootstrap was clean — 14 CodeNames, 32 modules, 11 buttons,
protection applied, verification clean, natural COM shutdown. The Bulk pass also
bootstrapped cleanly.

**The result vocabulary worked.** The gate refused to draw a conclusion:

```
FAIL|Bulk|RAISED|tblCostLines already holds 25 body rows where the fixture
                 needs 12. This builder never deletes rows; it runs against a
                 fresh disposable workbook.
EQUIV|<not evaluated>|invalid|comparison was not executed: only the Endpoints
                              pass completed
```

### What run 2 did NOT prove

Bulk fixture construction; fixture semantic equivalence; calculation equivalence;
Bulk authorisation. **No semantic mismatch has been demonstrated.** This
**must not be read as a fixture DIFFER** — no comparison ran.

```
Equivalence run 2: INVALID / NOT EVALUATED
```

---

## Reserved capacity is not semantic count

### The root cause

`Set-BenchmarkRegisterRowCount` took its argument as the number of body rows the
table should **end up with**, and refused when the table already held more. Stage
A builds `tblCostLines` with `reserved_rows: 25` — twenty-five blank body rows —
and twelve Cost Lines occupying twelve of them is not an error. It is the state
production reaches.

**Production's own rule**, `modDrivers.AddDriver`:

```vba
targetRow = FirstFreeRow(Kind, orphanRow)   ' a blank RESERVED row
If targetRow = 0 Then                       ' only when none is left
    register.ListRows.Add                   ' ...is the table grown
```

with its own comment beside it: *"Reserved rows were only ever initial capacity,
never a business maximum."*

### All five tables, audited

| table | Stage-A physical body rows | blank suffix permitted after the last semantic row | when production grows it |
|---|---|---|---|
| `tblCostLines` | 25 (`reserved_rows`) | **yes** — `FirstFreeRow` fills reserved rows | `AddDriver`, only when `targetRow = 0` |
| `tblRiskRegister` | 25 | **yes** — same code path, `Kind` switched | same |
| `tblCostProfiling` | 25 | **yes** — `SyncRows` **clears** the tail, never deletes | `SyncRows`, only when `writeRow > BodyRowCount(target)` |
| `tblRiskProfiling` | 25 | **yes** — same | same |
| `tblInflation` | 10 | **yes** — `SyncProfileRows` clears the tail | `SyncProfileRows`, only when `writeRow > BodyRowCount(target)` |

**No production path shrinks a body.** The two grids and the inflation grid are
production's alone: the builder writes weights into the rows `SyncRows` created
and never touches a grid's row count.

### The corrected algorithm

```
target physical rows = max(existing reserved capacity, semantic driver count)
```

* capacity already sufficient → **nothing is done**, and the blank suffix stays
  exactly as Stage A built it;
* capacity insufficient → grow by **exactly** the shortfall, inside the accepted
  fixture window, and never shrink afterwards;
* the register block covers **exactly** the semantic rows, so identifiers stop at
  N and the counter is N;
* the reserved suffix is **read back and proved blank** before `ApplyTimeline`
  synchronises the grids from that register — a populated unkeyed row is the
  orphan `AddDriver` refuses to mutate over.

### Executed, not asserted from source

`tests/phase10_reserved_rows_flow.ps1` lifts the grower and the block builder by
AST and drives them against a fake whose `ListRows.Add()` is counted. Excel is
never started.

| case | capacity | semantic | physical after | `ListRows.Add` calls |
|---|---|---|---|---|
| SMALL Cost Lines | 25 | 12 | **25** | **0** |
| SMALL Risks | 25 | 8 | **25** | **0** |
| exactly at capacity | 25 | 25 | 25 | 0 |
| one past capacity | 25 | 26 | 26 | **1** |
| MEDIUM Cost Lines | 25 | 60 | 60 | 35 |
| **LARGE Cost Lines** | 25 | 180 | **180** | **155** |
| **LARGE Risks** | 25 | 120 | **120** | **95** |

| drivers | block rows × cols | first ID | last ID | counter | IDs past N |
|---|---|---|---|---|---|
| 12 | 12 × 11 | CL-001 | CL-012 | 12 | **0** |
| 8 | 8 × 11 | CL-001 | CL-008 | 8 | **0** |
| 180 | 180 × 11 | CL-001 | CL-180 | 180 | **0** |

`category` and `uom` are `$null` in every row — blank in the endpoint-built
register too, because `Write-Phase5Driver` never writes them.

### PERF-SMALL expected state, and PERF-LARGE growth

SMALL: 12 semantic Cost Lines and 8 Risks, **physical body row count 25 in both
registers**, rows 13–25 and 9–25 blank, counters 12 and 8. LARGE: 180 and 120
semantic rows, physical 180 and 120 with **no** reserved suffix (capacity is
exhausted), counters 180 and 120, and **no** production Add commands — one
`ListRows.Add` per shortfall row instead.

### The snapshot was not weakened

`Get-EquivalenceSnapshot` is **byte-identical to `99cb472`**. Register bodies are
compared through `Get-TableBody`, which returns every **physical** row with blanks
as empty strings — so if Endpoints had 25 physical rows and Bulk had 12, that
would be a real `differ`. The correction makes Bulk match physically as well as
semantically rather than teaching the comparison to ignore the difference.


---

## Equivalence runs 3 and 4 — INVALID / NOT EVALUATED — STAGE-B WAS REFUSED

**Harness commit:** `1e0edb2`. Windows PowerShell 5.1. The same command was run
twice and failed identically both times.

### What both runs did

| step | outcome |
|---|---|
| bundles | built, and proved identical |
| Endpoints pass | **completed**, `PASS|Endpoints|COMPLETED` |
| Endpoints `PCCM_Calculate` | ran |
| Bulk Stage-B — owned Excel instance | opened |
| Bulk Stage-B — Stage-A workbook | opened |
| Bulk Stage-B — build | **raised `System.Runtime.InteropServices.COMException: Call was rejected by callee. HRESULT 0x80010001 RPC_E_CALL_REJECTED`** |
| shutdown | clean, no forced stop |
| Bulk fixture construction | **never began** |
| semantic comparison | **never ran** |

```
FAIL|Bulk|SETUP|BOOTSTRAP|RAISED
EQUIV|<not evaluated>|invalid|comparison was not executed: only the Endpoints
                              pass completed
```

### What runs 3 and 4 are not

No semantic comparison was executed. This record
**must not be read as a fixture DIFFER**.
They are **not** a Bulk semantic failure either — the Bulk builder was never
reached. They are **not** equivalence evidence in either direction.
**Bulk remains NOT authorised.**

```
Equivalence runs 3 and 4: INVALID / NOT EVALUATED
```

### The old Stage-B log did not identify the exact rejected COM operation

This is the finding that matters, and it is structural rather than inferred. The
Stage-B build is **one** `try`/`catch` around eleven distinct COM operations, and
its catch reported the whole region under one name:

```powershell
} catch {
    Add-Step 'Stage-B build' 'FAIL' (Format-Err $_)
```

So the transcript carried the HRESULT and the name of a **region**. It could not
distinguish `Workbooks.Open` completion from `SaveAs`, from the worksheet
collection, from a CodeName write, from `VBProject`, from `VBComponents`, from the
module import, from the `ThisWorkbook` write, from a button, from protection, or
from the final `Save`.

**`VBProject` acquisition is a CANDIDATE and nothing more.** It is the leading one
on this repository's own runtime precedent — the reopen verification path already
reads `VBProject` through `Invoke-ComRetryRead`, because that member was observed
to be refused at runtime — but precedent about a member is not identification of a
call. Two further candidates are named without preference: `SaveAs`, and the
`Worksheets` acquisition.

**No claim is made here about which call was refused.** The instrumentation added
in this batch exists so the **next** Windows run states it outright.

### What the next run will report

The build block now sets a label from a closed eleven-word vocabulary immediately
before each operation, and the catch reports it:

```
[FAIL] Stage-B build
       operation=vbproject.acquire; System.Runtime.InteropServices.COMException: ...
COMREJECT|build|vbproject.acquire|hresult=0x80010001|RPC_E_CALL_REJECTED (0x80010001)|the call was refused before it ran
```

A call Excel **accepted** and that then failed is reported differently, because it
is a different finding:

```
COMFAIL|build|thisworkbook.write|hresult=0x800a03ec|the call was accepted and failed; it was NOT refused
```

and a clean build says so rather than falling silent:

```
[PASS] Transient COM rejections (build)
       COMREJECT|build|none|attempts=0|waited=0
```

### A second defect, found while reading the gate

`Invoke-EquivalencePass` checked only that a Stage-B workbook **existed** after the
bootstrap. A build refused **after** its `SaveAs` leaves the `.xlsm` on disk with
no modules, no buttons and no protection, and the pass would have opened it and
reported a fixture result against a half-built workbook. The bootstrap's exit code
is now checked first, using the `BOOTSTRAP:` vocabulary the gate already had.

### What was deliberately NOT done

* **No readiness gate** after `Workbooks.Open`. Workbook readiness is not yet known
  to be the failing condition, and a poll inserted before the evidence exists could
  make the symptom disappear without proving its cause.
* **No inter-pass drain, sleep or quiet period.** Both Excel instances shut down
  naturally in the observed runs, so nothing links the second-pass failure to a
  lifecycle overlap.
* **No retry around any mutation** — `SaveAs`, `VBComponents.Import`,
  `CodeModule.AddFromString`, `Shapes.AddShape`, `Protect`, `Save`. If the next run
  names one of these, that is where it stops and a postcondition-aware correction
  is a separate decision.
* **No retry around a property SET**, including the CodeName writes.
* **`Invoke-ComRetryRead` is byte-identical to `1e0edb2`.** This batch adds four
  call sites, not capability.
* The two passes remain two bundles and two Excel sessions.

### The four reads that were opted in, and why only those

| read | label | why it is eligible |
|---|---|---|
| `$wb.FileFormat` | `saveas.xlsm` | property get; the reopen path already reissues `FileFormat` |
| `$wb.Worksheets` | `worksheets.acquire` | collection acquisition; reopen path reissues `Worksheets` |
| `$wb.VBProject` | `vbproject.acquire` | property get; reopen path reissues `VBProject` |
| `$vbproj.VBComponents` | `vbcomponents.acquire` | collection acquisition; reopen path reissues `VBComponents` |

The rule is **precedent, not preference**: a build read is eligible only where the
accepted verification block already reissues the same member on the same class of
object. `$excel.Workbooks` is a pure property get too, and is deliberately **left
out** — verification does not reissue it, so there is no accepted precedent to
apply. If the next run is refused there, that is exactly the outcome the labels
exist to report.

### A host behaviour that had to be worked around, and what it does not prove

PowerShell 7.4.6 **discards** an exception thrown by a **.NET** property getter:
`$obj.$member` yields `$null` and `$Error` is not even populated. A **method** call
raises a `MethodInvocationException` whose `InnerException` is the `COMException` —
which is exactly the chain `Get-ComRejectionName` walks, and why it walks one. So
the executed harness drives the retry loop through a method, and the wrapper's own
logic through a stubbed helper.

**This is a fact about the harness host, not about Excel.** COM objects use a
different PowerShell adapter, and the rejection actually observed on Windows
arrived as a catchable `COMException`. The wrapper nevertheless refuses a read that
neither raised nor answered, so if Excel's adapter ever swallows one, the operation
is still named at the point it happened.

---

## The diagnostic run that named the rejected call — `saveas.xlsm`

**Harness commit:** `3d34b26`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed.

### The instrumentation answered the question it was built for

```
[PASS] Read Stage-A build outputs
[PASS] Open an owned Excel instance
[PASS] Open the Stage-A workbook
[FAIL] Stage-B build
       operation=saveas.xlsm;
       System.Runtime.InteropServices.COMException: Call was rejected by callee.
       HRESULT 0x80010001 RPC_E_CALL_REJECTED
COMREJECT|build|none|attempts=0|waited=0
```

**The exact rejected BUILD operation is `saveas.xlsm` — `Workbook.SaveAs`.**

The `COMREJECT|build|none|attempts=0|waited=0` line matters as much as the label:
**no read was reissued**, so the four opted-in property gets had nothing to do
with this rejection. The refusal is in the save itself, and it happens before
worksheet acquisition, `VBProject`, `VBComponents`, the module import, the
`ThisWorkbook` write, the buttons, protection, the final `Save`, and any Bulk
fixture construction.

Endpoints completed and its real `PCCM_Calculate` ran. Shutdown was clean —
`Workbook.Close=True`, `Application.Quit=True`, natural PID exit, no emergency
cleanup, COM releases clean.

```
FAIL|Bulk|BOOTSTRAP|...
EQUIV|<not evaluated>|invalid|comparison was not executed: only the Endpoints
                              pass completed
```

No Bulk fixture was built. No comparison ran. This record
**must not be read as a fixture DIFFER**.
It is **INVALID / NOT EVALUATED**, and **Bulk remains NOT authorised**.

Runs 3 and 4 stay historically unresolved at exact-call level; the log that would
have named them did not exist yet. This run establishes the *current, repeatable*
rejection.

### The diagnostic wording, corrected

The line the build emits says *"the call was refused before it ran"*. That is what
the OLE message-filter **contract** says, and for a property get it is the whole
answer. **It is not proof that `SaveAs` made zero state change**, because `SaveAs`
is **non-idempotent**: it writes a file and rebinds the workbook. The contract is
therefore no longer used as the licence to reissue it. The licence is **observed
state**.

### Why the observation is conclusive here

The build **deletes the target immediately before the call**, so at the moment
`SaveAs` is attempted the target provably does not exist and the workbook is
provably bound to the Stage-A path in the Stage-A format. Three independent facts
then separate the outcomes:

| fact | read how |
|---|---|
| which file the workbook is bound to | `Workbook.FullName` |
| what format it is in | `Workbook.FileFormat` |
| whether the target is on disk | `Test-Path` |

Both COM reads go through the **accepted read-only retry helper**, so an
inspection that is itself refused is reissued rather than mistaken for a finding.

| state | requires | action |
|---|---|---|
| **COMPLETED** | bound to target **and** format 52 **and** target exists | accept; **never reissue** |
| **NOT EXECUTED** | bound to source **and** source format **and** target absent | one bounded retry permitted |
| **AMBIGUOUS** | anything else, including an unreadable inspection | **abort**; no retry, no cleanup |

The source format is **read before the call**, not assumed — "nothing moved" is
only provable against what the workbook was.

### The bounded algorithm

```
attempt SaveAs
  returned normally  -> verify all three facts; not COMPLETED is a FAILURE
  non-retryable error-> rethrow immediately, no inspection
  refused (0x80010001 / 0x8001010A)
        -> inspect postconditions BEFORE any second call
              COMPLETED     -> accept, do not reissue
              NOT EXECUTED  -> bounded backoff, then retry
              AMBIGUOUS     -> abort with the evidence intact
```

Bounds are the **accepted envelope's own values** — 12 attempts, 250 ms rising to
2000 ms, 15000 ms total — not new numbers. Exhaustion rethrows the **original**
rejection.

### Diagnostics

```
SAVEAS|attempt=1|success
SAVEAS|verified|path=True|format=52|exists=True

SAVEAS|attempt=1|rejected|0x80010001
SAVEAS|postcondition|not-executed|boundToTarget=False|boundToSource=True|format=51|targetExists=False|sourceExists=True|readError=none
SAVEAS|attempt=2|success

SAVEAS|postcondition|completed|...
SAVEAS|attempt=1|completed-despite-rejection

SAVEAS|postcondition|ambiguous|...
SAVEAS|attempt=1|error|0x800a03ec
```

### Executed, not asserted

`tests/phase10_stage_b_build_ops_flow.ps1` drives the real settlement against a
fake whose `SaveAs` is a **method** (so its exception really propagates) and which
can leave any partial state a half-done save could leave. Excel is never started.

| case | SaveAs calls | outcome |
|---|---|---|
| clean first attempt | **1** | completed, all three verified |
| refused, NOT EXECUTED | **2** | completed on the retry |
| refused `0x8001010A`, NOT EXECUTED | **2** | completed on the retry |
| refused but COMPLETED | **1** | accepted, **never reissued** |
| target exists, still bound to source | **1** | AMBIGUOUS, abort |
| rebound, no file, old format | **1** | AMBIGUOUS, abort |
| format moved, nothing else | **1** | AMBIGUOUS, abort |
| rebound + file, wrong format | **1** | AMBIGUOUS, abort |
| rebound + format, no file | **1** | AMBIGUOUS, abort |
| inspection unreadable | **1** | AMBIGUOUS, abort |
| `0x800A03EC` (accepted and failed) | **1** | rethrown, **no inspection** |
| not a COM failure | **1** | rethrown, no inspection |
| always refused, 3 attempts allowed | **3** | original `0x80010001` rethrown |
| always refused, 5 ms budget, 50 allowed | **3** | original rethrown |

Path comparison was proved too: a separator difference is **not** a rebind, and
two genuinely different paths are **not** equal.

### What is still NOT known, and is not claimed

**We know WHERE the rejection occurs. We do not know WHY Excel rejects the
second-session `SaveAs`.** No claim is made about lifecycle overlap, an Excel
readiness race, OneDrive, file locking, modal state or message-filter timing.
None of those has been independently proved, and this batch settles
**safe recovery, not root cause**. If the next run shows a `COMPLETED` or `NOT EXECUTED` settlement the
build proceeds; if it shows `AMBIGUOUS`, that is new evidence and a separate
decision.

### Still not done

No readiness gate. No inter-pass drain. No blanket sleep — the one sleep in the
bootstrap is the SaveAs backoff, reachable only after a postcondition **proved**
the save did not happen. No retry on the module import, the `ThisWorkbook` write,
button creation, protection or the final `Save`. No retry on any property SET.
`Invoke-ComRetryRead` byte-identical. Two bundles, two Excel sessions.

---

## Equivalence run 5 — INVALID / NOT EVALUATED — THE PRE-SAVE READ ANSWERED WITH NOTHING

**Harness commit:** `cc9cf8d`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed.

### What run 5 did

```
[PASS] Read Stage-A build outputs
[PASS] Open an owned Excel instance
[PASS] Open the Stage-A workbook
[FAIL] Stage-B build
       operation=saveas.xlsm;
       System.Management.Automation.RuntimeException:
       Invoke-StageBBuildRead: the Stage-A workbook FileFormat before SaveAs
       answered with nothing at saveas.xlsm.
       The read was not refused and it was not answered.
COMREJECT|build|none|attempts=0|waited=0
```

Endpoints completed. Shutdown was clean — `Workbook.Close=True`,
`Application.Quit=True`, natural PID exit, no emergency cleanup.

```
FAIL|Bulk|BOOTSTRAP
EQUIV|<not evaluated>|invalid
```

**`SaveAs` itself was NOT executed.** The failure was in the **pre-save FileFormat
observation**, before the call, and **no RPC rejection occurred** — `attempts=0`. No
Bulk fixture was built and no comparison ran, so this record
**must not be read as a fixture DIFFER**,
and it is **not** another SaveAs rejection. It is
**INVALID / NOT EVALUATED**, and **Bulk remains NOT authorised**.

It therefore tests none of: SaveAs recovery, Bulk fixture construction, the
reserved-row correction, semantic equivalence, `CALCEQUIV`.

### Root cause, narrowed by source elimination

The read went through `Invoke-StageBBuildRead`, a wrapper that took the COM object
and the member name and forwarded both to the accepted helper. `attempts=0` is what
makes the rest diagnosable — the retry loop answered on its **first** try, so the
reissue path, the rejection ledger and the backoff are all excluded. The single
`$value = $Target.$Member` in the accepted helper returned `$null` **without
raising**.

| candidate | verdict |
|---|---|
| telemetry polluting the return | **excluded** — the only emission was `return $record`; both side effects were `$null =`-suppressed |
| telemetry consuming the value | **excluded** — the branch was gated on `Attempts -gt 1`, and `attempts=0` proves it never ran |
| type coercion | **excluded** — the `[int]` cast was *outside* the wrapper and never reached |
| null/empty classification | **excluded as the cause** — it is the *report*; `$record.Value` was genuinely `$null`, because `$value` has exactly one assignment |
| wrapper output handling | **excluded for pollution**; the wrapper returned the helper's own object |
| dynamic member access through the extra hop | **the only surviving difference on the value path** |

**Why dynamic member access answered with nothing through that extra hop is NOT
established, and nothing here theorises about it.** The earlier PowerShell 7
observation about .NET property getters was explicitly *not* Windows evidence and
is not offered as one.

**What IS established is which form has worked.** The reopen verification reads
`FileFormat`, `Worksheets`, `VBProject`, `VBComponents`, `CodeName`, `Count`,
`Item`, `Shapes`, `OnAction` and `Name` through the accepted helper called
**directly**, and has done so on Windows in every accepted run. `Invoke-StageBBuildRead`
had **never once returned a value on Windows** — in runs 3 and 4 the SaveAs was
refused before any wrapped read ran, and run 5 is its first and only execution.

### The correction: the unproven layer is removed, not repaired

Every build read is now the proven form verbatim:

```powershell
Set-StageBBuildStep 'saveas.presave.fileformat'
$preFormat = Invoke-ComRetryRead -Target $wb -Member 'FileFormat' `
                 -Description 'the Stage-A workbook FileFormat before SaveAs'
Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $preFormat
$sourceFormat = Get-StageBScalarInt -Value $preFormat.Value `
                    -What 'the Stage-A workbook FileFormat before SaveAs'
```

The label is set **before** the call, the telemetry is recorded **after** it, and
both are separate statements — neither is in the expression that produces the
value, so neither can pollute or consume it. A COM object is taken straight off
`$read.Value` with no further parameter binding. **No third read mechanism was
invented**, and `Invoke-ComRetryRead` is byte-identical.

### Validation is at the consumer, by type, never by truthiness

`if (-not $value)` would refuse a legitimate `0`, which is a real value elsewhere
in the object model. So `FileFormat` requires a **scalar integer-compatible**
answer and `FullName` a **non-empty string**; a `$null`, an array, a blank string
and a non-numeric string are each refused with their own message.

### The pre-save observation contract

Before `SaveAs` is attempted the build reads the original `FullName` and
`FileFormat` and confirms the target is absent, then **requires the baseline to be
internally consistent**: the workbook must be bound to the Stage-A path, and the
target must not already exist. If either fails, **the save is not attempted** —
because `bound to source and target absent` cannot mean what NOT EXECUTED needs it
to mean without it. Run 5 is what an unchecked baseline looks like: `[int]$null`
would have become `0`, and every later NOT-EXECUTED verdict would have been wrong.

```
SAVEAS|baseline|fullname=ok|format=51|target-absent=True
```

### Precise sub-operation diagnostics

`saveas.xlsm` remains the region's name; the failure line now carries the exact
step:

```
saveas.presave.fullname   saveas.presave.fileformat   saveas.presave.target
saveas.call
saveas.post.fullname      saveas.post.fileformat      saveas.post.target
```

```
COMFAIL|build|saveas.presave.fileformat|hresult=none|…
```

A misspelt step is refused where it is written, and a new top-level operation
clears the step so a stale one cannot name a finished phase.

### The SaveAs settlement is unchanged

COMPLETED = bound to target **and** format 52 **and** target exists. NOT EXECUTED
= bound to source **and** the original source format **and** target absent.
AMBIGUOUS = everything else, never retried, never cleaned up. A retryable rejected
`SaveAs` may be reissued **only** after NOT EXECUTED is observed.

### Executed, not asserted

The harness drives the real read path: `51` and `52` come back as `Int32`, `"52"`
coerces to `52`, a `FullName` string survives as `String`, an object-valued read
survives as an object, and `$null`, a blank string, two values and `"xlsm"` are
each refused. The telemetry recorder emits **nothing** and leaves the value
unchanged (`51` before, `51` after); a read reissued twice returns its value with
`attempts=3` and records exactly one line naming `saveas.presave.fileformat`; an
exhausted read rethrows the original `0x80010001` and records nothing. Excel is
never started.

### Still not known

**Why Excel refuses the second-session `SaveAs` remains unknown**, and nothing here
claims lifecycle overlap, an Excel readiness race, OneDrive, file locking, modal
state or message-filter timing. Run 5 did not reach the save at all, so it adds no
evidence either way.

---

## Equivalence run 6 — INVALID / NOT EVALUATED — THE OPENED WORKBOOK DID NOT ANSWER YET

**Harness commit:** `c66e752`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed.

### The first Excel session was clean

```
SAVEAS|baseline|fullname=ok|format=51|target-absent=True
SAVEAS|attempt=1|success
SAVEAS|verified|path=True|format=52|exists=True
```

No build COM rejection occurred. The Endpoints fixture completed and the real
`PCCM_Calculate` ran.

### The second, isolated Bulk session did not

```
[PASS] Read Stage-A build outputs
[PASS] Open an owned Excel instance
[PASS] Open the Stage-A workbook
[FAIL] Stage-B build
       operation=saveas.presave.fullname;
       the Stage-A workbook FullName before SaveAs answered with nothing.
       The read was not refused and it was not answered.
COMREJECT|build|none|attempts=0|waited=0
```

Immediately after `Workbooks.Open`, before `SaveAs`. Shutdown was clean.
**`SaveAs` itself was NOT executed.** No Bulk fixture was built and no comparison
occurred. This record **must not be read as a fixture DIFFER**; it is
**INVALID / NOT EVALUATED**, and **Bulk remains NOT authorised**.

### The cross-run comparison, which is what licences the gate

| run | first post-open pre-SaveAs read | result |
|---|---|---|
| `cc9cf8d` | `FileFormat` | answered with nothing |
| `c66e752` | `FullName` — the wrapper had been removed, so this became first | answered with nothing |

Both with `attempts=0`, so **no RPC rejection was recorded** in either. And in the
**same** `c66e752` run the **first Excel session** observed `FullName` **and**
`FileFormat` successfully and completed its `SaveAs`.

So the evidence no longer supports a member-specific or wrapper-specific diagnosis.
What Windows shows is a **post-open readiness gap**: `Workbooks.Open` may return a
Workbook RCW in the second isolated session **before its own read-only properties
reliably answer**.

**That says WHERE the gap is observed. It does not say WHY.**

### The gate

Immediately after `$workbooks.Open($stageAPath)`, on the object it returned, before
any read, any baseline and any mutation. A workbook is READY only when:

| | observation |
|---|---|
| A | `FullName` is a non-empty scalar string |
| B | its normalised form equals the expected Stage-A path |
| C | `FileFormat` is a non-null scalar integer-compatible value |

The observed `FileFormat` **becomes** the original source format the SaveAs
NOT-EXECUTED verdict is measured against. It is **not** re-read afterwards, and no
literal `51` appears anywhere in the gate.

**No answer, or a refusal the accepted helper could not resolve within its own
bounds, means NOT READY YET** — poll again. Bounds are the accepted envelope's own
values: 12 attempts, 250 ms rising to 2000 ms, 15000 ms total. **A real but wrong
answer is not a delay**: a non-empty `FullName` that is the wrong path, or a
`FileFormat` that is not an integer, **aborts immediately** — waiting cannot change
which workbook this is. A non-retryable exception aborts on the first observation.

On exhaustion Stage-B **stops before `SaveAs`** — no call, no partial build, no
fixture pass, no comparison.

**Nothing sleeps on the way past a workbook that answers.** The success path breaks
out before the backoff, so a first-attempt readiness costs zero milliseconds.

```
READY|open|attempt=1|fullname=True|fileformat=51|waited=0
READY|open|attempt=3|fullname=True|fileformat=51|waited=2
READY|open|exhausted|attempts=12|waited=15000|fullname=no-answer|fileformat=ok
```

### Executed, not asserted

| case | outcome | attempts | waited | reads |
|---|---|---|---|---|
| answers immediately | READY | 1 | **0** | 1 + 1 |
| `FullName` silent twice | READY | 3 | >0 | 3 + 3 |
| `FileFormat` silent twice | READY | 3 | >0 | 3 + 3 |
| both silent three times | READY, format **52** observed | 4 | >0 | 4 + 4 |
| refused twice, then answers | READY | 3 | >0 | 3 |
| refused, then silent, then answers | READY | 3 | >0 | 3 |
| **wrong path** | **ABORTED** | — | — | **1 + 0** |
| **`FileFormat` = "xlsm"** | **ABORTED** | — | — | 1 + 1 |
| **non-retryable `0x800A03EC`** | **ABORTED**, code intact | — | — | **1** |
| 3 attempts allowed, always silent | ABORTED, exhausted line | 3 | 2 ms | 3 |
| 50 allowed, 5 ms budget | ABORTED | 3 | 4 ms | 3 |

A property getter cannot raise on the harness host — it is swallowed and the value
returns `$null` — so the refusal and real-error paths are driven through a scripted
stand-in for the accepted helper whose **contract** is what the gate depends on. An
earlier draft "threw" a rejection from a fake property and therefore exercised the
**null** path while claiming to test refusals; that case was replaced rather than
kept. Excel is never started.

### Unchanged

`Invoke-ComRetryRead` byte-identical. The SaveAs three-state settlement untouched,
including its postcondition inspection on **both** the refused and the normal-return
paths — **readiness success is not evidence that `SaveAs` will succeed**. No
inter-pass drain, no sleep between sessions, no PID quiet period, no session
spacing, no merging of the two sessions, no generic mutation retry.

### What is still NOT known

**Why the opened workbook does not answer for itself yet is not established.** No
claim is made about a message-filter race, modal state, OneDrive, a file lock,
lifecycle overlap, COM marshaling, or an Excel bug. The gate is a bounded
observation of a condition Windows has now shown twice; it is not a theory about its
cause.

---

## Equivalence run 7 — INVALID / NOT EVALUATED — TWO HARNESS DEFECTS

**Harness commit:** `f006ea2`. Windows PowerShell 5.1. Stage A immediately before
the run: 351 passed, 0 failed. Both starting bundles were built and proved
**identical**.

### What the two new settlements did

The post-open readiness gate worked in **both** Stage-B builds:

```
Endpoints  READY|open|attempt=1|fullname=True|fileformat=51|waited=0
Bulk       READY|open|attempt=1|fullname=True|fileformat=51|waited=0
```

`SaveAs` succeeded cleanly in **both**:

```
SAVEAS|attempt=1|success
SAVEAS|verified|path=True|format=52|exists=True
```

and no build COM rejection occurred anywhere:

```
COMREJECT|build|none|attempts=0|waited=0
```

**The post-open and SaveAs problems did not recur.** Neither settlement is
redesigned by this batch.

### Defect A — the reopen verification acquired a null Worksheets collection

The Endpoints Stage-B **build** completed in full: SaveAs, 14 CodeNames, all 32
modules, the ThisWorkbook module, 11 buttons, protection applied, final `Save`, and
the build instance closed naturally. Then the **reopen verification** failed every
sheet and every button with:

```
Invoke-ComRetryRead: no target object for Worksheets.Item(...)
```

beside `no verification read was refused; 0 ms waited`, and a shutdown ledger
reading `Worksheets2 | SKIPPED | reference was already null` while `Workbook2`,
`VBProject2` and `VBComponents2` all released cleanly.

**Cause, from source.** `$worksheets2 = (Invoke-ComRetryRead ... ).Value` took
whatever the helper read. To the helper a `$null` Value is a **success** — it read
the member and that is what came back — so the acquisition "succeeded" with nothing
in it and the first `Item()` call raised instead. The build path gained a non-null
guard in an earlier batch; **the verification path did not.**

**This is a verification-path no-answer defect. It is not a build failure and not a
production failure.**

**Correction.** A narrowly-scoped acquisition rule: a reopened object is usable only
when a **real object** comes back. `$null` means **not ready yet** and is retried
under the accepted envelope — 12 attempts, 250 ms rising to 2000 ms, 15000 ms total,
sleeping only after a no-answer or a refusal the helper could not resolve. A real
but **non-object** answer aborts, because waiting cannot turn an `Int32` into a
collection. A non-retryable exception aborts on the first observation. Applied to
all three reopened acquisitions — `Worksheets`, `VBProject`, `VBComponents`; scalar
reads and `Item` lookups are untouched.

**And it returns a record, never the object.** `Worksheets` is a collection, and a
collection written to the output stream is enumerated into its members — the Phase-10
Run-3 record collapse, and the very same hazard as Defect B below.

```
VERIFYREADY|Worksheets|attempt=1|ready=True|waited=0
VERIFYREADY|Worksheets|exhausted|attempts=12|waited=15000
```

### Defect B — the cost-profiling block lost its rectangular rank

The Bulk Stage-B bootstrap completed **fully**, including its reopen verification —
14 CodeNames, 32 modules, 11 buttons, clean COM lifecycle. Then Bulk fixture
construction raised:

```
FAIL|Bulk|RAISED|the bulk write for tblCostProfiling was handed a rank-1 array;
                 Excel accepts only a rectangular two-dimensional block
```

**Cause, from source.** `New-BenchmarkWeightBlock` ended with `return $block` on a
rank-2 `object[,]`. A multidimensional array written to the PowerShell output stream
is **enumerated into its elements** in row-major order, so the caller's assignment
received an `Object[]` of rows × years scalars. `New-BenchmarkRegisterBlock` never
had the defect because its matrix travels as a **property of a record**, which the
pipeline cannot enumerate — which is why the registers always wrote correctly.

`Set-BenchmarkRangeBlock`'s rank guard is what caught it **before** the write.

**This is a harness fixture-shape defect. It is not a production failure and it is
not a fixture DIFFER.**

**Correction.** The weight builder hands the matrix back the way the register
builder always has — as `.Block` on a record, alongside the `Rows`, `Columns` and
`Keys` it built — and both callers now assert the geometry they asked for before the
write. `return ,$block` would also work; a comma is one keystroke from being tidied
away by someone who does not know what it is holding up.

**The complete audit:** exactly two `object[,]` allocations and exactly two
`Set-BenchmarkRangeBlock` callers exist in the runner — the two registers and the
two profiling grids. The inflation grid is production's own via `SyncProfileRows`
and is not block-written.

### Executed, not asserted

| block | CLR type | rank | dimensions |
|---|---|---|---|
| register SMALL cost lines | `System.Object[,]` | 2 | 12 × 11 |
| register SMALL risks | `System.Object[,]` | 2 | 8 × 10 |
| register LARGE cost lines | `System.Object[,]` | 2 | 180 × 11 |
| weights SMALL cost profiling | `System.Object[,]` | 2 | 12 × 5 |
| weights SMALL risk profiling | `System.Object[,]` | 2 | 8 × 5 |
| weights LARGE cost profiling | `System.Object[,]` | 2 | 180 × 30 |
| **bare `return` (the defect)** | **`System.Object[]`** | **1** | **12 elements from 3 × 4** |
| **record property (the fix)** | `System.Object[,]` | 2 | 3 × 4 |

The last two lines are the same matrix handed back two ways — the defect and the fix
side by side, observed rather than claimed. The verification acquisition is proved
the same way: an object answers on attempt 1 with zero wait and comes back as
**one collection with three members**, a `$null` causes another bounded attempt, a
refusal retries, and a scalar, a string, a real error, the attempt bound and the
wait budget each abort. Excel is never started.

### Overall status

Endpoints never reached fixture execution, because Stage-B verification failed. Bulk
completed Stage-B but its fixture construction failed before completion. Therefore:
**no Endpoints snapshot, no Bulk snapshot, no semantic comparison, no CALCEQUIV.**

```
Equivalence run 7: INVALID / NOT EVALUATED
```

This record **must not be read as a fixture DIFFER**.
**Bulk remains NOT authorised.**
Both findings are **harness defects**; neither is a production failure and neither
is a build failure.

### Unchanged

The readiness gate, the SaveAs three-state settlement, `Invoke-ComRetryRead`,
`Set-BenchmarkRegisterRowCount`, `Get-BenchmarkPermanentId`,
`New-BenchmarkRegisterBlock`, `Set-BenchmarkRangeBlock`, the equivalence gate,
`Get-EquivalenceSnapshot`, the protection window, the timed path and production VBA.
No inter-pass drain, no sleep between sessions, no merged sessions.

---

## Equivalence runs 8 and 9 — INVALID / NOT EVALUATED — REFUSED INSIDE BULK FIXTURE CONSTRUCTION

**Harness commits:** run 8 at `6672b75` (the two harness corrections), run 9 at
`6672b75` again — the same command, run twice, failing at the same stage. Windows
PowerShell 5.1. Stage A immediately before each: 351 passed, 0 failed.

### The latest run, step for step

```
BUNDLE|identical|5 artifact(s)
```

**Endpoints:** post-open readiness on attempt 1 with 0 ms wait; `SaveAs` on
attempt 1; Stage-B build; reopen verification; fixture; **real `PCCM_Calculate`
ran**.

```
PASS|Endpoints|COMPLETED
```

**Bulk:**

```
READY|open|attempt=2|fullname=True|fileformat=51|waited=250
```

**The readiness gate absorbed the post-open no-answer gap** it was built for: the
first observation returned nothing, the gate waited 250 ms once, the second
observation answered, and the run continued. `SaveAs` on attempt 1. Stage-B build.
32 modules, 11 buttons and protection persisted. Reopen verification. Clean COM
lifecycle.

Then, **after Stage-B completed**, Bulk fixture construction raised:

```
FAIL|Bulk|RAISED|Call was rejected by callee. HRESULT 0x80010001 RPC_E_CALL_REJECTED
EQUIV|<not evaluated>|invalid|comparison was not executed: only the Endpoints pass completed
```

### What is proved, and what is not

Proved on Windows: the readiness architecture works; the SaveAs settlement path
works; Stage-B works for **both** passes; reopened verification works for **both**
passes. The repeated unresolved failure is now **inside Bulk fixture construction**.

**This is the SECOND consecutive rejection at the Bulk-fixture stage**, and the two
are identical.

**The exact Bulk fixture operation was NOT identified by the logging in place.** The
fixture ran inside one try/finally reporting one line for the whole region — the
Stage-B build's original defect, one stage later. **No claim is made here about
which fixture call was refused.** The instrumentation added in this batch exists so
the next run says so.

No Bulk snapshot. No semantic comparison. No `CALCEQUIV`. This record
**must not be read as a fixture DIFFER**.
**Bulk remains NOT authorised.**
It is **INVALID / NOT EVALUATED**. It is **not** a production failure, **not** a
Stage-B failure, **not** a readiness failure and **not** a SaveAs failure.

### The instrumentation, observational only

Bulk fixture construction now labels every COM and VBA call from a **closed
22-word vocabulary**, set immediately before the call it names. The gate prints the
label beside the failure, **separately** from and **after** the top-level line:

```
FAIL|Bulk|RAISED|Call was rejected by callee. ...
BULKFAIL|bulk.costprofiling.write|hresult=0x80010001|RPC_E_CALL_REJECTED (0x80010001)
```

A call Excel accepted and that then failed is a different finding and says so:

```
BULKFAIL|bulk.timeline.apply|hresult=0x800a03ec|not a refused call
```

**Captured at the throw, not read at the catch.** The fixture runs inside a
`finally` that closes the protection window, and that close is itself a labelled
operation — so an outer catch asking "which operation was in flight" would have been
told `bulk.window.close` every time. The label is saved in an inner catch before the
`finally` runs. This defect was found and closed while building the harness, before
any Windows run could have printed the wrong name.

**Nothing is retried, nothing sleeps, nothing about the fixture window changed.**
The window is opened and closed exactly where it was, by the same shim, with the
same rollback. Labels sit around calls; they are never inside the frozen functions
(`Set-BenchmarkRangeBlock`, `Set-BenchmarkRegisterRowCount`) and never inside
helpers the Endpoints fixture shares.

### Every Bulk fixture COM/VBA operation, in order, classified

A = read / property get / Item acquisition · B = idempotent value write ·
C = structural or non-idempotent mutation · D = `Application.Run` production endpoint

| # | label | what it does | class |
|---|---|---|---|
| 1 | `bulk.window.open` | `Run('P10FW_Begin')`, `Run('P10FW_State')` — the accepted shim | D |
| 2 | `bulk.registers.assert-empty` | `Get-TableBody` on both registers | A |
| 3 | `bulk.inputs.write` | four `Names.Item → RefersToRange → Value2 =` | A + B |
| 4 | `bulk.fx.reset` | `Set-TableCell` clears/seeds; `Get-NamedValue` reporting currency | A + B |
| 5 | `bulk.fx.write` | `Add-BlankTableRow` (a **search** of reserved rows, no `Add`) + `Set-TableCell` | A + B |
| 6 | `bulk.profiles.master` | `Set-TableCell` / reserved-row search | A + B |
| 7 | `bulk.cost.register.grow` | `Get-TableRowCount`; **`ListRows.Add()` only on shortfall** (0 for SMALL, 155 for LARGE) | A, then **C** if it grows |
| 8 | `bulk.cost.register.write` | `Worksheets/ListObjects/DataBodyRange/Cells/Resize` then **one rectangular `Value2 =`** | A + B |
| 9 | `bulk.cost.counter.write` | `Names.Item → RefersToRange → Value2 =` | A + B |
| 10–12 | `bulk.risk.register.grow` / `.write`, `bulk.risk.counter.write` | as 7–9, for `tblRiskRegister` | as 7–9 |
| 13 | `bulk.registers.readback` | `Get-TableBody` on both registers, reserved suffix proved blank | A |
| 14 | `bulk.timeline.apply` | `Run('PCCM_AutomationBegin')`, `Run('PCCM_ApplyTimeline')`, `Run('PCCM_AutomationResult')` | **D** |
| 15 | `bulk.timeline.coherence` | `Run('PCCM_StructuralReport')` | D (read-only report) |
| 16 | `bulk.inflation.rates` | `Set-TableCell` per rate | A + B |
| 17 | `bulk.costprofiling.acquire` | `Get-TableBody` on the grid production just keyed | A |
| 18 | `bulk.costprofiling.write` | range acquisition then **one rectangular `Value2 =`** | A + B |
| 19–20 | `bulk.riskprofiling.acquire` / `.write` | as 17–18 | as 17–18 |
| 21 | `bulk.final.coherence` | `Run('PCCM_StructuralReport')` | D |
| 22 | `bulk.window.close` | `Run('P10FW_End')` | D |

**If the next run names a class-A operation**, a bounded read retry may later be
considered. **If it names B, C or D — the block writes, the counter writes, the
register growth, or `PCCM_ApplyTimeline` — the evidence is returned and nothing is
retried** without a separate decision. No pre-authorisation exists.

### Executed, not asserted

`tests/phase10_bulk_ops_flow.ps1` lifts the real orchestrator and the real model,
identifier and block builders by AST, replaces each Excel-touching helper with a
recording stand-in, feeds it the **real** manifest, inspection and plan, and drives
it.

| case | result |
|---|---|
| PERF-SMALL, complete | all 22 labels in order; 32 helper calls; 0 grows; four blocks rank 2 |
| PERF-LARGE, complete | same 22 labels; **155 and 95 `ListRows.Add`**; 180×11, 120×12, 180×40, 120×40 |
| refused at each of 8 chosen operations | `BULKFAIL` names **that** operation, `hresult=0x80010001`, **no label touched after it**, each touched **once**, window still closed |
| accepted-and-failed at `bulk.timeline.apply` | `BULKFAIL\|bulk.timeline.apply\|hresult=0x800a03ec\|not a refused call` |
| misspelt label | **refused** where written; the label in effect unchanged |

Excel is never started. Nothing is retried: every helper under a failing label is
touched exactly once.

## Equivalence run 10 — THE FIRST COMPLETED COMPARISON — A REAL DIFFER ON TWO TRACE CELLS

**Harness commit:** `d90a186` (the Bulk operation labelling). Windows PowerShell 5.1.
Stage A immediately before: 351 passed, 0 failed.

### What happened, step for step

For the first time, **both equivalence passes completed** and the comparison ran.

```
BUNDLE|identical|5 artifact(s)
```

**Endpoints:** readiness, `SaveAs`, the Stage-B build, the reopen verification,
the fixture, and the real `PCCM_Calculate`:

```
PASS|Endpoints|COMPLETED|fixture built and PCCM_Calculate ran
```

**Bulk:** the second Excel session again needed one readiness poll — the gate
absorbed it exactly as designed, and this architecture is not reopened:

```
READY|open|attempt=2|fullname=True|fileformat=51|waited=250
PASS|Bulk|COMPLETED|fixture built and PCCM_Calculate ran
```

No Bulk RPC refusal occurred. No `BULKFAIL` line was emitted. The Bulk fixture
completed, so the operation map of runs 8 and 9 was not exercised by a failure.

### The comparison

Every compared field matched **except two**:

```
EQUIV|cost_profiling.body|differ
EQUIV|risk_profiling.body|differ
```

All headers matched. All weights matched. The reserved blank suffix matched. The
only visible differences are the **Description cells of the FINAL semantic driver**
in each profiling grid:

| grid | row | Endpoints | Bulk |
|---|---|---|---|
| Cost Profiling | `CL-012` | *blank* | `GateB CL-012` |
| Risk Profiling | `R-008` | *blank* | `GateB R-008` |

Every prior semantic row (`CL-001`–`CL-011`, `R-001`–`R-007`) carries its expected
description in **both** passes. The weight cells of every row, `CL-012` and `R-008`
included, are identical in both passes.

### The calculation

Both real production calculations succeeded, and the fingerprints are identical:

```
CALC|Endpoints|OK|Calculation committed.|CURRENT|2DA8A0F6092AEA4B
CALC|Bulk|OK|Calculation committed.|CURRENT|2DA8A0F6092AEA4B
CALCEQUIV|match|the production calculation fingerprint is identical
```

### Verdict

This run is a **REAL DIFFER**: a comparison that ran, on two passes that completed.
It is not INVALID and not NOT EVALUATED. The deterministic calculation inputs and
the calculation result state are equivalent; the remaining difference is confined
to the final-row profiling Description. **The equivalence contract requires
workbook-state equivalence, so equivalence is NOT yet accepted, the two differences
are not waived on the strength of `CALCEQUIV|match`, and Bulk remains NOT
authorised.**

### The cause, from source — not from the pattern

The Windows pattern (every row present except the last) suggested an
order-dependent synchronisation. That was an inference. Source confirms it:

1. `modDrivers.AddDriver` writes **only** the new permanent identifier, then runs
   `modProfiling.SyncRows Kind` (`src/vba/modDrivers.bas`, `AddDriver`: the ID
   write at `IdColumn(Kind)).Value = newId` precedes `modProfiling.SyncRows Kind`).
2. `modProfiling.SyncRows` rewrites the grid in register order and copies the
   register's trace column — `COL_COST_LINES_DESCRIPTION` / `COL_RISK_REGISTER_RISK_NAME`
   — into grid column 2, preserving every weight by permanent ID
   (`src/vba/modProfiling.bas`, `SyncRows`).
3. The accepted Gate-B fixture (`Invoke-Phase5AddDriverAndRequireSuccess`) invokes
   `PCCM_AddCostLine` / `PCCM_AddRisk`, proves the issued identifier, and **only
   then** writes that driver's business fields, Description included
   (`Write-Phase5Driver`). So the Description is written **after** the SyncRows that
   ran inside its own Add.
4. Driver N's trace is therefore refreshed by driver N+1's Add. The final driver's
   Add is followed by no further Add, and steps G–H of the fixture
   (`Write-Phase5InflationRates`, `Write-Phase5Weights`, the coherence report) run
   no synchronisation. Nothing refreshes `CL-012` or `R-008`.
5. The Bulk fixture writes every register row — Description included — as one
   block **before** its single `PCCM_ApplyTimeline`, whose `SyncRows` then sees
   every trace text. Its grids are fully synchronised.
6. There is no `Worksheet_Change` handler in production; the applied trace is
   copied at structural operations only. This is the accepted Phase-4 contract,
   `docs/phase4.md`: *"Trace columns are refreshed, not live. Description and Risk
   Name are copied into the profiling grids by synchronisation; editing the
   register updates them at the next structural operation, not instantly."*

This is a **fixture-order matter, not a production defect**: the interactive
workflow is Add, then type — the same order the fixture uses — and the accepted
contract says the trace follows at the next structural operation. The owners that
refresh it are `modProfiling.SyncRows` through Add, Delete, Apply / Update Timeline
and Repair Profiling. **No production change** is made or proposed for this.

### The canonical state

The profiling Description column is model-controlled trace metadata owned by
`modProfiling.SyncRows`. The one production-reachable canonical state is the
**fully synchronised** one: every user input entered, then the accepted structural
operation completed. Bulk already reaches it. The Endpoints fixture is left exactly
as accepted (`phase5_gate_b_scenarios.ps1` is not edited) and is now followed, in
the runner and in the gate, by **one more `PCCM_ApplyTimeline`** — the accepted
public structural command both fixtures already rely on — inside the fixture
window, inside the setup stopwatch, before the timed loop, before the snapshot.
With the entered timeline unchanged it deletes and adds no column, `SyncRows`
preserves every weight by permanent ID, `SyncProfileRows` preserves every rate by
profile and year, and the trace columns are copied from the now-complete
registers. No description is written by any fixture; the Bulk final descriptions
are not blanked to imitate the old artifact.

### Executed, not asserted

`tests/phase10_profiling_sync_flow.ps1` lifts the accepted fixture and the bulk
builder by AST over an emulation of `AddDriver` and `SyncRows` written from the VBA.
On PERF-SMALL: the accepted fixture alone leaves `CL-012` and `R-008` blank with
every prior row present — the run-10 pattern, reproduced from source; with the
resynchronisation all twelve and all eight descriptions are present; the corrected
Endpoints grids are byte-identical to the Bulk grids in the snapshot's own
rendering; every weight equals the model's in all three scenarios; the reserved
suffix is 13 and 17 blank rows in all three; no fixture wrote a grid column below
the first year column.

### What the next run must show

```
EQUIV|cost_profiling.body|match
EQUIV|risk_profiling.body|match
CALCEQUIV|match
```

with every other `EQUIV` line `match`. Both are required; neither excuses the
other. Until then, Bulk remains NOT authorised.

## Equivalence run 11 — INVALID / NOT EVALUATED — REFUSED BEFORE THE FIRST BULK LABEL

**Harness commit:** `1eb6395` (the profiling resynchronisation). Windows PowerShell
5.1. Stage A immediately before: 351 passed, 0 failed.

### What happened, step for step

```
BUNDLE|identical|5 artifact(s)
```

**Endpoints:** Stage-B completed; the fixture completed; the real `PCCM_Calculate`
ran:

```
PASS|Endpoints|COMPLETED|fixture built and PCCM_Calculate ran
```

**Bulk Stage-B:** readiness PASS on attempt 1; `SaveAs` PASS on attempt 1; 14
CodeNames, 32 modules and 11 buttons persisted; reopen verification PASS; clean COM
lifecycle. Then the Bulk pass raised:

```
FAIL|Bulk|RAISED|Call was rejected by callee. HRESULT 0x80010001 RPC_E_CALL_REJECTED
BULKFAIL|<before the first bulk operation>|hresult=0x80010001|RPC_E_CALL_REJECTED (0x80010001)
```

The rejected call was therefore **outside the labelled Bulk operation region**: the
instrumentation in place **did not cover** it. No Bulk fixture completion, **no
snapshot, no EQUIV, no CALCEQUIV**.

### Verdict

**INVALID / NOT EVALUATED.** This is **NOT a fixture DIFFER**: no comparison ran.
It **does not invalidate** run 10, which completed both passes and proved the
profiling-Description difference and its source cause; it means only that the
`1eb6395` correction has not yet received a completed Windows equivalence run.
**Bulk remains NOT authorised.**

### Why the label could not name it — from source

`Invoke-EquivalencePass` set its first label, `bulk.window.open`, immediately
before `Open-BenchmarkFixtureWindow` — correctly for that call — but only there.
Between the Stage-B bootstrap's return (`$bootstrapExit`, `Test-Path`) and that
line, the pass makes nine Excel calls with no label in flight:

| # | statement | what Excel is asked |
|---|---|---|
| 1 | `$excel = New-Object -ComObject Excel.Application` | start a new Excel instance |
| 2 | `$excel.Visible = $false` | property set |
| 3 | `$excel.DisplayAlerts = $false` | property set |
| 4 | `$excel.AskToUpdateLinks = $false` | property set |
| 5 | `$workbooks = $excel.Workbooks` | property get |
| 6 | `$wb = $workbooks.Open($stageB)` | open the freshly built workbook |
| 7 | `$excel.Run('PCCM_AutomationBegin', $true, '')` | production VBA |
| 8 | `Save-Phase5LockedFxSeed` | `Worksheets`, `ListObjects`, `Value2` reads |
| 9 | `Import-BenchmarkFixtureWindow` | `VBProject`, `VBComponents`, `Import`, `Name`, `Run('P10FW_Ping')` |

The failure was raised in one of those nine — the diagnostic proves that much and
no more. **It is not assigned to any call.** Nothing here claims the Excel start,
the workbook open, the automation envelope, the seed read or the shim import;
the next run's label will say which. No claim is made about readiness, timing,
lifecycle overlap or any other cause.

### The instrumentation, observational only

Every one of those calls is now named from the closed vocabulary immediately before
it runs, in the gate and in the runner (whose Bulk mode makes three further reads
of its own: the Excel identity, the environment inventory and the protection
state), and a catch around the prefix saves the label at the throw and rethrows,
exactly as the builder's catch does. The window open is inside that guard in the
gate. A refused property get can answer with nothing — Stage-B saw that on
Windows — so the `Workbooks` acquisition is followed by a null guard under its own
label, so that a swallowed refusal there is not reported one label later.
`Reset-BulkOp` now runs before the prefix, so the sentinel
`<before the first bulk operation>` can only name a failure that precedes the
first instrumentable operation.

| # | label | call | class |
|---|---|---|---|
| 1 | `bulk.preflight.excel.create` | `New-Object -ComObject Excel.Application` | A |
| 2 | `bulk.preflight.excel.identity` (runner only) | `Get-ExcelIdentity` — `Hwnd` read | A |
| 3 | `bulk.preflight.excel.visible` | `Visible = $false` | B |
| 4 | `bulk.preflight.excel.displayalerts` | `DisplayAlerts = $false` | B |
| 5 | `bulk.preflight.excel.asktoupdatelinks` | `AskToUpdateLinks = $false` | B |
| 6 | `bulk.preflight.workbooks.acquire` | `Workbooks` get, then the null guard | A |
| 7 | `bulk.preflight.workbook.open` | `Workbooks.Open` | **C** (a second Open is a second session state) |
| 8 | `bulk.preflight.environment.read` (runner only) | `Version`, `Build`, `Calculation`, `Workbooks.Count` | A |
| 9 | `bulk.preflight.automation.begin` | `Run('PCCM_AutomationBegin')` | D |
| 10 | `bulk.preflight.fxseed.read` | `Worksheets`/`ListObjects`/`Value2` reads | A |
| 11 | `bulk.preflight.window.import` | `VBProject`, `VBComponents.Import`, `Name`, `Run('P10FW_Ping')` | **C** + D |
| 12 | `bulk.preflight.protection.read` (runner only) | `Run('P10FW_State')` | D |
| 13 | `bulk.window.open` | `Run('P10FW_Begin')` and the state read | D |

A property set that Excel silently refused would take effect nowhere and raise
nothing; that case is not detectable by labelling alone and is recorded as such.

**If the next run names a class-A operation**, a bounded read retry may later be
considered. **If it names B, C or D**, the evidence is returned and nothing is
retried. **No pre-authorisation exists.**

### Executed, not asserted

`tests/phase10_bulk_ops_flow.ps1` now cuts the gate's Bulk pass (from its
`Reset-BulkOp` through the window close) and the runner's Bulk prefix (through its
catch) out of the real files by their first and last statements and runs them over
a fake Excel whose every property get, property set and `Run` records the label in
flight and can be told to refuse there; `New-Object` itself is stood in for so the
COM request returns the fake. A successful prefix meets the labels in order and
hands over to `bulk.registers.assert-empty`. Refused at each of the ten gate prefix
operations and each of the twelve runner prefix operations, `BULKFAIL` names that
operation with its HRESULT, no label is touched after it, each is touched once, and
the window is neither opened nor closed. The swallowed `Workbooks` get is named
under `bulk.preflight.workbooks.acquire` with `hresult=none` and the message *the
Workbooks collection read answered with nothing*.

The `1eb6395` resynchronisation is byte-identical and still called once in each
file, immediately after `Set-Phase5Fixture`, inside the window and before the timed
loop. `Get-EquivalenceSnapshot`, the window functions, the block builders,
`build_stage_b.ps1`, `com_lifecycle.ps1` and production VBA are byte-identical.

## Architecture simplification after run 11 — ONE Stage-B, TWO COPIES

**Not a Windows run.** This records a change of harness architecture made after
run 11, so the sections above remain accurate about what earlier runs executed
and this one is accurate about what the next run will execute.

### What changed

Runs 3 to 11 spent every failure inside work that was never the semantic
question: a second Stage-B bootstrap, a second `SaveAs`, a second module import,
a second protection pass, a second reopen verification, and finally the Excel
calls a second pass makes before its first fixture step. The question is only:
*starting from the exact same verified Stage-B workbook image, do the Endpoints
fixture and the Bulk fixture reach the same workbook state and the same
production calculation fingerprint?* That needs one Stage-B, not two.

`tests/phase10_fixture_equivalence.ps1` now:

1. prepares **ONE** starting bundle from the Stage-A build and runs the Stage-B
   bootstrap **once**, exactly as it ships — build, `SaveAs`, modules, buttons,
   protection, its own reopen verification, its own Excel shutdown;
2. takes that verified `.xlsm` as the canonical baseline and prints its SHA-256
   (`BASELINE|StageB|verified|sha256=…`);
3. makes **two filesystem copies**, `PCCM_equiv_endpoints.xlsm` and
   `PCCM_equiv_bulk.xlsm`, prints each SHA-256, and requires
   canonical == Endpoints copy == Bulk copy before any Excel is started
   (`COPIES|identical`, or `COPIES|differ|…` and nothing is opened);
4. opens each copy in its own clean Excel instance and, **after the open and
   before anything is asked to change** — before `PCCM_AutomationBegin`, the seed
   read, the shim import, the window, the fixture and the calculation — runs one
   bounded **read-only readiness barrier** (`READY|<mode>|attempt=N|waited=X`):
   `FullName` must be this copy, `Worksheets` must answer, one known worksheet
   must be acquirable; every read goes through the accepted `Invoke-ComRetryRead`
   envelope; an answer of nothing waits 250 ms doubling to 2000 ms, at most 12
   attempts and 15000 ms; a real answer naming the wrong workbook aborts at once;
5. runs the existing Endpoints fixture plus the `1eb6395`
   `Invoke-BenchmarkEndpointsResync`, or the existing Bulk fixture; the real
   `PCCM_Calculate`; the unchanged `Get-EquivalenceSnapshot`;
6. compares with the existing contract, every EQUIV field and `CALCEQUIV`
   independently required.

**No mutation is retried** anywhere: no property set, write, `ListRows.Add`,
import, `Application.Run`, window, protection, `Calculate` or `ApplyTimeline` is
wrapped in a loop or a sleep; there is no inter-pass drain and no process spacing.
If a mutation is refused after readiness succeeded, that is reported and stops
the run; no further instrumentation subsystem is built.

### What was retired

The `24e1a14` **prefix instrumentation is retired**: the twelve `bulk.preflight.*`
labels, the prefix try/catch in the gate and the runner, the `Workbooks` null
guard, the lifted-region harness section and their controls and mutations are
removed. The runner is byte-identical to `1eb6395` again. The single-baseline
gate has no duplicated prefix to name; the readiness barrier is the settlement.
The Bulk builder's 22 operation labels and the `BULKFAIL` line stay, because a
refused builder step is exactly what PERF-LARGE would need named.
`Test-BundleIdentity` (two bundles) is replaced by `Test-CopyIdentity` (three
digests). `Get-EquivalenceSnapshot`, `New-EquivalenceBundle`,
`Get-BundleArtifacts`, the window functions, the block builders,
`Invoke-BenchmarkEndpointsResync` and its two call sites, production VBA,
`build_stage_b.ps1` and `com_lifecycle.ps1` are byte-identical.

Run 11's section above records the prefix instrumentation as it stood at
`24e1a14`; that is history and is not rewritten.

### Acceptance

ONE successful run closes Bulk fixture equivalence if: the canonical Stage-B
build and verification passed; both copies are hash-identical to it; readiness
passed in both passes; both fixtures completed; every `EQUIV` field is `match`;
`CALCEQUIV|match`; the COM lifecycle is clean. Until then, **Bulk remains NOT
authorised**.

## Equivalence run 12 — INVALID / NOT EVALUATED — THE HARNESS CRASHED AT THE COMPARISON

**Harness commit:** `0a6df69` (one Stage-B, two copies). Windows PowerShell 5.1.

Under the single-baseline architecture, everything up to the comparison passed for
the first time:

```
BASELINE|StageB|verified
COPY|Endpoints|<same hash>
COPY|Bulk|<same hash>
COPIES|identical
PASS|Endpoints|COMPLETED|fixture built and PCCM_Calculate ran
PASS|Bulk|COMPLETED|fixture built and PCCM_Calculate ran
```

Then the harness raised, before any `EQUIV` line:

```
The property 'State' cannot be found on this object.
    foreach ($field in @($reference.State.Keys)) {
```

**INVALID / NOT EVALUATED** — no comparison ran; **NOT a fixture DIFFER**; Bulk
remains NOT authorised. Both fixtures completed and both real calculations ran.

### Cause, from source

`Invoke-EquivalencePass` wrote its `READY|…` line with `Write-Output` from inside
the function. A PowerShell function's output stream *is* its return value, so the
caller received a two-element array — the string and the result record — and
`.State` is not a property of that array. The two `SHUTDOWN|…` catches in the
pass's `finally` were the same trap, latent.

### Correction

The pass returns exactly one record; readiness (`Ready`) and any shutdown notes
(`Shutdown`) travel in it, and the caller prints `READY|…` and any notes from the
record before `PASS|…`. The caller also refuses a result that is not one record
carrying `State`. `Get-EquivalenceSnapshot`, the comparison block, the readiness
barrier, the canonical/copy architecture, the fixtures, the resync, production
VBA, the Stage-B builder and the timed paths are unchanged.
`tests/phase10_result_shape_flow.ps1` reads the record's keys from the pass's own
return statement, proves the pass emits nothing else, drives the real snapshot for
its real field set (28), and runs the real comparison block over two such
records: every field iterated, `CALCEQUIV` evaluated after the last `EQUIV`, in
both the identical and the differing case.

## Equivalence run 13 — ACCEPTED — BULK FIXTURE AUTHORISED

**Harness commit:** `8caffb0`. Windows PowerShell 5.1. Stage A immediately before:
**351 passed, 0 failed**.

```
BASELINE|StageB|verified|sha256=10E06F40BD856D15823F419381BC309F08A19F7726E490DF6242D0E3A0B09483
COPY|Endpoints|sha256=10E06F40BD856D15823F419381BC309F08A19F7726E490DF6242D0E3A0B09483
COPY|Bulk|sha256=10E06F40BD856D15823F419381BC309F08A19F7726E490DF6242D0E3A0B09483
COPIES|identical
READY|Endpoints|attempt=1|waited=0
PASS|Endpoints|COMPLETED|fixture built and PCCM_Calculate ran
READY|Bulk|attempt=1|waited=0
PASS|Bulk|COMPLETED|fixture built and PCCM_Calculate ran
```

The canonical Stage-B build and its reopen verification passed once; the
Endpoints and Bulk copies carried the identical SHA-256; the readiness barrier
passed on attempt 1 with 0 ms waited in both passes; both fixtures completed.

**EVERY EQUIV FIELD MATCHED**: cost-line identifiers and order; risk identifiers
and order; both counters; both register bodies; the FX table; the Config profile
master; the applied timeline names; the Cost Profiling, Risk Profiling and
inflation headers and bodies; `nmStructuralState`; `PCCM_StructuralReport`;
`PCCM_CurrentInputFingerprint`; `PCCM_CurrentSimulationRequestFingerprint`;
`PCCM_ModelCheckCalculationState`. Including the two fields run 10 had found
different:

```
EQUIV|cost_profiling.body|match
EQUIV|risk_profiling.body|match
```

Both real production calculations committed as CURRENT with the identical
fingerprint:

```
CALC|Endpoints|OK|Calculation committed.|CURRENT|2DA8A0F6092AEA4B
CALC|Bulk|OK|Calculation committed.|CURRENT|2DA8A0F6092AEA4B
CALCEQUIV|match|the production calculation fingerprint is identical
```

### Decision

**BULK FIXTURE EQUIVALENCE = CLOSED / ACCEPTED.** The Bulk fixture is
**AUTHORISED for performance benchmark construction**. The equivalence harness is
not reopened unless future evidence shows a genuine semantic defect. The default
`-FixtureMode` stays `Endpoints` — the method the accepted SMALL and MEDIUM
baselines were built by — and PERF-LARGE is run with `-FixtureMode Bulk`, which the
artifact records as `fixture_mode`.

### PERF-LARGE, from the current harness

Contract: PERF-LARGE = 180 Cost Lines + 120 Risks = 300 drivers, 40 project years,
iterations 10,000 and 50,000 (100,000 is forbidden for Large by the plan and the
runner refuses it). Setup, all outside the timed region: Stage-B bootstrap; Excel
start; workbook open; `PCCM_AutomationBegin`; FX seed; shim import; protection
read; fixture window open; the Bulk fixture — inputs, FX, profile master, both
register blocks with `ListRows.Add` growth 25 → 180 and 25 → 120, counters,
readback, **one** real `PCCM_ApplyTimeline` (year columns, profiling rows, inflation
rows, ValidateStructure), inflation rates, both weight blocks, final coherence;
window close; seed write. Then the timed matrix, in plan order, one cold run and
three warm runs each, median of the warm three: Calculate; workbook recalculation
(`Application.CalculateFull`); Simulation @ 10,000; Sensitivity @ 10,000; Annual
@ 10,000; Simulation @ 50,000; Sensitivity @ 50,000; Annual @ 50,000 — eight runs,
thirty-two timed executions.
