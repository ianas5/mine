# Phase 5 — Gate A — Step 7: transactional reporting and orchestration

**Status: correction round 2 — ready for independent review.**

Step 6 is accepted and closed at `12ade1d`. This step adds the final Phase-5
production module — `modCalcReport` — with the transactional write-back, the
`calc_state` maintenance and the six calculation endpoints, plus the static Linux
suite that reads it and the narrow plumbing the endpoint declaration needs.

The first submission (`a9de9b3`) was rejected with five blocking defects, all of
which independent review has since confirmed closed; they are recorded under
**[Correction round 1](#correction-round-1)**. Correction round 1 (`73adbb0`) was
then rejected with two further blockers — an uncontained normal-path cleanup and
a commit failpoint that was not at the commit boundary — recorded under
**[Correction round 2](#correction-round-2)**. Each defect keeps the tests that
would fail if it returned. The accepted Step-7 architecture is otherwise
unchanged; `modCalcFingerprint` was reopened once, in round 1, for exactly one
visibility change under explicit authorisation, and is untouched this round.

---

## What this step does NOT claim

1. **NO VBA WAS EXECUTED.** There is no VBA interpreter on Linux, and none was
   simulated.
2. **No transaction has been observed to commit, and no rollback has been
   observed to restore anything.** Every assertion in
   `tests/test_phase5_report_source.py` is a statement about source text: what is
   written where, in what order, inside which error envelope.
3. **The six-row status matrix is not demonstrated here.** Gate A proves the
   source is *capable* of every row and contains no ordering that makes one
   unreachable. Demonstrating them needs real Excel.
4. **No workbook was written.** No Windows run, no harness extension.
5. **The module has never been compiled.**

---

## Final Phase-5 architecture

Fifteen modules, asserted against a freshly emitted manifest (`test_02`):

| | |
| --- | --- |
| Phase 4 | `modConstants` *(generated)*, `modWorkbook`, `modAppState`, `modTimeline`, `modDrivers`, `modProfiling`, `modInflation`, `modStructuralCheck` |
| Phase 5 | `modCalcContract` *(generated)*, `modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint`, `modCalcResolve`, `modCalcCheck`, `modCalcReport` |

Generated remains exactly `{modConstants, modCalcContract}`. `test_03` refuses
`modCalcOrchestrator`, `modCalcState`, `modCalcTransaction`, `modCalcApi` and
three more names by which the work might have been split.

### Line metrics

| Module | Raw | Blank | Comment | Code | code < 900 | raw < 1200 |
| --- | --- | --- | --- | --- | --- | --- |
| `modCalcReport` | 1113 | 78 | 242 | 793 | yes | yes |

It fits, with 107 code lines of headroom. No extra module was needed, and none
was added in either correction round.

---

## The six endpoints

```vba
Public Sub      PCCM_Calculate()
Public Function PCCM_CalculationStatus() As String
Public Function PCCM_CalculationAttemptResult() As String
Public Function PCCM_CalculationAttemptDetail() As String
Public Function PCCM_CalculationFingerprint() As String
Public Function PCCM_CurrentInputFingerprint() As String
```

Exactly six, asserted in both directions (`test_04`). `PCCM_CalculationRefusal`
appears nowhere (`test_05`) — it was replaced by the attempt axis.

**No other Public procedure exists in the module** (`test_08`). Two Public
*constants* do: `FAILPOINT_ANALYTICAL_WRITE` and `FAILPOINT_SUCCESS_COMMIT`, and
they are public precisely because a later Gate-B harness arms a failpoint **by
name**. `test_08` asserts those two are the only ones.

### They are declared, and not one is a button

`spec/structure_contract.yaml` gains a third PCCM_ group, `api_procedures`.
`entry_points` are button-bound and `harness_procedures` are the Phase-4 harness
helpers; neither describes a calculation endpoint, and putting the six into
either would have made the vocabulary lie. The loader parses the new key, the
emitter projects it into the manifest, and `test_08_no_orphan_pccm_macro_exists`
now accounts for all three groups — the rule itself is unchanged.

`test_06` and `test_07` assert against a **freshly emitted manifest** that there
are exactly five buttons, that no endpoint is bound to one, and that no shape is
named for Calculate.

---

## One preparation pipeline, three consumers

`PCCM_Calculate`, `PCCM_CalculationStatus` and `PCCM_CurrentInputFingerprint`
all go through `PrepareCurrentCalculation` (`test_09`). There is exactly one
definition of "valid current inputs", so a state the write path would refuse can
never be reported CURRENT because a partial digest happened to be constructible.

The sequence, asserted by statement index (`test_10`):

1. `modCalcResolve.ResolveModel` — which itself invokes the Phase-4 structural gate
2. `modCalcCheck.CheckResolvedModel`
3. `BuildFactorTables` — inflation factors per referenced profile, discount factors once, `YearFactors`
4. `BuildDriverFactors` — one `DriverFactors` per driver, with `Knom` and `Kpv`
5. `BuildAudits` — `modCalcAnalytical.BuildDriverAudit`
6. `modCalcAnalytical.AccumulateTotals`
7. `BuildAnnual` — `modCalcAnalytical.BuildAnnualSeries`
8. `modCalcAnalytical.Reconcile`
9. `modCalcAnalytical.AllIdentitiesHold` — every identity must hold
10. `BuildFingerprint`

`test_13` asserts preparation writes nothing at all. `test_12` asserts a
preparation that produced no digest does not report success.

`test_14` pins the numerical surface to thirteen named calls into the three
accepted modules and refuses compounding, distribution arithmetic and any
fingerprint constant in the reporter. Nothing is reimplemented.

### The carry convention lives in one place

A cost line carries its Quantity and a Probability of 1; a risk carries a
Quantity of 1 and its Probability. `test_39` asserts `= 1#` appears in
`BuildDriverFactors` and **nowhere else** — not in the audit block, not in the
fingerprint record.

---

## Two axes that never mix

| Axis | Values |
| --- | --- |
| **Derived status** — what do CURRENT inputs say about the stored snapshot? | `NOT CALCULATED`, `CURRENT`, `STALE`, `INVALID` |
| **Attempt result** — what happened last time Calculate was attempted? | `NONE`, `SUCCESS`, `REFUSED`, `FAILED` |

`test_31` collects every value assigned to `DeriveStatus` and requires the set to
be exactly the four statuses. `test_32` refuses any read of
`CALC_STATE_ROW_LAST_ATTEMPT_RESULT` inside the derivation: a refusal last week
does not make today's matching fingerprint stale.

### The rule

```
if the current inputs do not prepare      -> INVALID
else if no stored successful fingerprint  -> NOT CALCULATED
else if current = stored                  -> CURRENT
else                                      -> STALE
```

**No empty digest is ever compared.** `test_33` asserts the invalid check
precedes the read of the stored digest, and the `Len(stored) = 0` guard precedes
the equality comparison. Two blanks are two absences, not a match.

`PCCM_CalculationStatus` writes **only C19:C20**, as one 2×1 assignment
(`test_24`), and touches no analytical block and no part of the last-success
record. Status is last-evaluated, not live; there is no change event.

### Status-matrix capability

Gate A cannot demonstrate the six rows, but it can show nothing forbids them:

| Row | What the source provides |
| --- | --- |
| 1 success, nothing changed | current digest equals stored → `CURRENT`; commit wrote `SUCCESS` and a blank detail |
| 2 fingerprinted input changed | digests differ → `STALE`; the attempt block is untouched by a status refresh |
| 3 invalid input, no Calculate | preparation fails → `INVALID`; the attempt block is untouched |
| 4 invalid input + Calculate | `RecordRefusal` writes `REFUSED` + detail + `INVALID`, touching C17:C20 only |
| 5 prior inputs restored after a refusal | status is re-derived from the current digest → `CURRENT`, while C17:C18 still hold the refusal |
| 6 injected write failure | rollback, then `RecordFailure` writes `FAILED` with a **freshly derived** status |

---

## The transaction

Asserted by statement index (`test_16`):

```
On Error GoTo PreWriteFailed
    PrepareCurrentCalculation           pure; nothing analytical is touched
    CaptureSnapshot
On Error GoTo 0
    If Not prepared Then RecordRefusal  REFUSED; no rollback, nothing mutated
On Error GoTo TransactionFailed
    WriteAnalytical                     five tables + calc_totals
    FailPointCheck FAILPOINT_ANALYTICAL_WRITE
    VerifyAnalytical                    read back and compare
    FailPointCheck FAILPOINT_SUCCESS_COMMIT
    BuildSuccessBlock                   ONE 8x1 block, ONE captured moment
    WriteSuccessCommit                  ONE assignment of THAT block to C13:C20
    VerifySuccessCommit                 all eight cells against THAT block
    committed = True
On Error GoTo 0
```

**The commit and its verification are inside the envelope** (`test_17`): the
assignment can fail and so can its verification, so both sit inside the rollback
boundary. Success is not published before the analytical snapshot verifies
(`test_18`). A failed verification raises rather than continuing (`test_19`).

Preparation and the snapshot sit in **their own** envelope. A *controlled*
refusal from the accepted machinery is `REFUSED`; an *unexpected* runtime fault
in the same region is `FAILED`, with no rollback because nothing was mutated
(`test_56`). Downgrading a runtime fault to a refusal would report a model
problem the user does not have.

There is no `On Error Resume Next`. `test_20` asserts the handler set is exactly
the six reviewed envelopes — `InvocationFailed`, `CleanupFailed`,
`PreWriteFailed`, `TransactionFailed`, `RollbackFailed`, `BookkeepingFailed` —
and that every armed label is defined in the procedure that arms it.

### Snapshot set

Five ListObjects through `modWorkbook.SnapshotTable`, plus `C23:C32` and
`C13:C20` as values (`test_26`). Values only — the snapshot captures no
`NumberFormat`, `Interior` or `ColumnWidth`, because labels, notes and formats
are build-owned.

### Rollback

All five tables through `modWorkbook.RestoreTable`, plus both scalar blocks
(`test_27`). No second rollback mechanism exists (`test_30`).

**Rollback happens first, metadata second** (`test_28`). The first observable
moment after a failure is the previous successful snapshot, exactly. Only then
does `RollbackAndRecord` write one 4×1 block to C17:C20 with `FAILED`, the
specific detail, a **freshly derived** status and a fresh timestamp
(`test_34`).

So the final observable state after a failure is:

* **C13:C16** — the previous success block, restored exactly
* **C17:C20** — `FAILED`, the new detail, the current derived status, a new timestamp
* **C23:C32** and the five tables — the previous successful snapshot, restored exactly

The document does **not** claim final C13:C20 equals its previous value. Only
C13:C16 does.

**If the restore itself fails, no failed-attempt metadata is written at all.**
That record asserts *"the previous snapshot stands"*, and writing it after a
failed restore would assert something nobody established. Both diagnostics — the
original failure and the restore failure — are carried to the caller instead
(`test_28`).

If the failure-metadata assignment itself fails, the already-successful rollback
is **not** undone; there is no recovery transaction around failure bookkeeping,
and `test_28` refuses a `RestoreSnapshot` in the bookkeeping handler.

`test_29` asserts a committed operation can never become `FAILED`: the envelope
is disarmed between `committed = True` and the procedure's exit, so no handler
can fire after the commit, and nothing on that path writes.

### Pre-write refusal boundary

A refusal happens before any analytical write (`test_22`). `RecordRefusal` writes
one 4×1 block to **C17:C20** and touches nothing else — `test_22` refuses
`CALC_STATE_VALUE_RANGE`, `CALC_TOTALS_VALUE_RANGE`, `WriteAnalytical` and
`WriteSuccessCommit` inside it.

**A refused calculation does not leave the workbook byte-identical.** The
attempt and status metadata deliberately change. C13:C16 and every analytical
block do not.

### Failpoints

Two, through the accepted Phase-4 `modAppState.FailPointCheck` (`test_21`):

| Constant | Stage name | Source location | Position |
| --- | --- | --- | --- |
| `FAILPOINT_ANALYTICAL_WRITE` | `Phase5AnalyticalWrite` | `RunCalculation`, `modCalcReport.bas:263` | after `WriteAnalytical`, before `VerifyAnalytical` — analytical state is half-written |
| `FAILPOINT_SUCCESS_COMMIT` | `Phase5SuccessCommit` | `WriteSuccessCommit`, `modCalcReport.bas:744` | the statement **immediately before** `CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block` |

The commit hook sits at the assignment, not upstream of it:

```vba
Private Sub WriteSuccessCommit(ByRef block As Variant)
    modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT
    CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block
End Sub
```

`WriteSuccessCommit` is called from inside the `TransactionFailed` envelope,
after the analytical snapshot has been written and verified, so an injected
exception there rolls back from the **final commit boundary** with C13:C20 not
yet published.

`test_21` inspects the writer's own executable body — the hook must be exactly
one statement before the assignment, with nothing between them and nothing after
it — and refuses the hook anywhere in `RunCalculation`, refuses a second wiring,
and refuses `gAutomationFailAfterStage`: no second injection framework. Phase-4
failpoint machinery is untouched, and the Windows harness was neither extended
nor run.

---

## The audit blocks

Every anchor, column ordinal and range comes from the generated
`modCalcContract`. No second copy of a `_Calc` coordinate exists, and no label,
note or number format is written during a calculation (`test_47`).

**`tblCalcYears`** — one row per applied project year, from the in-memory
`YearFactors`.

**`tblCalcInflationFactors`** — one row per referenced profile per calendar year
over `BaseYear .. LastProjectYear`. The Base-Year row has a **BLANK** annual rate
— a model-controlled blank, never a fabricated zero — and a cumulative factor of
1 (`test_40`). Where `BaseYear < StartYear` the pre-project compounding years
stay visible, which is why the span is audited rather than the project years.

**`tblCalcFX`** — referenced currencies only, with `Referenced By` counted by
exact binary key comparison. The global reporting-currency invariant creates no
audit row; `test_41` refuses `REPORTING_CURRENCY` anywhere in the module. An
empty driver set produces zero rows.

**`tblCalcDrivers`** — 21 columns, one row per driver. **An inapplicable field is
BLANK.** `test_38` requires each of the eight kind-specific columns to be
explicitly `Empty` on the side it does not apply to, and refuses any `= 1#` or
`= 0#` assignment into an audit column. Never the in-memory identity, never zero.

**`tblCalcAnnual`** — one row per applied project year, from the already-built
`AnnualRow` array.

**`calc_totals`** — the ten totals from `AnalyticalTotals`. `test_43` asserts all
ten come from memory and that the builder contains no `Range(` and no
`CalcSheet`. The worksheet is the record of a calculation, never an input into
one (`test_44`).

### Empty semantic tables

A physical ListObject placeholder is not a semantic record. Each block builder
returns `Empty` for a zero-row table and `WriteTable` clears the physical body
(`test_42`). No `Project Index 0`, no SAR row, no empty driver ID, no zero-valued
analytical row.

### Verification

A write is not proven by the absence of an error. `VerifyTable` reads the body
back and compares it cell by cell against the prepared block (`test_45`), and
`SameCell` matches a blank only against a blank — the blank check precedes the
numeric comparison, so **a blank never verifies as zero** (`test_46`).

---

## Application state

`PCCM_Calculate` captures, begins, runs, finishes and publishes through the
accepted Phase-4 discipline (`test_48`), and a failed cleanup is reported rather
than swallowed. No `MsgBox`, no change handler (`test_49`), no RNG or `_SimData`
(`test_50`). `modAppState` was not redesigned and is byte-frozen (`test_51`).

The endpoint is wrapped in a top-level envelope installed **before the first
fallible operation** — before `CaptureAppState`, before `BeginOperation` — with
an explicit `stateCaptured As Boolean`:

```vba
On Error GoTo InvocationFailed
state = modAppState.CaptureAppState()
stateCaptured = True
modAppState.BeginOperation
result = RunCalculation(committed)
On Error GoTo 0

On Error GoTo NormalCleanupFailed
cleanupAttempted = True
cleanup = modAppState.FinishOperation(state)
On Error GoTo 0
stateCaptured = False

If Len(cleanup) > 0 Then result = CleanupOutcome(result, committed, cleanup)
modAppState.Announce result
```

### Cleanup is contained on both paths, and attempted at most once

`FinishOperation` returns a diagnostic `String` for a restoration it could not
complete — but it is an Excel call and can also **raise**. There are exactly two
contexts that may call it, and each runs inside its own cleanup-failure
envelope:

| Context | Envelope | Reached when |
| --- | --- | --- |
| normal path | `NormalCleanupFailed` | the transaction returned, whatever its outcome |
| recovery path | `CleanupFailed` | an error occurred **before** the normal cleanup |

Two Boolean facts carry the distinction, and the guards read those facts rather
than inferring them from where a label sits:

* `stateCaptured` — a snapshot exists and something still owes a restore
* `cleanupAttempted` — the one permitted attempt has been spent, raise or not

`cleanupAttempted` is set **before** each call, so a call that raises still
counts as the attempt. The recovery guard is
`If stateCaptured And Not cleanupAttempted Then`, so it can never retry an
attempt the normal path already spent. Neither handler calls `FinishOperation`
again (`test_48`, `test_57`).

When the normal cleanup **raises**, `NormalCleanupFailed` contains it, captures
`Err.Description`, and publishes through `CleanupOutcome` — the same rule the
returned-diagnostic path uses, so the committed/uncommitted distinction applies
identically:

* **committed = True** — C17 keeps saying `SUCCESS`, no rollback, nothing
  rewritten in `calc_state`; the announcement states that the calculation
  committed and application restoration then failed
* **committed = False** — the existing failed or refused calculation outcome is
  preserved and the cleanup exception is merged into its detail

Nothing on that path touches an analytical block, `WriteAttemptBlock`,
`WriteStatusBlock` or `RestoreSnapshot` — `test_57` refuses each by name.

Every terminal path publishes exactly one outcome (`test_55`); there are now
four.

---

## Corrected: the vacuous manifest proof

Three suites carried a guard of the shape

```python
if not path.is_file():
    return
```

around an assertion about the Stage-B manifest. That passes loudest exactly when
the build is broken. All three now **produce** the manifest by invoking the real
emitter into a fresh temp tree and require the artifact to exist. `test_nc_18`
scans the four relevant suites and fails on any surviving instance of the
pattern — asserted against the real files, not a synthetic string.

The Step-3 absence assertion in `test_phase5_stage_a.py` had also run out of
names ahead of it; it now asserts the boundary that remains — the six endpoints
are declared and none is bound to a button.

---

## Tests

| Module | Tests |
| --- | --- |
| `test_phase1_manifest_validation.py` | 10 |
| `test_phase1_structure.py` | 21 |
| `test_phase2_contract_validation.py` | 42 |
| `test_phase2_inputs.py` | 40 |
| `test_phase3_driver_contract_validation.py` | 31 |
| `test_phase3_drivers.py` | 28 |
| `test_phase3_verifier_intersection.py` | 12 |
| `test_phase4_oracle.py` | 68 |
| `test_phase4_stage_b_source.py` | 155 |
| `test_phase4_structure.py` | 43 |
| `test_phase4_structure_contract_validation.py` | 55 |
| `test_phase5_calc_contract_validation.py` | 151 |
| `test_phase5_fingerprint.py` | 52 |
| `test_phase5_numeric.py` | 94 |
| `test_phase5_oracle.py` | 111 |
| `test_phase5_stage_a.py` | 57 |
| `test_phase5_vba_source.py` | 122 |
| `test_phase5_resolve_source.py` | 91 |
| `test_phase5_check_source.py` | 60 |
| `test_phase5_report_source.py` | **107** |
| **Total** | **1350** |

The Step-6 baseline was 1241; the Step-7 submission was 1313; correction round 1
was 1340. **No test was removed and none was weakened.** Round 1 added 25 tests
to the reporter suite and 2 to the kernel suite; round 2 adds 10 more to the
reporter suite and retargets five in place (`test_20`, `test_21`, `test_48`,
`test_55`, `test_57`). The Step-5 (91) and Step-6 (60) suites are unchanged in
count, and the Stage-A verifier still reports 351/351.

### Mutation evidence

Twenty-four regressions were planted into a scratch copy of `modCalcReport.bas`
and the suite was run against each. **All twenty-four were caught**:

analytical write before preparation · split success metadata writes · commit
outside the rollback envelope · rollback omitting a table · rollback omitting
C23:C32 · rollback omitting prior C13:C20 · FAILED metadata before rollback ·
refusal overwriting C13:C16 · refusal clearing analytical tables · status derived
from the attempt result · REFUSED as a status · FAILED as a status · empty
digests compared as CURRENT · risk `Quantity = 1` in the audit · cost
`Probability = 1` in the audit · Base-Year rate written as zero · reporting
currency seeded into the FX audit · a fabricated zero semantic row · a seventh
endpoint · a report value recomputed from the sheet · write verification removed
· commit verification removed · error suppression · headers rewritten during
calculation.

Twenty further negative controls plant the same defect classes as synthetic
module text.

**Correction round.** Twenty-four *new* regressions were planted into a scratch
copy of the corrected sources — one per defect and per near neighbour — and the
suite was run against each. **All twenty-four were caught**, the last of them
only after the detector for it was strengthened (see below):

| # | Planted regression | Caught by |
| --- | --- | --- |
| M01 | header canonicalised as a number, framed as text | `test_14` |
| M02 | header field assembled by hand from `FP_TAG_NUMBER` | `test_53` |
| M03 | `CalcFpNumberField` made `Private` again | `test_51` |
| M04 | the framer rerouted through the text encoder | `test_51` |
| M05 | `ReportResult` restored in place of `Announce` | `test_54` |
| M06 | a handler path stops announcing | `test_55` |
| M07 | envelope armed after `CaptureAppState`/`BeginOperation` | `test_48` |
| M08 | the recovery cleanup removed | `test_48` |
| M09 | `stateCaptured` never cleared, so cleanup can run twice | `test_57` |
| M10 | preparation and snapshot left uncovered | `test_56` |
| M11 | a runtime fault recorded as a model refusal | `test_56` |
| M12 | FAILED metadata written after a *failed* rollback | `test_28` |
| M13 | a failed record re-running the rollback | `test_28` |
| M14 | `On Error Resume Next` reintroduced | `test_20` |
| M15 | an unreviewed handler added | `test_20` |
| M16 | post-commit cleanup rewriting C17 to `FAILED` | `test_29` |
| M17 | the commit fact never leaving the transaction | `test_29` |
| M18 | the envelope left armed past the commit | `test_29` |
| M19 | C20 dropped from the commit verification | `test_59` |
| M20 | the verifier regenerating `Now` | `test_59` |
| M21 | a second captured moment in the commit block | `test_25` |
| M22 | the commit marked before it is proven | `test_16` |
| M23 | the writer assembling part of its own block | `test_25` |
| M24 | a different block reaching the verification | `test_19` |

**M11 was MISSED on the first sweep.** The suite distinguished refusal from
failure by inspecting `RecordRefusal` and `RecordFailureWithoutRollback` in
isolation, which says nothing about *which handler reaches which*. Swapping the
`PreWriteFailed` handler to call `RecordRefusal` left both procedures intact and
passed. `test_56` now partitions `RunCalculation` at its handler labels and
asserts the routing at the point the two outcomes diverge; `test_nc_27b` plants
the same swap and watches the detector see it.

**Correction round 2.** Twenty-four further regressions were planted into a
scratch copy and the suite run against each. **All twenty-four were caught** —
eleven on the cleanup-containment boundary, eight on the failpoint boundary, and
five re-testing round-1 invariants that must not regress while this round moves
code around them:

| # | Planted regression | Caught by |
| --- | --- | --- |
| N01 | the submitted shape: normal cleanup with the handler disarmed | `test_48`, `test_57` |
| N02 | normal cleanup covered by the top-level envelope instead of its own | `test_48`, `test_57` |
| N03 | the normal-cleanup handler retries `FinishOperation` | `test_48`, `test_57` |
| N04 | the normal-cleanup handler announces nothing | `test_55`, `test_57` |
| N05 | the normal-cleanup handler rewrites `calc_state` | `test_57` |
| N06 | the normal-cleanup handler ignores the committed axis | `test_57` |
| N07 | the attempt marked spent *after* the call, so a raise leaves it unmarked | `test_57` |
| N08 | the exactly-once state removed entirely | `test_48`, `test_57` |
| N09 | the recovery guard drops the spent flag | `test_48`, `test_57` |
| N10 | the recovery handler retries `FinishOperation` | `test_48`, `test_57` |
| N11 | the normal-cleanup envelope left open past the call | `test_57` |
| N12 | the submitted shape: commit hook upstream of `BuildSuccessBlock` | `test_21` |
| N13 | commit hook *after* the C13:C20 assignment | `test_21` |
| N14 | a statement between the hook and the write | `test_21`, `test_25` |
| N15 | the commit hook removed | `test_21` |
| N16 | the hook wired from two places | `test_21` |
| N17 | a hand-rolled second injection mechanism | `test_21` |
| N18 | the analytical hook moved before `WriteAnalytical` | `test_21` |
| N19 | the analytical hook moved past the analytical verification | `test_21` |
| N20 | `Announce` swapped back to `ReportResult` | `test_54`, `test_55` |
| N21 | the top-level envelope armed late again | `test_48`, `test_56` |
| N22 | C20 dropped from the commit verification | `test_59` |
| N23 | the commit marked before it is proven | `test_16`, `test_59` |
| N24 | `On Error Resume Next` reintroduced | `test_20`, `test_48`, `test_57` |

The round-2 suite was run against the round-1 commit `73adbb0` with the
production source restored to its submitted state. **Five reporter tests fail
there**, covering both blockers: blocker 1 (`test_48`, `test_57`, and `test_20`
and `test_55` for the handler and terminal path the fix adds) and blocker 2
(`test_21`).

The corrected suite was also run against the submitted commit `a9de9b3` with the
production sources restored to their submitted state. **Sixteen reporter tests,
two kernel tests and one checker test fail there**, covering all five blocking
defects: blocker 1 (`test_14`, `test_53`, `test_43`, `test_64i`), blocker 2
(`test_54`, `test_55`), blocker 3 (`test_20`, `test_48`, `test_56`, `test_57`),
blocker 4 (`test_29`, `test_58`), blocker 5 (`test_25`, `test_59`).

---

## Correction round 1

Independent review rejected the submission at `a9de9b3` with five blocking
defects. Each is recorded here with what was wrong, what the source does now, and
which test would fail if it came back.

### 1 — the fingerprint header framed numbers as text

`BuildFingerprint` called a private helper that canonicalised each header scalar
as a number and then framed the *result* as an **S** field. Base Year, Start
Year, Duration and Discount Rate are **N** fields in the locked schema, so the
digest covered four text fields carrying numeric text where the contract says
number. Canonicalising and framing are one decision.

The authorised repair was narrow: `CalcFpNumberField` in `modCalcFingerprint`
changed from `Private` to `Public`, and **nothing else in that module changed**.
Its body, its framing, `CalcFpCanonicalNumber`, `CalcFpField`, the UTF-16
handling, `FP_VERSION`, the reducer, the digest, the record schema, the sorting,
the reference digest and the reference stream are all untouched. The reporter now
calls it and neither `CalcFpCanonicalNumber` nor `CalcFpCanonicalText`, and it
does not reproduce N framing of its own.

`modCalcFingerprint` is consequently frozen **twice** — at its current bytes, and
at its Step-4 *executable text* with comments and blanks removed, whitespace
collapsed and that one keyword normalised back to `Private`
(`test_64j`, `test_44` in the Step-6 suite, `test_51` here). The second digest,
`f6e8313b…`, is what "visibility only" means as a check rather than a promise.
The Public fingerprint surface goes from 10 names to 11, and
`CalcFpNumberField` is **not** added to the no-cross-module-caller exception set:
`test_64i` asserts `modCalcReport` really is its caller. `test_64h`'s caller
corpus was widened from the three kernel modules to every hand-written module,
because scoping it to the kernel would have called a genuine consumer "no caller"
— and would equally have let a kernel name be justified by a caller that does not
exist.

The reference digest `50B6EB0E26857EA7` over 366 UTF-16 code units is unchanged.
The Python oracle already framed the four scalars as N fields
(`tests/test_phase5_fingerprint.py:136`); the VBA had disagreed with it, and now
does not.

### 2 — the endpoint bypassed the harness-aware announcement surface

`PCCM_Calculate` ended through `modAppState.ReportResult`, which shows a modal
dialog unconditionally. Gate B drives this endpoint through automation, and a
modal dialog blocks it and leaves the run with no recorded result.

It now ends through `modAppState.Announce`, which records the outcome for
automation and shows the dialog only when automation is inactive. **No new
automation reporter was created and `modAppState` was not modified**;
`test_54` refuses `ReportResult`, refuses six invented channel names, and refuses
`Announce` anywhere but the endpoint.

### 3 — error and cleanup containment started too late

The envelope was armed *after* `CaptureAppState` and `BeginOperation`, so the two
operations that make restoration necessary ran outside it, along with the
preparation, the snapshot and every bookkeeping write. A fault in any of them
escaped raw, past the restoration it had itself made necessary, leaving
`EnableEvents`, `Calculation` mode and `ScreenUpdating` dirty — worse than a
failed calculation.

Now: a top-level envelope installed before the first fallible operation, an
explicit `stateCaptured` flag, `FinishOperation` attempted exactly once, a
`CleanupFailed` envelope over the recovery cleanup, and separate `PreWriteFailed`
and `TransactionFailed` envelopes that keep a controlled refusal and an
unexpected fault apart. A failure before analytical mutation is `FAILED` with no
rollback; a failure after it rolls back first; a **failed rollback writes no
FAILED metadata at all**; and a failed metadata write does not undo a successful
rollback. There is no `On Error Resume Next`.

`test_20` previously asserted that `TransactionFailed` was the only handler in
the module — the assumption that produced the defect. It was **retargeted, not
deleted**: it now asserts the handler set is exactly the six reviewed envelopes,
that none suppresses, and that every armed label is defined where it is armed.

### 4 — post-commit cleanup could retroactively falsify a committed success

After the commit, C17 = `SUCCESS` is committed workbook truth. A later
`FinishOperation` problem is an application/invocation cleanup failure, not a
failed analytical transaction, and rewriting the attempt to `FAILED` would leave
the workbook and the reported outcome contradicting each other.

`RunCalculation` now carries the commit fact out through
`ByRef committed As Boolean`, and `CleanupOutcome` decides on that basis. When
the run committed, it **writes nothing** — no `WriteAttemptBlock`, no
`WriteStatusBlock`, no rollback, no `.Value2` — and reports the cleanup problem
on the invocation axis with a message that states plainly that the calculation
committed.

**On the §4.1 STOP condition:** the existing `OperationResult` shape *can*
represent "committed calculation success plus cleanup/invocation failure"
without ambiguity, so no interface limitation is reported. The two axes are
already separate objects: `PCCM_AutomationResult()` reports the invocation
outcome (`FAIL|Calculate|<detail>`) while `PCCM_CalculationAttemptResult()`
independently reports `SUCCESS` read from committed `calc_state`, and the detail
text names the situation explicitly. Nothing rewrites C13:C20 after the commit.
This is the one place in the correction round where the instruction offered a
stop, and the reviewer should treat the assessment above as the claim to reject
if it is wrong.

### 5 — the success commit was not verified against the exact block written

C20 was excluded from verification, and the success timestamp was only checked
for being non-blank — a check that passes over a stamp written into the wrong
cell. The verifier compared selected semantic fields, not the commit.

The 8×1 block is now built **once** by `BuildSuccessBlock`, with a single
captured `stamp As Date` written into both `built(1,1)` and `built(8,1)`;
`WriteSuccessCommit` performs one assignment of that block; and
`VerifySuccessCommit` compares C13:C20 against **that same block**, all eight
cells, generating no second `Now`. `committed = True` is set only after the whole
block verifies, and a failed verification raises. The old selected-field verifier
was removed rather than kept under the same procedure name — two procedures of
one name is a VBA compile error, and the exact comparison subsumes it.

---

## Correction round 2

Independent review confirmed the five round-1 areas substantially closed and
rejected `73adbb0` with two remaining blockers.

### 1 — the NORMAL-path `FinishOperation` was outside every envelope

Round 1 armed `CleanupFailed` over the **recovery** cleanup only. The normal path
disarmed the top-level envelope first and then called `FinishOperation` with
nothing armed:

```vba
On Error GoTo 0                                   ' <- envelope closed here

cleanup = modAppState.FinishOperation(state)      ' <- and this could raise
```

`FinishOperation` returns a diagnostic `String` for a restoration it could not
complete, so the returned-diagnostic case was handled. But it is an Excel call
and can also **raise**, and a raise there escaped `PCCM_Calculate` with
`stateCaptured` still `True`, no `CleanupOutcome`, no `Announce`, and no
invocation result recorded for automation. That was exactly one uncovered
application-state cleanup path, and round 1 had required the invocation envelope
to contain cleanup failures as well.

**The fix** is a narrow envelope of its own — `NormalCleanupFailed` — armed
before the call and closed immediately after it, plus an explicit
`cleanupAttempted As Boolean` so the exactly-once rule is carried in state rather
than inferred from statement position. The flag is set *before* each call, so a
call that raises still counts as the attempt; the recovery guard became
`If stateCaptured And Not cleanupAttempted Then`; and neither handler calls
`FinishOperation` again. The raised case publishes through the same
`CleanupOutcome` rule the returned-diagnostic case uses, so a committed
calculation keeps C17 = `SUCCESS` and an uncommitted one keeps its own outcome
with the cleanup exception merged in. `modAppState` was not modified, no new
application-state framework was created, and there is no `On Error Resume Next`.

`test_48` now walks **every** `FinishOperation` in the endpoint and requires a
cleanup-failure handler to be armed and still open above each one; `test_57`
proves both contexts independently — armed, spent-before-call, no retry, no
`calc_state` write, and an announcement on the raised path. `test_nc_38` plants
the shipped `On Error GoTo 0` / `FinishOperation` pair verbatim and watches the
detector see it.

### 2 — the success-commit failpoint was not at the C13:C20 assignment

Gate-B plan §25.7 requires the rollback proof at two injection boundaries: after
analytical blocks have been mutated, and **at the final C13:C20 commit
assignment**. The hook was in `RunCalculation`, several statements upstream:

```vba
modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT
BuildSuccessBlock package, successBlock
WriteSuccessCommit successBlock
```

so it exercised a failure during commit *preparation* — before the block was
built and before anything was published — not a failure at the commit boundary.
The test was `assert around < commit`, which any earlier statement satisfies.

**The fix** moves the existing Phase-4 `FailPointCheck` into `WriteSuccessCommit`
as the statement immediately before the assignment. Nothing stands between them.
The writer is called from inside the `TransactionFailed` envelope, after the
analytical snapshot has been written and verified, so an injected exception now
rolls back from the real commit boundary with C13:C20 unpublished. No second
injection mechanism was added, and `FAILPOINT_ANALYTICAL_WRITE` is unchanged:
still after `WriteAnalytical`, still before `VerifyAnalytical`.

`test_21` no longer compares statement indices across procedures. It reads
`WriteSuccessCommit`'s own executable body, requires the hook at exactly
`assignment - 1`, names whatever stands between them if anything does, refuses
the hook anywhere in `RunCalculation`, and refuses a second wiring. `test_nc_43`
plants the shipped upstream placement and asserts that the superseded index check
passes on it while the body check does not — the weak assertion and the strong
one, on the same defect, in one test.

---

## Retargeted invariants

Eight tests named the world before this module existed. Each was retargeted, not
deleted — including two that had asserted the *absence* of what Step 7 builds:

| Test | Now asserts |
| --- | --- |
| `test_04_step_7_does_not_exist_yet` (Step-6 suite) | renamed `test_04_the_checker_owns_no_orchestration`; the checker declares no endpoint and does not reach into the reporter |
| `test_08_the_deferred_phase_6_surface_does_not_exist_yet` (Step-4 suite) | the list has run out of names ahead of it, so it asserts the Step-4 **kernel** does not reach forward into any later module or endpoint |

Correction round 1 retargeted nine more. None was deleted, and each now asserts
the invariant that survived the defect rather than the assumption that caused it:

| Test | Was | Now asserts |
| --- | --- | --- |
| `test_14` | pinned a call list containing `CalcFpCanonicalNumber` + `CalcFpCanonicalText` | the list contains `CalcFpNumberField` and neither primitive |
| `test_17` | took the *first* `On Error GoTo 0` in the procedure | the first disarm **after** the envelope was armed — the earlier one belongs to the pre-write envelope and made the ordering vacuous |
| `test_19` | named `VerifySuccessCommit(package)` | `VerifySuccessCommit(successBlock)` |
| `test_20` | `TransactionFailed` is the **only** handler | the handler set is exactly the six reviewed envelopes, none suppresses, every armed label is defined where armed |
| `test_25` | the commit is one 8-row assignment inside `WriteSuccessCommit` | built once with one captured moment, written once, the writer assembles nothing |
| `test_28` | `RestoreSnapshot` before `RecordFailure` inline in `RunCalculation` | the ordering inside `RollbackAndRecord`, plus: a failed rollback writes no metadata, and a failed record does not re-run the rollback |
| `test_29` | a `committed` guard precedes the restore in the handler | the envelope is disarmed before the commit is published, and the post-commit path writes nothing |
| `test_34` | one `RecordFailure` re-derives the status | **both** failure recorders re-derive it, and `FAILED` is never used as a status |
| `test_48` | capture, begin, finish appear and run before cleanup | the envelope is armed before capture, and the `stateCaptured` flag governs the recovery cleanup |
| `test_64i` (kernel suite) | `CalcFpNumberField` is Private | it is Public, has a real caller in `modCalcReport`, and is not excused by the exception set |
| `test_44` / `test_51` | `modCalcFingerprint` frozen at its Step-4 bytes | frozen at its current bytes **and** at its Step-4 executable text with the one authorised keyword normalised away |

Correction round 2 retargeted five more, again in place:

| Test | Was | Now asserts |
| --- | --- | --- |
| `test_20` | six reviewed handlers | seven — `NormalCleanupFailed` joins them, and every armed label is still defined where armed |
| `test_21` | `around < commit` by statement index in `RunCalculation` | the hook sits at `assignment - 1` inside `WriteSuccessCommit`'s own body, is absent from `RunCalculation`, and is wired exactly once |
| `test_48` | capture/begin/finish appear, envelope armed before capture | **every** `FinishOperation` runs with a cleanup-failure envelope armed and open above it |
| `test_55` | three terminal paths announce | four |
| `test_57` | one cleanup on each path, guarded by `stateCaptured` | both contexts contained independently; the attempt is spent before the call; neither handler retries; the raised path announces and writes nothing |

---

## What remains for Gate B

Everything about behaviour. Whether the transaction commits, whether rollback
restores, whether an injected failure at either failpoint produces the state this
document describes, whether the six status rows come out as specified, and
whether the fingerprint matches `50B6EB0E26857EA7` on real Excel arithmetic and a
real locale. Gate A has established what the source says, and only that.

**No VBA was executed.** No Windows run was made, no Excel COM session was
opened, and nothing in this document or in any test in the repository claims
otherwise — `test_52` scans this suite for such a claim. Everything below is
unproven and remains unproven until Gate B runs on Windows:

* that `modCalcFingerprint` and `modCalcReport` compile at all, including that
  the module now has exactly one `VerifySuccessCommit`
* that the four header scalars produce the N fields the oracle produces, and that
  the digest over a real model is `50B6EB0E26857EA7`
* that `modAppState.Announce` suppresses the dialog under automation and records
  the outcome where the harness reads it
* that `FinishOperation` actually restores `EnableEvents`, `Calculation` mode and
  `ScreenUpdating`, and that the `stateCaptured` path is reached when
  `CaptureAppState` or `BeginOperation` raises
* that a `FinishOperation` which **raises** on the normal path is contained by
  `NormalCleanupFailed`, reaches `Announce`, and leaves a committed C17 reading
  `SUCCESS` — the source says so; no run has shown it
* that an injected failure at `FAILPOINT_SUCCESS_COMMIT` now fires with C13:C20
  unpublished and rolls back from that boundary
* that a fault at either failpoint rolls back to the previous snapshot, that a
  *failed* rollback leaves C17:C20 untouched, and that a failed metadata write
  leaves the rollback standing
* that the eight-cell commit verification passes on real Excel — in particular
  that a `Date` written through `.Value2` reads back equal to the value written
* that a post-commit cleanup failure leaves C17 reading `SUCCESS` while the
  invocation axis reports the failure

Gate B has not been started. No harness work, no Windows execution, no RNG or
Monte Carlo, and nothing from Phase 6 is present in this round.
