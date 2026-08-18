# Phase 5 — Gate A — Step 7: transactional reporting and orchestration

**Status: ready for independent review.**

Step 6 is accepted and closed at `12ade1d`. This step adds the final Phase-5
production module — `modCalcReport` — with the transactional write-back, the
`calc_state` maintenance and the six calculation endpoints, plus the static Linux
suite that reads it and the narrow plumbing the endpoint declaration needs.

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
| `modCalcReport` | 952 | 67 | 176 | 709 | yes | yes |

It fits, with 191 code lines of headroom. No extra module was needed.

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
CaptureSnapshot
On Error GoTo TransactionFailed
    WriteAnalytical                     five tables + calc_totals
    FailPointCheck FAILPOINT_ANALYTICAL_WRITE
    VerifyAnalytical                    read back and compare
    FailPointCheck FAILPOINT_SUCCESS_COMMIT
    WriteSuccessCommit                  ONE 8x1 assignment to C13:C20
    VerifySuccessCommit
    committed = True
On Error GoTo 0
```

**The commit and its verification are inside the envelope** (`test_17`): the
assignment can fail and so can its verification, so both sit inside the rollback
boundary. Success is not published before the analytical snapshot verifies
(`test_18`). A failed verification raises rather than continuing (`test_19`).

There is no `On Error Resume Next`; `test_20` asserts `TransactionFailed` is the
only handler in the module.

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
does `RecordFailure` write one 4×1 block to C17:C20 with `FAILED`, the specific
detail, a **freshly derived** status and a fresh timestamp.

So the final observable state after a failure is:

* **C13:C16** — the previous success block, restored exactly
* **C17:C20** — `FAILED`, the new detail, the current derived status, a new timestamp
* **C23:C32** and the five tables — the previous successful snapshot, restored exactly

The document does **not** claim final C13:C20 equals its previous value. Only
C13:C16 does.

If the failure-metadata assignment itself fails, the already-successful rollback
is not undone; there is no recovery transaction around failure bookkeeping.

`test_29` asserts a committed operation can never be rolled back: the `committed`
guard precedes the restore in the handler.

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

| Constant | Stage name | Position |
| --- | --- | --- |
| `FAILPOINT_ANALYTICAL_WRITE` | `Phase5AnalyticalWrite` | after `WriteAnalytical`, before `VerifyAnalytical` — analytical state is half-written |
| `FAILPOINT_SUCCESS_COMMIT` | `Phase5SuccessCommit` | immediately before `WriteSuccessCommit` |

`test_21` also refuses `gAutomationFailAfterStage` in the module: no second
injection framework. Phase-4 failpoint machinery is untouched, and the Windows
harness was neither extended nor run.

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

`PCCM_Calculate` captures, begins, runs, finishes and reports through the
accepted Phase-4 discipline (`test_48`), and a failed cleanup is reported rather
than swallowed. No `MsgBox`, no change handler (`test_49`), no RNG or `_SimData`
(`test_50`). `modAppState` was not redesigned.

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
| `test_phase5_vba_source.py` | 120 |
| `test_phase5_resolve_source.py` | 91 |
| `test_phase5_check_source.py` | 60 |
| `test_phase5_report_source.py` | **72 (new)** |
| **Total** | **1313** |

The Step-6 baseline was 1241. No test was removed and none was weakened.

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

---

## Retargeted invariants

Eight tests named the world before this module existed. Each was retargeted, not
deleted — including two that had asserted the *absence* of what Step 7 builds:

| Test | Now asserts |
| --- | --- |
| `test_04_step_7_does_not_exist_yet` (Step-6 suite) | renamed `test_04_the_checker_owns_no_orchestration`; the checker declares no endpoint and does not reach into the reporter |
| `test_08_the_deferred_phase_6_surface_does_not_exist_yet` (Step-4 suite) | the list has run out of names ahead of it, so it asserts the Step-4 **kernel** does not reach forward into any later module or endpoint |

---

## What remains for Gate B

Everything about behaviour. Whether the transaction commits, whether rollback
restores, whether an injected failure at either failpoint produces the state this
document describes, whether the six status rows come out as specified, and
whether the fingerprint matches `50B6EB0E26857EA7` on real Excel arithmetic and a
real locale. Gate A has established what the source says, and only that.
