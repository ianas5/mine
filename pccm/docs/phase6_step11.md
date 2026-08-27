# PCCM Phase 6 — Step 11 authority record

Step 11 adds **one module**, **one procedure appended to an accepted one**, and
**activates the D6-11 endpoint grant**:

```
src/vba/modSimReport.bas       the orchestrator, the dual-bank publication, the seven Phase-6 procedures
src/vba/modCalcReport.bas      + CalcPrepareSimulationInputs, APPENDED after every accepted line
spec/structure_contract.yaml   modSimReport enters the registry; RunSimulation is scoped to it
```

**Not in this step.** No Windows or Excel runtime ran. No workbook was opened,
no macro executed, no simulation produced a number on this machine. Nothing was
added to `spec/sim_contract.yaml`, `spec/workbook.yaml`, `spec/input_contract.yaml`,
`spec/calc_contract.yaml` or `spec/driver_contract.yaml`; nothing in `builder/`,
`evidence/` or `bootstrap/` moved; the accepted Phase-6 kernels
(`modSimRng`, `modSimSample`, `modSimEngine`, `modSimStats`, `modSimFingerprint`)
and `modCalcFingerprint` are byte-for-byte unchanged. **There is no Step 12.**

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-6 / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Accepted Step-7 sampler authority | `f2f654eadba4f5196c795e4167b71f7002e1f727` |
| Accepted Step-8 engine authority | `39415f3` |
| Accepted Step-9 statistics authority | `e760c50361f03bce4a393de64614b1cac45d7d29` |
| Accepted Step-10A request-fingerprint grammar | `34a7c467a2e22c3f896cdc10487a1b3922b4536b` |
| Accepted Step-10 modSimFingerprint | `09a56b2` |
| Accepted Step-11A publication authority | `5e7e7b6` |
| Step-11 modSimReport | this commit |

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

**SOURCE CONFORMANCE — proved now, on Linux.** The public surface, the scoped
D6-11 activation, the transaction ORDER, the ownership of every number, the
shape of every worksheet write, what each failure class may and may not touch,
and the byte-identity of every accepted line of `modCalcReport.bas`.

**VBA EXECUTION CONFORMANCE — NOT proved.** *There is deliberately no
transcription in this step.* Steps 7–10 compiled VBA arithmetic to Python and
ran it; this layer is COM and worksheet orchestration, which the accepted
Phase-6 transcriber models not at all. Building a fake `Worksheet` to run
`modSimReport` against would have proved that the fake behaves — not that Excel
does. **Nothing in this record may be read as "VBA ran a simulation."** Every
runtime claim is Gate-B work and is listed in §8.

---

## 3. The module

`src/vba/modSimReport.bas`, 1065 lines, `Option Explicit`, **no module-level
variable and no `Static`**. The staged run lives in one `Private Type
SimRunPackage` — private, so there is no caller-writable result boundary.

### 3.1 The public surface is exactly seven procedures

```
PCCM_RunSimulation                        Sub    the endpoint
PCCM_SimulationStatus                     String derived, and the only accessor that writes
PCCM_SimulationRequestFingerprint         String the STORED fingerprint of the published bank
PCCM_SimulationResultDigest               String the STORED digest of the published bank
PCCM_CurrentSimulationRequestFingerprint  String what the workbook WOULD request now
PCCM_SimulationAttemptResult              String
PCCM_SimulationAttemptDetail              String
```

Nothing else is `Public`. There is no `PCCM_SimulationRunId`, no
`PCCM_SimulationEffectiveSeed`, no `PCCM_SimulationIterations` — the run id and
the effective seed are committed facts on `_SimData`, not a second API that
could disagree with them.

### 3.2 It owns no mathematics

`test_07` refuses any of `SafeProduct`, `SafeSignedSum`, `SafeMultiply`,
`SafeDivide`, `SafeAdd`, `SimRngNextUniform`, `SimRngJumpNextStream`,
`SimRngBuildComponentStreams`, `SimSample`, `BuildDriverFactors`,
`BuildInflationFactors`, `BuildDiscountFactors`, `BuildKnom`, `BuildKpv`,
`CalcFp`, `Sqr(`, `Log(`, `Exp(`, `FP_BASE`, `FP_MOD_`; pins the kernel entry
points it may name to exactly

```
SimEngineRun  SimStatsDescribe  SimStatsContingency
SimFpBuildRequestFingerprint  SimFpResultDigest
SimRngAutoSeedFromNonce  CalcPrepareSimulationInputs
```

and — because a statistic can be written without naming a single banned
identifier — refuses `/`, `*` and `^` anywhere in executable code, and refuses
any procedure that returns a `Double`. Row and index arithmetic needs `+` and
`-` only. A mean, a share, a variance or an interpolated rung cannot be
expressed here.

---

## 4. The one accepted Phase-5 bridge

`CalcPrepareSimulationInputs` is **appended** to `src/vba/modCalcReport.bas`
after a unique banner. The accepted prefix is proved byte-identical:

```
accepted prefix SHA-256   5d4568aef01037fd2999915da87a550d02033441b8c26c80f9386d4fcf8b087f
```

This is the Step-10 technique, and it is strictly stronger than re-pinning the
whole file: the frozen digest gates in the Phase-5 suites were re-aimed at the
text BEFORE the banner and keep their ORIGINAL literals, so the test proves both
that the accepted bytes did not move AND that exactly one named block exists
beyond them. `test_08` additionally asserts that the only procedure declared
after the banner is `CalcPrepareSimulationInputs`.

The bridge **reuses** the accepted `PrepareCurrentCalculation`, requires
`CALC_STATUS_CURRENT`, and then **projects only**:

```
Drivers   DriverCount   Fingerprint   Totals.ANom   Totals.APv
AppliedTimelineText(package)   HostDecimalSeparator()
```

It recomputes nothing: `BuildFactorTables`, `BuildDriverFactors`, `BuildAudits`,
`BuildAnnual`, `BuildFingerprint`, `AccumulateTotals`, `ResolveModel` and
`Reconcile` are all asserted absent from its body, as is any `.Value2 =` write
and any recalculation.

---

## 5. The transaction

```
PrepareRun            bridge -> iterations -> seed -> machine prerequisites
AllocateAutoNonce     persisted and read back BEFORE any sampling
  [FAILPOINT Phase6AfterNoncePersisted]
RunKernels            engine -> describe x2 -> ladder identity -> contingencies
                      -> request fingerprint -> result digest      (in memory)
package.Stamp = Now
package.TargetBank = InactiveBank(package.ActiveBank)
  [FAILPOINT Phase6CandidateBank]
PublishCandidate      three bulk writes + the iteration bank, INACTIVE bank only
VerifyCandidateBank
  [FAILPOINT Phase6FinalCommit]
FinalCommit           ONE write to D22:D30, active bank as field 9
```

Everything before publication **refuses** (`SIM_ATTEMPT_REFUSED`, three sites);
publication and commit **fail** (`SIM_ATTEMPT_FAILED`, two sites). They are
different records because they mean different things: a refusal changed nothing,
a failure may have left an unpublished candidate bank behind.

### 5.1 Two banks, one selector, one moment

The candidate is written to whichever bank is **not** published; the first
success targets `A`. The reporter's `InactiveBank` yields exactly
`["SIM_BANK_A", "SIM_BANK_B", "SIM_BANK_A"]` for the blank / A / B cases, which
is the contract's `candidate_target` mapping and admits no third bank.

`SIM_IDENTITY_ROW_ACTIVE_BANK` is named in **one** procedure, `ReadActiveBank`,
which reads it. The bank becomes active as the ninth field of the committed
block or not at all. `FinalCommit` contains exactly two `.Value2 =` writes — the
commit and the restoration of the block captured before it — so no field is ever
published on its own.

If the commit does not verify, the captured block goes back and the detail says
`remains authoritative`. If the restoration *also* does not verify, the detail
says so plainly: `could not be restored` … `requires recovery`. It is not
glossed.

### 5.2 A cleanup problem after the commit does not unpublish anything

`CleanupOutcome` is reached with a `committed` Boolean carried out of
`RunSimulation`. Once the active bank has moved, the workbook says SUCCESS and
that is committed truth: the announcement reports an invocation failure, and
nothing rewrites the attempt record, unpublishes a bank or rolls anything back.

### 5.3 The consumed nonce is never rolled back

An AUTO nonce is allocated, persisted and read back **before** any sampling, and
no path decrements it. A refusal after allocation leaves the counter advanced —
which is correct: the seed derived from it may already have been used.

---

## 6. Reporting boundary

- **Results is never written and never reached.** `modWorkbook.Sh(...)` resolves
  to `SIM_DATA_SHEET` and nothing else; `"Results"`, `shResults` and `Results!`
  are absent from the source *with the string literals still in*.
- **Selected CL does not participate in a run.** `inpSelectedConfidenceLevel`,
  `SelectedConfidence`, `NM_INPUT_SELECTED_CONFIDENCE_LEVEL` and `SelectedPx`
  are absent from the source, again scanned with the strings intact — a name
  reached by literal rather than by constant is exactly the evasion this
  forbids. The full ladder is precomputed and published, so moving the selector
  never requires a rerun or a worksheet subtraction.
- **The stored accessors never recompute.** `PCCM_SimulationRequestFingerprint`
  and `PCCM_SimulationResultDigest` read the published snapshot and name no
  `SimFp*`. `PCCM_CurrentSimulationRequestFingerprint` computes and reads no
  snapshot. Only `PCCM_SimulationStatus` writes, and only the derived status
  block.
- **The status derivation is attempt-orthogonal.** It names no
  `SIM_ATTEMPT_*`. A blank active bank is the ABSENCE of a publication, not a
  fourth state: `If Len(active) = 0 Then Exit Function`.
- **The request identity is the request.** Both fingerprint call sites pass
  `package.HasSuppliedSeed, package.SuppliedSeed` and never `EffectiveSeed` — an
  AUTO run must not be hashed as though the caller had supplied the nonce it
  happened to draw, or the same request would get a new fingerprint every run.

---

## 7. D6-11 is activated atomically

`RunSimulation` becomes a scoped construct in the **same commit** that
introduces its owner. Neither the grant nor the module has ever existed without
the other. The structured rules now read exactly:

```
MRG32k3a      -> ["modSimRng"]
RunSimulation -> ["modSimReport"]
Percentile    -> []            (still forbidden everywhere)
```

No button binds to the endpoint in this step.

---

## 8. Deferred to Gate B — NOT executed on Linux

Nothing below ran. Do not read any of it as executed.

**Carried forward, still open:**

- **GATE-B TEMP-DIR CLEANUP DEBT — OPEN.**
- Two stale `"15"` strings.
- `SimStatsLadderExtent` raising arm.
- `SimFpRetainedExtent` raising arm.
- `CalcFpContinueDigest` real-VBA execution.
- `AscW` vs `Asc` behaviour on a real host.

**Opened by this step:**

- Every worksheet read and write in `modSimReport` — no COM ran.
- The three named failpoints (`Phase6AfterNoncePersisted`,
  `Phase6CandidateBank`, `Phase6FinalCommit`) and the crash/recovery cases they
  exist for.
- The bulk-write and chunked iteration-bank writes at real iteration counts.
- `VerifyCandidateBank` and `SameBlock` against real Excel round-tripping.
- The final commit's atomicity as Excel actually performs it.
- The AUTO nonce persist-and-read-back.
- The application-state envelope, `FinishOperation` and every `Announce`.
- `CalcPrepareSimulationInputs` against a real Phase-5 CURRENT workbook.
- The derived status against a real published bank.
- Results presentation reading a published bank.

**Results stale-text debt.** Two Results strings are stale —
"Simulation output - no statistics are implemented in Phase 1" and
"Not implemented yet. No percentiles, moments or cash flows are implemented
yet." `spec/workbook.yaml` was deliberately NOT modified in this source
implementation round; this is carried to Step-12 cleanup together with the two
stale `"15"` strings.

**Stage-A workbook binary SHA.** It is not a stable semantic authority unless
`PCCM_BUILD_TIMESTAMP` is pinned. The Stage-A SHA reported in the submission is
build-instance evidence only, and does not claim to reproduce any earlier digest
without the same timestamp override.

---

## 9. Tests

| Suite | Count |
|---|---|
| `tests/test_phase6_sim_report_vba.py` — source conformance | 47 |
| `tests/test_phase6_sim_report_vba_validation.py` — mutation controls | 55 + baseline |

Every control damages one of the three authorities this step touches, reruns the
**whole** Step-11 conformance battery against the damaged copy, and requires a
**named** detector among the refusers. Damaged copies live in a temporary
directory; nothing is written to the repository.

**A vacuity defect was found and fixed while building this layer.** The
conformance module bound its source paths as *default arguments*, evaluated at
import time, so re-aiming them at a damaged copy silently kept reading the
accepted file and 43 of the 56 controls "passed" without proving anything. The
paths are now resolved at call time, and the comment in the source says why.

Seven detectors were **strengthened** — never weakened — because a control
survived them:

| Detector | What it now also refuses |
|---|---|
| `test_07` | `/`, `*`, `^`, and any `Double`-returning procedure |
| `test_13` | the banned control names **as string literals** |
| `test_15` | `EffectiveSeed` inside either fingerprint call |
| `test_22` | a qualified ladder assignment (`package.PvLadder(...) =`) |
| `test_26` | the three `InactiveBank` outcomes, in the contract's order |
| `test_31` | any third write inside `FinalCommit` |
| `test_33` | `SIM_IDENTITY_ROW_ACTIVE_BANK` outside `ReadActiveBank` |

No test was deleted, skipped or weakened anywhere in the repository.

---

## 10. What changed

```
NEW   src/vba/modSimReport.bas
NEW   tests/test_phase6_sim_report_vba.py
NEW   tests/test_phase6_sim_report_vba_validation.py
NEW   docs/phase6_step11.md
MOD   src/vba/modCalcReport.bas          the appended bridge ONLY
MOD   spec/structure_contract.yaml       modSimReport registered; RunSimulation scoped
MOD   18 existing test files             inventory / scope / endpoint expectations
```

Every one of the eighteen existing test files moved for the same reason: a
newly authorised module, endpoint, bridge or scoped token changed the EXACT
expectation the test states. Each remains exact — a set, a sorted list or a
named inventory — and none became "contains at least".

The repo-wide D6-11 sweeps in `tests/test_phase4_stage_b_source.py` now assert
both scoped grants per module, which is a stronger statement than the blanket
ban they replace: each construct must be present in its owner and absent from
every other module.

**Digests**

```
src/vba/modSimReport.bas        a0b9a738b8f7346efd7f5964c311861d975075786072e2ec7b7c7773afd0c363
src/vba/modCalcReport.bas       8252b935b256b1abad9b26ca6b1d90c92c5e0d7566906308b191cd03dd6a71b3
  its accepted prefix           5d4568aef01037fd2999915da87a550d02033441b8c26c80f9386d4fcf8b087f
build/stage_b_manifest.json     51335e3339ab480b28760a2c58fbe83e72ce3d1be54554766857598db7272049
```

**Unchanged, as required**

```
build/vba/modSimContract.bas    1d949be659d0afc3e18501a34b7d372bab3df575fc1a981cfd60dcf1f293a753
build/phase6_cases.json         98f835375f5b8f548172c21ae6102b50fef7e6a001e196ece0741c987d78b6d1
src/vba/modSimRng.bas           3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3
src/vba/modSimSample.bas        5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877
src/vba/modSimEngine.bas        f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab
src/vba/modSimStats.bas         98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0
src/vba/modSimFingerprint.bas   9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda
src/vba/modCalcFingerprint.bas  2efbb30c6f915c04b9c07adec07e25e11f4b5bd2b98e3efa818631dc510ce847
```
