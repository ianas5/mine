# Phase 5 — Gate A — Step 6: the numerical prerequisite checker

**Status: ready for independent review.**

Step 5 is accepted and closed at `277b51e`. This step adds one hand-written VBA
module — `modCalcCheck` — that validates the already-resolved model against the
Phase-5 numerical prerequisites, plus the static Linux suite that reads it, plus
the one-line inventory declaration it needs.

---

## What this step does NOT claim

1. **NO VBA WAS EXECUTED.** There is no VBA interpreter on Linux, and none was
   simulated. Nothing here observed the checker accept or refuse a model.
2. **No model was checked.** Every assertion in
   `tests/test_phase5_check_source.py` is a statement about source text: which
   predicates exist, in what order they run, and which authority each number
   comes from.
3. **No refusal has been seen to reach anyone.** Whether a diagnostic is
   readable, whether Excel behaves as the source expects — Gate B's, on real
   Excel on Windows.
4. **The checker has never been compiled.** `Option Explicit` is present and the
   source tests check balance and declaration shape, but a static reader is not
   a compiler.
5. **Nothing was written anywhere.** No `_Calc` write-back, no attempt or status
   metadata, no `MsgBox`, no `PCCM_Calculate`, no button, no harness change, no
   RNG.

---

## Responsibility

`modCalcCheck` is the Phase-5 **numerical** prerequisite checker. It reports, it
refuses, and it never repairs.

### Three owners, and this is the third

| Owner | Concern |
| --- | --- |
| **Phase 4** | the STRUCTURAL prerequisites — the structural state and `ValidateStructure()`: ID patterns, duplicates, orphans, profiling and inflation grid shape, counters. Already invoked by `modCalcResolve` before anything is resolved. |
| **Step 5**, `modCalcResolve` | RESOLUTION — reference sets, referenced-only FX and inflation, D2, exact identifiers, strict numeric typing, blank versus zero, Permanent-ID profiling. |
| **Step 6**, `modCalcCheck` | the NUMERICAL prerequisites over the resolved model. |

None of the first two is re-implemented. `test_41` refuses every Phase-4
structural identifier in the checker; `test_42` refuses every Step-5 resolution
identifier; `test_46` in the Step-5 suite was retargeted to assert the split from
the other side — the resolver does not call the checker and does not evaluate its
predicates.

### It checks the resolved model, not the workbook again

Every value the checker needs is already in `ResolvedModel`. Re-reading a cell
would create a second resolution authority and break the locked pipeline —
*resolve everything into memory, validate everything in memory, calculate
everything in memory*. So there is **no worksheet access at all**: no `Range`, no
`ListObject`, no defined name, no `modWorkbook` (`test_05`). `test_06` goes
further and refuses the resolver's own entry points and table constants by name,
so a second read path cannot appear by reaching sideways instead of downwards.

No new worksheet read was needed for any locked predicate in this step.

### It owns no arithmetic

The profiling sum goes through `modCalcFactors.SafeSignedSum` and the difference
from 100% through `SafeSubtract`. `test_43` asserts those are the **only** two
calls into the numerical kernel. No new primitive was required, and no accepted
one was altered.

---

## Public API

```vba
Public Function CheckResolvedModel(ByRef model As ResolvedModel, _
                                   ByRef detail As String) As Boolean
```

One entry point, whitelisted exactly in both directions (`test_08`). Everything
else is a private helper: `CheckTimeline`, `CheckDiscountRate`, `CheckDriver`,
`CheckOrdering`, `CheckQuantity`, `CheckProbability`, `CheckProfileSum`,
`DriverLabel`, `OrderingFailure`.

No `PCCM_` procedure exists. The Phase-5 automation surface belongs to the later
orchestration step.

---

## The predicates

### Applied Base Year ≤ Start Year

```vba
If timeline.BaseYear > timeline.StartYear Then
```

The price Base Year may **equal** the Start Year — the ordinary one-year case —
and may precede it, in which case pre-project years participate in inflation
compounding. It may not follow it. `test_16` asserts the comparison is `>` and
not `>=`, so an equal pair is not refused.

Phase 4 may also report this relationship structurally. That does not remove the
Phase-5 predicate: two consumers at different boundaries is fine, and a numerical
layer that assumed someone else had checked would be assuming rather than
checking.

The refusal names both values. Nothing is altered and no replacement year is
derived.

### Discount rate — D3

```vba
If timeline.DiscountRate <= -1# Then
```

`1 + r > 0`. Step 5 has already proven the value is a usable Double, so for a
finite Double the condition is exactly `r <= -1` and no arithmetic is needed.

The rate is **not** clamped, **not** defaulted to zero and **not** replaced with
an identity factor. `BuildDiscountFactors` stays in `modCalcFactors`, unweakened
and unreferenced from here (`test_18`); it is simply never reached with a rate
that cannot produce a factor.

### Profiling sum — the rule Step 5 deferred

For every driver, the project-year weights must sum to `1` within
`TOL_PROFILING_SUM_ABSOLUTE`.

```vba
modCalcFactors.SafeSignedSum(weights, count, total)
modCalcFactors.SafeSubtract(total, PROFILE_SUM_TARGET, difference)
If Abs(difference) > TOL_PROFILING_SUM_ABSOLUTE Then
```

**Signed, and through the accepted primitive.** A hand-written accumulation would
be a second summation rule with none of `SafeSignedSum`'s behaviour — in
particular none of the tier-2 exact rescue, so a profile whose partial sums step
outside Double range would be refused rather than producing its representable
answer (erratum C2). `test_20` refuses a `total = total + …` loop.

**The tolerance is the contract's.** `test_21` refuses any tolerance literal in
the module and asserts the only constant it declares is `PROFILE_SUM_TARGET`.

**No individual weight is judged by sign.** The locked rule is about the sum. A
numeric zero is legitimate — a driver may genuinely spend nothing in a year — and
so is a negative weight: a credit, a transfer out. `test_24` refuses any
per-weight sign test. Blank weights never reach here; Step 5 already refused them.

A sum that cannot be represented is a controlled refusal, never a fabricated zero
(`test_23`). The failure names the Permanent ID, the resolved sum, the target and
the tolerance.

### Distribution ordering

Dispatched on the `DistKind` Step 5 already resolved — the distribution **name**
is not mapped a second time (`test_28`), and appears only in the diagnostic.

| Kind | Rule |
| --- | --- |
| Triangular, Beta-PERT | `Min <= Most Likely <= Max` |
| Uniform | `Min <= Max` only |

**D1: for Uniform a populated Most Likely is accepted and IGNORED.** The cell may
hold a leftover from another choice of distribution, and refusing it would block
a valid model. `test_27` reads the statements of the `Case DIST_UNIFORM` branch
and asserts none of them mentions `MostLikely` at all — the value is not
consulted, not merely tolerated. `HasMostLikely` appears nowhere.

**No positivity rule is invented.** A correctly ordered set of negative values is
a valid distribution and no accepted contract says otherwise (`test_29`).

**No statistic is computed.** An ordering check that called a mean would be doing
the calculation early, and could refuse for a representability reason that has
nothing to do with ordering (`test_30`).

An unrecognised kind refuses rather than passing silently — unreachable through
the resolver, but a silently unchecked driver is worse than an unexpected
refusal.

### Cost Quantity, Risk Probability

| Kind | Rule |
| --- | --- |
| Cost line | `Quantity > 0` |
| Risk | `0 <= Probability <= 1`, both boundaries valid |

A cost line's check does not mention Probability, and a risk's does not mention
Quantity (`test_33`, `test_35`). The `Quantity = 1 for risks, Probability = 1 for
cost lines` convention belongs to the in-memory `DriverFactors` carry type the
calculation and the simulation share. It is not a user input, so there is nothing
for a user to have got wrong, and no fabricated field appears in the resolved
model.

---

## The empty driver set

A model with no cost lines and no risks is valid. The checker evaluates the
model-level predicates — Base/Start and D3 — and then succeeds, producing zero
ordering checks, zero scalar checks and zero profiling checks (`test_37`).

The count is tested **before** any bound of `Drivers` or `Weights` is read
(`test_38`), because a VBA array cannot represent a zero-element set and an
unallocated dynamic array raises on `LBound`. `test_39` asserts the model-level
predicates still run: an empty driver set does not excuse a bad timeline or an
impossible discount rate. A negative count is refused.

---

## Reports, never repairs

This is the defining property, and it is asserted structurally rather than
described. `test_09` collects every statement that assigns **into** `model`,
`driver`, `timeline` or `weights` and requires the list to be empty — the one
permitted exception being the read of a weight *out of* the model into a local
array. `test_10` additionally refuses clamping, normalising and defaulting
vocabulary.

There is no `On Error` at all (`test_12`): every refusal is a returned `False`
with a diagnostic. `test_13` requires every procedure that can refuse to set one,
and `test_14` requires a failed sub-check to return immediately, so a specific
diagnostic naming a driver and a rule is never overwritten by a generic one.

Nothing is published. No `MsgBox`, no `modAppState`, no `_Calc`, no attempt
metadata (`test_11`). Later orchestration owns telling the user; the checker
returns status and detail.

---

## Line metrics

| Module | Raw | Blank | Comment | Code | code < 900 | raw < 1200 |
| --- | --- | --- | --- | --- | --- | --- |
| `modCalcCheck` | 284 | 18 | 135 | 131 | yes | yes |

---

## Build plumbing

`spec/structure_contract.yaml` now declares fourteen modules — the accepted
thirteen plus `modCalcCheck`, `generated: false`. The generated set remains
exactly `{modConstants, modCalcContract}` (`test_02`). Nothing else in the
builder changed: `modConstants.bas` and `modCalcContract.bas` render identically,
and the Stage-B manifest gains only the one inventory entry.

### Retargeted invariants

Six existing tests named the world before this module existed. Each was
retargeted, not deleted.

| Test | Was | Now |
| --- | --- | --- |
| `test_05_no_module_is_a_dumping_ground` | four Phase-5 modules under the code/raw limits | five; the checker is measured by the same policy |
| `test_02_step_4_added_exactly_three_modules_and_no_fourth` | inventory = Phase 4 + kernel + the resolver | + exactly one Step-6 module; `KERNEL_MODULES` is still exactly three |
| `test_08_the_deferred_phase_6_surface_does_not_exist_yet` | `modCalcCheck` on the deferred list | it left at Step 6; the reporter and the five accessors remain |
| `test_08_no_phase_5_endpoint_or_later_module_was_added` (Step-5 suite) | same | same |
| `test_46_the_numerical_checker_still_does_not_exist` (Step-5 suite) | the checker must be absent | renamed `test_46_the_resolver_does_not_perform_the_numerical_checks`; it now asserts the SPLIT from the resolver's side — the resolver neither calls the checker nor evaluates its predicates |
| `test_the_stage_b_manifest_carries_the_implemented_modules_and_nothing_later` | the checker must be absent from the manifest | it must be present; the reporter and endpoints must be absent |

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
| `test_phase5_check_source.py` | **60 (new)** |
| **Total** | **1241** |

The Step-5 baseline was 1181. No test was removed and none was weakened.

### Mutation evidence

Eighteen regressions were planted into a scratch copy of `modCalcCheck.bas` and
the suite was run against each. **All eighteen were caught**, none silently:

| Planted regression | Result |
| --- | --- |
| Base/Start check removed | caught |
| D3 removed | caught |
| profiling sum replaced with a naive accumulation | caught |
| tolerance hard-coded instead of the contract constant | caught |
| a negative individual weight rejected | caught |
| Uniform made to require `Min <= ML <= Max` | caught |
| Quantity accepts zero | caught |
| Probability accepts `< 0` | caught |
| Probability accepts `> 1` | caught |
| empty driver set refused | caught |
| repairs `Quantity = 1` | caught |
| normalises the weights | caught |
| re-reads the workbook | caught |
| an early `PCCM_Calculate` endpoint | caught |
| a second distribution-name mapping | caught |
| a generic diagnostic overwrites a specific one | caught |
| `On Error Resume Next` | caught |
| an array bound read before the empty branch | caught |

Fourteen further negative controls plant the same defect classes as synthetic
module text and assert the sweeps see them.

### The suite guards its own language

`test_45` reads the file back and fails if it ever contains a phrase claiming the
VBA ran or a real model was checked. The forbidden phrases are assembled from
parts so the guard cannot trip over its own wording.

---

## What remains for Step 7

The checker returns a verdict and a diagnostic. It does not publish either.
Still ahead:

* `modCalcReport` and the transactional write-back to `_Calc`
* `PCCM_Calculate` and the five Phase-5 status accessors
* `calc_state` maintenance and the attempt state machine
* status and fingerprint orchestration, and the Calculate button
* the Gate-B Windows harness

## What cannot be proven until Gate B

Everything about behaviour. That these predicates accept the models they should
and refuse the ones they should; that `SafeSignedSum` returns what the reference
returns on real Excel arithmetic; that a diagnostic reaches a user in a readable
form. Gate A has established what the source says, and only that.
