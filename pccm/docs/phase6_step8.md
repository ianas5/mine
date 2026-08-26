# PCCM Phase 6 — Step 8 authority record

Step 8 adds **one module**:

```
src/vba/modSimEngine.bas        the Monte Carlo iteration loop and accumulation
spec/structure_contract.yaml    modSimEngine enters the registry
```

**Not in this step.** No statistic, no mean, no standard deviation, no
percentile or quantile, no contingency, no result digest, no request
fingerprint, no SIM fingerprint, no `run_id`, no AUTO nonce allocation or
persistence, no `_SimData`, no Results, no simulation or attempt state, no
`PCCM_RunSimulation`, no workbook publication, no sensitivity and no annual
stochastic output. Step 9 has not begun. No Windows or Excel runtime ran.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-5 Stage-A authority | `fa30424bfe5e782756fc44b0912c4f113781606c` |
| Accepted Step-6 / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Accepted Step-7 sampler authority | `f2f654eadba4f5196c795e4167b71f7002e1f727` |
| Accepted dependency-corrected head | `78f42439e5e799d860b24475167b77a1472af43a` |
| Step-8 modSimEngine | this commit |

`builder/pccm_builder/sim_oracle.py` remains the single definition of these
semantics. `modSimEngine.bas` is their VBA implementation and creates no
authority of its own.

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

**SOURCE CONFORMANCE — proved now, on Linux.** Purity, the public surface, the
canonical execution order, factor semantics, the consumption contract, the
preparation boundary, transactional output — and the arithmetic those statements
describe, against the accepted Python oracle and the accepted Step-5 corpus.

**VBA EXECUTION CONFORMANCE — NOT proved, deferred to Gate B.** No VBA runtime
exists in this step. No result here may be read as "VBA produced this total".

**The comparison policy is not strengthened.** Each corpus case is checked under
the policy the corpus assigns it — `EXACT`, `TOLERANCE_BOUNDED` under the bound
the accepted Step-0 policy owns, `STATISTICAL` as a distributional statement,
and `SAME_RUNTIME_ONLY` as a relation between two runs. The Beta case is
**not** strengthened to cross-language retained-value equality.

### 2.1 Two primitives are borrowed, not transcribed

`SafeProduct` and `SafeSignedSum` have accepted Phase-5 VBA bodies whose second
tier uses an exact-arithmetic UDT with dynamic limbs and scoped error handlers —
constructs the test transcriber does not model. Their accepted **Python**
counterparts in `calc_numeric` are bound instead, and their real VBA
**signatures are read out of `modCalcFactors.bas`** so the ByRef/ByVal call
convention stays the module's own rather than something retyped in a test. This
is the same discipline Step 7 used for `IsUsableDouble`, applied from the other
side.

### 2.2 The transcriber was extended, not replaced

Three mechanical additions, no second interpreter:

```
Private Type            recognised as well as Public Type, because the engine
                        keeps its prepared representation private
signature_only=         register a procedure's SIGNATURE from source while its
                        body is bound externally (§2.1)
```

Everything else is the Step-7 engine unchanged, and all Step-6 and Step-7 tests
still pass against it.

---

## 3. `modSimEngine.bas`

6 procedures, **one** public, one private `Type`.

| Procedure | Does |
|---|---|
| `SimEngineRun` *(public)* | preflight, prepare once, run the loop, commit the two arrays |
| `SimEnginePrepare` | identifiers, seed, streams, identity verification, Beta shapes |
| `SimEngineClaim` | one component verified and matched to exactly one unclaimed record |
| `SimEngineAdopt` | copy across only what the driver's own kind owns |
| `SimEngineValidateFactor` | the engine-relevant factor domain, and no more |
| `SimEngineSampleValue` | select the family, call the accepted sampler |

### 3.1 The prepared representation is PRIVATE

Step 7 proved a `Public` VBA UDT is caller-writable and had to grow a structural
validator for its prepared Beta shape. The engine declines to create a second
such boundary: `SimEngineDriver` is `Private`, built inside the one public call,
and never leaves it. There is no public prepared-engine carrier to harden.

### 3.2 It owns no randomness and resolves no model

Exactly two `modSimRng` entry points are called — `SimRngStateFromFixedSeed` and
`SimRngBuildComponentStreams` — and **no generator constant is read at all**.
Every draw during the loop goes through `modSimSample`; the engine does not
implement Bernoulli, an inverse CDF or any Cheng arithmetic. No `Rnd`, no
`Randomize`, no D6-11 token, and **no new scoped exception was required or
taken**.

It never touches a worksheet, never reads a register, never builds `Knom` or
`Kpv`, never derives a timeline, and never recomputes FX, inflation, profiles or
discounting. `Knom` and `Kpv` arrive already collapsed by the accepted Phase-5
factor boundary. The only Phase-5 procedures it borrows are `IsUsableDouble`,
`SafeProduct` and `SafeSignedSum` — `test_11` asserts that set exactly.

---

## 4. The locked rules

| Rule | Why it is stated rather than assumed |
|---|---|
| **Preflight first** | Iteration bounds are checked before allocation, before the seed is expanded, before a stream exists and before a draw. `iterations` is a `Long`, so there is no non-whole case at this boundary. |
| **Two limits, no third** | `SIM_MIN_ITERATIONS` is the business minimum and `SIM_MAX_ITERATIONS` the technical ceiling. 100,000 is a performance target, not a cap; `test_12` refuses any four-digit-or-longer literal in the procedure. |
| **The seed arrives selected** | Step 8 does not decide FIXED versus AUTO, allocates no nonce and persists no counter. It expands what it is given through `SimRngStateFromFixedSeed`. |
| **Zero drivers is legal** | Carried forward from the dependency correction. The caller's array may be unallocated, so **no bound of it is read** on that path, no component is inspected, and every retained total is `0#` through `SafeSignedSum(..., 0)`. |
| **Canonical order is READ, not re-derived** | Cost Lines then Risks, ordinal by Permanent ID — taken off the component sequence `SimRngBuildComponentStreams` produced, so there is exactly one collation implementation in the phase. `CL-1000` precedes `CL-999`. |
| **Every component identity is verified** | Kind, role, contiguous stream index, non-blank id, one record per component and one component per record; and the Risk pair must be the same id, adjacent, occurrence first, on two different streams. |
| **Each kind reads only what it owns** | A Cost Line's `Probability` and a Risk's `Quantity` are never adopted, so neither can change preparation, refusal, consumption or output. A Uniform's `Most Likely` is not adopted, not validated, and `SimSampleUniform` has no parameter that could receive it. |
| **The sample is UNIT COST** | `Quantity` is deterministic, sits outside the distribution and is applied exactly once. |
| **D6-18b** | The occurrence draw comes first, exactly once per Risk per iteration at every Probability including 0 and 1; the severity sampler is then invoked **unconditionally**. Consumption is a property of the distribution, not of the occurrence. |
| **A Risk carries no Quantity** | And `Probability` is spent on the Bernoulli draw and appears in no factor. |
| **Two independent accumulators** | PV is built from the same unit cost and the same Quantity against `Kpv`, never by discounting the nominal term. |
| **`SafeProduct` and `SafeSignedSum`** | Not chained multiplication, not a running total: a canonical sequence can overflow at an intermediate point while the final signed sum remains representable. |
| **Allocate once** | Nothing inside the loop reallocates, prepares, sorts, jumps, seeds or reaches a workbook. |
| **No partial success** | The totals are staged locally and committed only after the last iteration. |

### 4.1 Retained-array indexing is zero-based

The authorisation permits one-based indexing *"if consistent with the existing
Step-6/Phase-6 authority"*. It is not: every array `modSimRng` and
`modSimSample` declare is `0 To n - 1`. A one-based retained array would be the
only exception in the phase, and a caller mixing `LBound` conventions across the
three modules is a hazard worth refusing. **Element `i - 1` holds iteration `i`**,
and the module says so where a later reader would ask.

---

## 5. Tests

### 5.1 New files

| File | Contents |
|---|---|
| `tests/test_phase6_sim_engine_vba.py` | 54 Step-8 conformance tests |
| `tests/test_phase6_sim_engine_vba_validation.py` | 42 tests — a baseline and 41 mutation controls |

**Group A — declaration and surface** (`test_01`–`test_05`): the attribute and
`Option Explicit`; the registry entry with nothing beyond it and D6-11
untouched; **exactly one public procedure** with five private helpers and one
**`Private`** `Type`; the signature taking `DriverFactors` and no object; and a
prepared `Type` with no worksheet anchor of any kind.

**Group B — purity and leakage** (`test_06`–`test_11`): the whole module
worksheet-independent, not merely the loop; no module-level or `Static` state;
exactly two `modSimRng` entry points and no generator constant; no sampler
mathematics; no statistic, digest or publication; no model resolution, with the
borrowed Phase-5 set asserted exactly.

**Group C — preflight, seeding and the empty model** (`test_12`–`test_17`):
preflight before everything and no invented cap; out-of-range counts refused
with the caller's arrays untouched; the seed expanded through `modSimRng` and an
inadmissible one refused; a zero-driver model succeeding with `iterations`
all-zero totals **agreeing with the accepted oracle**; that path reading no
driver-array bound and never inspecting the one-slot carrier; and a zero-driver
model still validating its seed.

**Group D — corpus conformance** (`test_18`–`test_28`): `unit_interval` and
`dyadic_mixed` EXACT against the corpus head, tail and distinct count;
`quantity_applied_once` linear and not quadratic, with the corpus's own
"applied twice would be" figures; `d6_18b_unconditional_severity` proving the
severity sequence is **identical draw-for-draw at p = 0.2 and p = 0.8** while
the occurrence counts differ; `degenerate_severity_zero_consumption` invoked
1000 times at zero consumption; `no_beta` under the accepted tolerance;
`with_beta` compared distributionally and not strengthened; replay identity and
seed divergence; a fully degenerate fixture identical for every seed; row-order
invariance over four permutations; and ordinal ordering with no numeric-suffix
parsing.

**Group E — factor semantics** (`test_29`–`test_32`): a Cost Line's
`Probability`, a Risk's `Quantity` and a Uniform's `Most Likely` each poisoned
with `0.5`, `-7`, `NaN` and `inf` and the retained arrays **bit-identical** each
time; and `Central`/`MeanValue`/`CentralBasis` never inspected.

**Group F — accumulation** (`test_33`–`test_37`, `test_54`): independent
accumulators; the accepted primitives with the right logical factor counts; a
constructed non-associative fixture **validated through the independent Python
primitive before** the engine is asked about it; allocate-once; exactly two
retained arrays committed after the loop; and a two-applied-year discounted
fixture where `Kpv ≠ Knom`, so a shared accumulator is arithmetically visible
and not merely textually detectable.

**Group G — preparation and the Step-7 carry-forward** (`test_38`–`test_42`):
nothing prepared or allocated in the loop; **a Beta shape produced only by
`SimSamplePrepareBetaPert`**, with no `.BetaShape.<field> =` anywhere and the
only shape-mentioning assignment being the `HasBetaShape` flag; two Beta drivers
over 1000 iterations costing exactly two preparations; prepared initial states
copied and never mutated; and every component identity verified.

**Group H — failure and transactional output** (`test_43`–`test_49`,
`test_50`–`test_53`): the whole §32 list, including a refusal injected at
iteration 1, 500 and 1000 each naming the iteration and the driver and leaving
the caller's arrays untouched; a numerical refusal naming the driver and the
measure; the Bernoulli trace proving `["occurrence", "severity"] * 2000`; and
the source shape of both contribution arms.

### 5.2 Mutation controls

`test_00` first asserts the accepted source passes **every** detector. Each
control then damages a copy in a temporary directory, reruns the **whole**
battery, and requires a **named** detector among the refusers.

| # | Mutation | Named detector |
|---|---|---|
| 1 | physical input order used as execution order | `test_27` |
| 2 | Risks evaluated before Cost Lines | `test_18` |
| 3 | numeric Permanent-ID suffix sort | `test_28` |
| 4 | the mapping trusts a mismatched identity | `test_45` |
| 5 | the occurrence/severity roles are no longer checked | `test_46` |
| 6 | a duplicate or aliased stream is accepted | `test_42` |
| 7 | a Beta shape assembled by hand | `test_39` |
| 8 | a Beta shape field mutated after preparation | `test_39` |
| 9 | a Uniform's Most Likely is read | `test_31` |
| 10 | a Cost Line's Probability is read | `test_29` |
| 11 | a Risk's Quantity is read | `test_30` |
| 12 | Quantity omitted | `test_20` |
| 13 | Quantity applied twice | `test_20` |
| 14 | total cost sampled instead of unit cost | `test_20` |
| 15 | the Cost nominal term uses `Kpv` | `test_33` |
| 16 | the Cost PV term derived from the nominal one | `test_33` |
| 17 | a Risk contribution gains a Quantity factor | `test_52` |
| 18 | Probability folded into the contribution | `test_21` |
| 19 | severity skipped when the Risk did not occur | `test_50` |
| 20 | severity skipped at Probability 0 | `test_50` |
| 21 | the occurrence draw taken after the severity sample | `test_50` |
| 22 | the engine implements the occurrence comparison itself | `test_09` |
| 23 | `SafeProduct` replaced by chained multiplication | `test_34` |
| 24 | `SafeSignedSum` replaced by a running addition | `test_33` |
| 25 | nominal and PV share one accumulator | `test_33` |
| 26 | the accumulation order reversed | `test_35` |
| 27 | a `ReDim` inside the iteration loop | `test_38` |
| 28 | Beta preparation inside the loop | `test_38` |
| 29 | stream construction inside the loop | `test_38` |
| 30 | a direct generator draw | `test_08` |
| 31 | a generator constant read directly | `test_08` |
| 31a | the D6-11 algorithm token introduced | `test_08` |
| 32 | a worksheet reference introduced | `test_06` |
| 33 | module-level mutable state introduced | `test_07` |
| 33a | a `Static` local introduced | `test_07` |
| 34 | a partial output committed during the loop | `test_47` |
| 35 | a zero-driver model refused, in two spellings | `test_15` |
| 36 | the zero-driver carrier slot treated as a real component | `test_16` |
| 37 | a driver-array bound read on the zero-driver path | `test_16` |
| 38 | a mean implemented prematurely | `test_10` |
| 38a | a result digest implemented prematurely | `test_10` |

**Mutation 21 is unobservable in the numbers** — the occurrence and severity
draws are on different streams, so their order within one iteration changes
nothing arithmetically. Only a source detector can see it, which is exactly why
`test_50` carries one alongside its behavioural trace.

**Two fixtures had to be built specifically so a control could not be vacuous.**
The accumulation-order control needs a term sequence where binary64 addition is
genuinely non-associative *under the transformation the mutation performs* — the
Cost terms reversed, the Risk terms left in place — so `test_35` validates
`[1.0, 1e16, -1e16, 1.0] → 1.0` against `[-1e16, 1e16, 1.0, 1.0] → 2.0` through
the independent Python primitive first. And every other fixture in the suite
collapses to `Knom = Kpv = 1`, which would have made a shared accumulator
invisible; `test_54`'s two-applied-year discounted fixture is what makes it
visible.

### 5.3 Existing inventory tests

The new registry member invalidated eleven tests whose assumption was the prior
exact inventory. None was deleted, skipped or weakened; each consumes a **named**
Phase-6 inventory rather than a count, so the addition stays a visible edit. **No
numeric module-count authority was introduced** anywhere.

One rule was genuinely narrowed, and deliberately: `test_10` of the Step-7 suite
and `test_43` of `test_phase6_sim_sample.py` asserted that the token `Cheng`
appears in no other module's **raw text**. `modSimEngine` states in prose that it
contains no Cheng arithmetic, and a rule that stopped a module naming what it
refuses to contain would forbid the clearest thing it can say. Both now scan
**comment- and string-stripped code** — the same discipline D6-11 enforcement
uses — and the Step-7 suite additionally drops `BetaPert` from its token list,
because `SimSamplePrepareBetaPert` is the accepted **public constructor** and a
module naming it is calling the sampler, not copying it. Every Cheng term and
every projected `SIM_CHENG_*`/`SIM_PERT_*` constant is still forbidden outside
`modSimSample`.

---

## 6. Verification

### 6.1 Python suite

```
2626 passed, 0 failed          (804.40s)
2626 collected
```

| Count | What |
|---|---|
| 54 | Step-8 conformance tests |
| 42 | Step-8 tests — a baseline and 41 mutation controls |
| 74 + 51 | Step-7 sampler conformance and controls — **still green, unmodified** |
| 51 + 32 | Step-6 RNG conformance and controls — **still green, unmodified** |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2530** to **2626**: +96, the 54 conformance tests and 42
controls. **No test was deleted, skipped or weakened**; the eleven inventory
tests the new module invalidated were pointed at a named Phase-6 inventory, and
one rule was narrowed with its reasoning recorded in §5.3.

### 6.2 Stage A

```
351 passed, 0 failed
Stage A build complete.
```

`build/PCCM_stageA.xlsx` — 35 736 bytes, `.xlsx`, no `vbaProject.bin`, no
`.bas`/`.bin` member and no `modSim` member. `modSimEngine.bas` is source VBA
for Stage B; Step 8 does not embed it. **No Windows or Excel runtime ran.**

### 6.3 Module registry after Step 8

```
 1. modConstants        generated       11. modCalcAnalytical
 2. modWorkbook                         12. modCalcFingerprint
 3. modAppState                         13. modCalcResolve
 4. modTimeline                         14. modCalcCheck
 5. modDrivers                          15. modCalcReport
 6. modProfiling                        16. modSimContract   generated
 7. modInflation                        17. modSimRng
 8. modStructuralCheck                  18. modSimSample
 9. modCalcContract     generated       19. modSimEngine              <- new
10. modCalcFactors
```

No count is hardcoded anywhere; P5-M and P5-D8 remain manifest-driven.

### 6.4 Artefact hashes

| Artefact | SHA-256 | Status |
|---|---|---|
| `src/vba/modSimEngine.bas` | `f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab` | **new** |
| `build/stage_b_manifest.json` | `8a606ceb0488adfcfe7e549f74ee26ec675b1238b82f07a748175e8f3da65fb2` | **changed, registry only** |
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` | **byte-identical** |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` | **byte-identical** |
| `build/vba/modSimContract.bas` | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` | unchanged |
| `build/phase6_cases.json` | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` | unchanged |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | byte-identical |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | byte-identical |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | byte-identical |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | byte-identical |

`git diff 78f4243 -- src/vba/modSimRng.bas src/vba/modSimSample.bas
src/vba/modCalcFactors.bas evidence/ builder/ bootstrap/` is **empty**.

### 6.5 The manifest movement is the registry and nothing else

Rebuilt from the Step-7 contract and compared leaf by leaf:

```
keys added      3     (.vba.modules[18].name / .generated / .responsibility)
keys removed    0
values changed  0
```

and the structured forbidden-rule projection is **bit-identical**:

```json
{"construct": "MRG32k3a",      "allowed_in": ["modSimRng"]},
{"construct": "RunSimulation", "allowed_in": []}
```

Step 8 required no new scoped exception and took none.

---

## 7. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed, not touched.

**Two stale descriptive strings** in
`bootstrap/windows/phase4_functional_test.ps1` and
`docs/phase5_gate_b_harness.md`, both still saying "back to 15". Neither is an
executable gate. All three must be settled before the next real Gate-B/Windows
execution.

**Carried to Step 9 and beyond:** the engine's prepared driver representation is
`Private` and never leaves `SimEngineRun`, so the Step-7 requirement — prove
statically that stored prepared Beta shapes originate from
`SimSamplePrepareBetaPert` and are not field-mutated before sampling — is
discharged *here* by `test_39`, and any later module that stores a prepared
shape inherits the same obligation.

---

## 8. Step-8 acceptance gate — self-check

| Gate condition | Status |
|---|---|
| `modSimEngine` exists | yes, 6 procedures, 1 public |
| entirely worksheet-independent | `test_06` |
| input is resolved `DriverFactors`, never worksheet data | `test_04`, `test_11` |
| preflight before stream construction and any draw | `test_12`, `test_13` |
| effective seed expanded through `modSimRng` | `test_14` |
| zero-driver model succeeds with all-zero totals | `test_15` |
| zero-driver path reads no driver-array bound | `test_16` |
| canonical order is Cost Lines then Risks, ordinal by Permanent ID | `test_27`, `test_28` |
| component identities and roles verified before use | `test_42`, `test_45`, `test_46` |
| Beta shapes come only from `SimSamplePrepareBetaPert` | `test_39` |
| no prepared Beta field is mutated afterwards | `test_39` |
| Cost samples UNIT COST; Quantity applied exactly once | `test_20`, `test_53` |
| Cost Probability and Risk Quantity are irrelevant | `test_29`, `test_30` |
| Bernoulli once per Risk per iteration, first | `test_50` |
| severity invoked unconditionally per D6-18b | `test_21`, `test_22`, `test_50` |
| `p = 0` / `p = 1` do not bypass severity | `test_50` |
| Risk contribution carries no Quantity; Probability not folded in | `test_52` |
| `SafeProduct` for every non-zero contribution | `test_34`, `test_53` |
| `SafeSignedSum` for both totals | `test_33` |
| nominal and PV independent | `test_33`, `test_54` |
| contribution arrays canonical and allocated once | `test_36` |
| no allocation, preparation or worksheet access in the hot loop | `test_38` |
| exactly two iteration arrays retained | `test_37` |
| transactional output, no partial success | `test_47` |
| no statistics, fingerprint, digest or publication | `test_10` |
| D6-11 scope unchanged | `test_02`, §6.5 |
| `modSimRng` and `modSimSample` byte-identical | §6.4 |
| no Step 9 exists | no `modSimStats`, `modSimFingerprint`, `modSimReport`, `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B debt | **OPEN**, §7 |
