# PCCM Phase 6 — Step 7 authority record

Step 7 adds **one module**:

```
src/vba/modSimSample.bas        the pure stochastic transforms
spec/structure_contract.yaml    modSimSample enters the registry
```

**Not in this step.** No simulation iteration orchestration, no D6-18 severity
invocation, no Cost Line or Risk contribution arithmetic, no `Quantity`, `Knom`
or `Kpv`, no retained iteration arrays, no statistic, no quantile, no
contingency, no request fingerprint, no result digest, no `_SimData`, no
Results, no run state, no run counters and no `PCCM_RunSimulation`. Step 8 has
not begun. No Windows or Excel runtime ran.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Step-4 oracle authority | `614be1acf0f69c16443ace5381edf6157e0f57d3` |
| Accepted Step-5 Stage-A authority | `fa30424bfe5e782756fc44b0912c4f113781606c` |
| Accepted Step-6 VBA RNG / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Step-7 modSimSample | this commit |

`builder/pccm_builder/sim_sample.py` remains the single definition of these
semantics. `modSimSample.bas` is their VBA implementation and creates no
authority: every operational value it uses is read from `modSimContract`, which
projects `spec/sim_contract.yaml`.

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

**SOURCE CONFORMANCE — proved now, on Linux.** The module is read as text: its
purity, its public surface, its use of the projected constants, the shape of
every locked formula and branch operator, the state-consumption contract, the
preparation boundary, and the arithmetic those formulas describe — against the
accepted Step-0 Cheng vectors and the accepted Step-5 corpus.

**VBA EXECUTION CONFORMANCE — NOT proved, deferred to Gate B.** No VBA runtime
exists in this step and none was authorised. `tests/test_phase6_sim_sample_vba.py`
carries the distinction in its own banner:

> No test in this file may be read as "VBA reproduced Cheng vector 24".

**The comparison policy is not strengthened.** `build/phase6_cases.json`
classifies transformed sampler outputs as `TOLERANCE_BOUNDED`, and they are
compared under a bound this suite **reads out of** the accepted Step-0 evidence
policy (§10.3 of `docs/phase6_step0.md`) rather than restating — `rel ≤ 1e-12`
with an `abs ≤ 1e-12·s` floor for transformed samples, `rel ≤ 1e-11` for the
deterministic Cheng vectors. They are **not** compared with `==`, however close
the transcription lands. What stays EXACT is what the corpus says stays exact:
draw counts, consumption, proposal attempts, RNG states, dispatch, orientation
and Bernoulli decisions.

### 2.1 The transcriber was reused, not rebuilt

Step 6's test suite already contained a source-transcription engine. It was
extracted **mechanically** into `tests/phase6_vba_transcribe.py` so Step 7 could
reuse it rather than grow a second transcription language. The Step-6 semantics
are unchanged and all 47 Step-6 conformance tests and 27 Step-6 mutation
controls still pass unmodified. The additions are:

```
Do ... Loop                      an unconditional loop, as the samplers are written
Log / Exp / Sqr / Abs            mapped to DOTLESS names, because the member-access
                                 rewrite would read `math.log` as a UDT field
Boolean UDT field default        False, not 0.0
multi-module compilation         modSimRng + modSimSample in one namespace
a per-module procedure filter    so ONE accepted Phase-5 predicate can be borrowed
```

It lives under `tests/`, carries **no authority**, and nothing in `builder/`,
`src/`, `spec/` or `bootstrap/` imports it.

---

## 3. `modSimSample.bas`

13 procedures, 5 public, 1 public `Type`.

| Public | Does |
|---|---|
| `SimSampleUniform` | one uniform, the stable convex transform |
| `SimSampleTriangular` | one uniform, the conditioned inverse CDF |
| `SimSamplePrepareBetaPert` | the per-DRIVER shape. Draws nothing |
| `SimSamplePreparedBeta` | the per-ITERATION path. Cheng BB or BC |
| `SimSampleBernoulli` | one uniform, one occurrence decision |

| Private | Does |
|---|---|
| `SimSampleChengBB`, `SimSampleChengBC` | the two rejection samplers |
| `SimSampleOrientedBeta` | the shared return rule, so the two cannot drift |
| `SimSampleOrderedTriple` | finite and `Min ≤ Most Likely ≤ Max` |
| `SimSampleScale` | `s = max(|a|, |m|, |b|)`, never zero |
| `SimSampleShapeInFamily` | the accepted `[1, 5]` shape bound |
| `SimSampleMinOf`, `SimSampleMaxOf` | the orientation primitives |

### 3.1 It owns no randomness

Every uniform arrives through exactly two accepted `modSimRng` entry points —
`SimRngValidateState` and `SimRngNextUniform`. There is no recurrence, no jump,
no seeding, no `Rnd`, no `Randomize`, and **not one generator constant is read**
(`test_08` sweeps every `SIM_RNG_*`, `SIM_JUMP_*`, `SIM_SEED_*`, `SIM_NONCE_*`,
`SIM_AUTO_*`, `SIM_STREAM_*` and `SIM_COMPONENT_*` name out of the source).

**Step 7 required no new D6-11 exception and took none.** `MRG32k3a` remains
scoped to `modSimRng` alone; `modSimSample` does not contain the token.
`RunSimulation` remains global.

### 3.2 Purity

No `Worksheet`, `Range`, `Cells`, `ListObject`, `Workbook`, `Application`,
`Evaluate`, `MsgBox`, file I/O, `Environ`, `Shell`, `Date`, `Now`, `Timer` or
`DoEvents`. No module-level variable, no `Static` local. No public parameter is
`Object`, `Variant` or a COM type. No caller row number anywhere.

### 3.3 Failure and commit semantics

Every state-consuming sampler validates, copies the caller's state to a **local
working copy**, draws against that copy, validates the candidate, and commits
last. On failure the caller's state, sample and counts are exactly as they were —
`test_35`–`test_44` prove each of those, including a Cheng failure injected
*after* three proposals have already been consumed.

---

## 4. The locked shapes

| Rule | Why it is stated rather than assumed |
|---|---|
| **Uniform has no Most Likely parameter** | Accepted Phase-5 D1 ignores it numerically. An argument that cannot be passed cannot be read. |
| **Uniform degeneracy is `a = b`, only** | A shared `a = m = b` predicate would make a degenerate Uniform depend on the input the family ignores: `Min = Max` with a populated unrelated Most Likely would enter the sampler, consume a uniform, and shift every later draw on that component. |
| **`x = (1-u)·a + u·b`** | Not `a + u·(b-a)`: for `a = -MAX, b = +MAX` the difference overflows while every convex result is representable. `test_17` asserts `b - a` is *not* finite for the corpus's extreme span, so the case still exercises what it was written for. |
| **Triangular is conditioned** | `(b-a)(m-a)` overflows for supports near Double maximum long before the answer does. |
| **`u <= c`** | At `u = c` exactly the result is the mode. `m = a` gives `c = 0` and `m = b` gives `c = 1`, by the arithmetic, with no endpoint special case. |
| **Degeneracy precedes parameterisation** | `a = m = b` would make `(m-a)/(b-a)` an evaluated `0/0`. The order of the two steps is the reason it never is. |
| **`min(alpha, beta) > 1` is BB; equality belongs to BC** | `m = a` gives `alpha = 1` and `m = b` gives `beta = 1`, so both endpoints reach BC **by the rule**. |
| **BB orients min, max; BC orients max, min** | Opposite. Inverting one is silent: the sampler still returns a valid Beta variate, of the mirrored distribution. |
| **Two uniforms per proposal, `u1` then `u2`** | And a rejection consumes them exactly as an acceptance does. |
| **No rewind, ever** | A rejected proposal advances the local working state and the retry continues from there. |
| **`Log(u1 / (1 - u1))`** | `Log(u1) - Log(1 - u1)` is the same function of a real number and a different function of a Double. |
| **The BB acceptance order** | Squeeze, then the log test, then the full test. Reordering changes *which* proposals are accepted, and therefore consumption and every draw after it. |
| **`u1 < 0.5`, `z <= 0.25`, `z >= k2`, final `>=`** | Operators are authority. The `z <= 0.25` arm accepts **immediately**, without the final test. |
| **Bernoulli is strict `<`** | Because a raw uniform is strictly inside `(0,1)`, strictness is what makes `p = 0` never occur and `p = 1` always occur — exactly, with no special case. At `u = p` the answer is False. |
| **One uniform for every valid `p`** | `p = 0` and `p = 1` included. Skipping the draw would desynchronise the component stream against every other run of the same model. |
| **Nothing clamps** | Not the probability, not `y`, not the sample. Out of domain is refused. |

### 4.1 One deliberate non-gate: `alpha + beta = 6`

The accepted family satisfies `alpha + beta = 6` in real arithmetic, and the
module checks that both shapes lie in `[SIM_PERT_SHAPE_LOWER,
SIM_PERT_SHAPE_UPPER]`. It does **not** gate on the sum.

`1 + 4r` and `1 + 4(1-r)` can each round, and a sweep of 400,000 values of `r`
found **2,083** where the binary64 sum is one ulp from six (`r =
0.00027500068750171877` gives `5.999999999999999`). Gating on it would refuse
correct shapes, and a tolerance inside a sampler is exactly what the numerical
authority forbids. The identity is **evidence**, asserted against the accepted
corpus by `test_24` and `test_25`, and the module says so in the comment where a
later editor would be tempted to add the gate.

### 4.2 Hot-loop precomputation

`SimSamplePrepareBetaPert` settles the conditioning, `alpha`/`beta`, the BB-vs-BC
dispatch, the Cheng orientation and every derived term — the square root, gamma,
delta, `k1`, `k2` — **once per driver**. `test_52` proves the proposal loops
contain no `Sqr(`, no `SimSampleMinOf`/`MaxOf`, no `SIM_PERT_*`, and no
assignment to any `Cheng*` term; that every `prepared.` reference in the loop is
a *read*; and that no procedure other than the preparation writes the shape at
all.

**There is no one-shot public Beta sampler.** `test_14` proves no public sampler
calls another, so no convenience entry point can quietly become the hot path.

---

## 5. Tests

### 5.1 New files

| File | Contents |
|---|---|
| `tests/phase6_vba_transcribe.py` | the transcriber, extracted mechanically from the Step-6 suite — test-only, no authority |
| `tests/test_phase6_sim_sample_vba.py` | 61 Step-7 conformance tests |
| `tests/test_phase6_sim_sample_vba_validation.py` | 39 mutation controls (the 34 required, a baseline, and four extras) |

**Group A — the module exists and is declared** (`test_01`–`test_05`): the
`VB_Name` attribute and `Option Explicit` with no `Option Base`; the registry
entry as `generated: false`; the exact five-name public surface with all eight
numerical helpers private; no public parameter that is an object, a `Variant` or
a COM type and every parameter explicitly typed; and one public `Type` whose
seventeen fields are exactly the prepared shape, with `UseChengBB` a **Boolean**
rather than a magic string and no RNG state, worksheet object or driver row.

**Group B — purity and provenance** (`test_06`–`test_10`): the
workbook/environment sweep; no module-level or `Static` state; all randomness
through exactly `SimRngValidateState` and `SimRngNextUniform` with no generator
constant read; the D6-11 grant unchanged and the token absent; and the Cheng
formulation unique to this module.

**Group C — no engine leakage** (`test_11`–`test_14`): the whole §26 token list
absent; no `PCCM_` macro and no overlap with the declared VBA surface; Bernoulli
a primitive that nothing else in the module calls; and no public sampler calling
another.

**Group D — source arithmetic conformance** (`test_15`–`test_34`): the tolerance
policy read from its owner rather than restated; the accepted convex Uniform
transform on every injected row; the extreme span staying finite where the naive
difference does not; negative and crossing-zero supports staying inside their own
support and monotone in `u`; both degenerate Uniform cases at zero draws; both
Triangular branches and the branch point; a mode at either endpoint; a
crossing-zero extreme support; the degenerate Triangular; the five-rung PERT
ladder with its dispatch; equality reaching BC at both endpoints; degeneracy
before the ratio with every downstream term left at zero; preparation taking no
RNG state and drawing nothing; **all five accepted Cheng vector families**
reproduced sample by sample — value under tolerance, attempts, consumption,
cumulative consumption and every intermediate RNG state exact; the corpus totals;
the opposite orientations; the prepared terms rebuilt from the projected
literals; the Bernoulli decision table including the `u = p` boundary; one draw
at both probability extremes; and a probability outside `[0,1]` refused with
every output untouched.

**Group E — failure and state atomicity** (`test_35`–`test_44`): the §24 list in
full, including an invalid state refused on all three degenerate paths, an
invalid ordering refused before any draw, an unprepared shape refused, a
rejection proving it advanced the working state by replaying the consumed draws,
a Cheng failure injected mid-walk leaving the caller untouched, and the accepted
result committing exactly the final state.

**Group F — shape conformance** (`test_45`–`test_61`): the only numeric literals
in the module are `0` and `1`; every locked Cheng literal read from the
projection and never re-derived; the convex Uniform form with no Most Likely
anywhere in the procedure; the Triangular conditioning with the **right form in
the right arm**; the conditioning scale; the dispatch boundary with degeneracy
ordered before the ratio; the opposite orientation arms; no prepared term
recomputed in either proposal loop; two uniforms per attempt with no saved state
to restore; the BB acceptance order and the single-expression logit; the BC
operators and the immediate-acceptance arm; the convex Beta rescale with no
clipping; the shared, unmirrored return rule; strict Bernoulli with no extreme
special case; every output committed after its last check with a local working
copy; and the finite predicate borrowed rather than reimplemented.

### 5.2 Mutation controls

`test_00` first asserts the accepted source passes **every** detector. Each
control then damages a copy in a temporary directory, reruns the **whole**
conformance battery, and requires a **named** detector among the refusers. A
five-second per-test budget catches a mutation that fails to terminate.

| # | Mutation | Named detector |
|---|---|---|
| 1 | Uniform reads Most Likely | `test_47` |
| 2 | Uniform degeneracy changed to `a = m = b` | `test_47` |
| 3 | Uniform uses `a + u*(b-a)` | `test_47` |
| 4 | a degenerate distribution consumes a draw | `test_19` |
| 5 | Triangular `<= c` changed to `< c` | `test_48` |
| 6 | Triangular branches swapped | `test_48` |
| 7 | Triangular conditioning removed | `test_48` |
| 8 | invalid ordering silently repaired | `test_38` |
| 9 | PERT lambda independently hardcoded | `test_45` |
| 9a | PERT lambda changed | `test_24` |
| 10 | BB/BC boundary gives equality to BB | `test_50` |
| 11 | BB min/max orientation reversed | `test_51` |
| 12 | BC max/min orientation reversed | `test_51` |
| 13 | a locked BB literal replaced by `Log(4)` | `test_46` |
| 14 | a locked BC decimal replaced by `7/9` | `test_46` |
| 15 | logit split into two logarithms | `test_54` |
| 16 | BB `>=` acceptance tightened to `>` | `test_54` |
| 16a | BB acceptance tests reordered | `test_54` |
| 17 | BC `u1 < 0.5` loosened to `<=` | `test_55` |
| 18 | BC `z <= 0.25` tightened to `<` | `test_55` |
| 19 | BC `z >= k2` changed | `test_55` |
| 20 | final BC `>=` changed | `test_55` |
| 21 | a rejected proposal rewinds the RNG state | `test_53` |
| 22 | a Cheng attempt consumes one uniform | `test_53` |
| 23 | prepared constants recomputed in the proposal loop | `test_52` |
| 24 | Beta orientation return mirrored | `test_57` |
| 25 | unsafe Beta rescale introduced | `test_56` |
| 26 | the Beta variate is clipped | `test_56` |
| 27 | Bernoulli `<` loosened to `<=` | `test_58` |
| 28 | `p = 0` special-cased to zero draws | `test_58` |
| 29 | probability clamped instead of refused | `test_34` |
| 30 | a direct `Rnd` introduced | `test_08` |
| 31 | a generator constant read directly | `test_08` |
| 31a | the D6-11 algorithm token introduced | `test_09` |
| 32 | a worksheet reference introduced | `test_06` |
| 33 | module-level mutable sampler state introduced | `test_07` |
| 33a | a `Static` local introduced | `test_07` |
| 34 | the caller state committed before success | `test_59` |

**Two detectors were strengthened by their own controls.** Mutations 6 and 24
survived the first draft of `test_48` and `test_57`, because those tests asserted
that both branch statements were *present* — which a swap preserves. Both now
split the `If`/`Else` and require the right form in the right arm.

### 5.3 Existing inventory tests

Adding a real module invalidated thirteen tests whose assumption was the prior
exact inventory. None was deleted, skipped or weakened; each was pointed at a
named Phase-6 inventory rather than a count, so the addition stays a **visible
edit**:

```
test_phase4_stage_b_source.py          PHASE6_VBA_MODULES
test_phase5_check_source.py            PHASE6_HANDWRITTEN
test_phase5_gate_b_harness_source.py   _PHASE6_MANIFEST_MODULES / _PHASE6_HANDWRITTEN (new)
test_phase5_report_source.py           PHASE6_INVENTORY (new)
test_phase5_vba_source.py              PHASE6_MODULES
test_phase6_sim_contract.py            "only the authorised Phase-6 VBA exists"
test_phase6_sim_rng.py                 the two authorised modules named
test_phase6_sim_rng_vba.py             the registry tail
test_phase6_sim_sample.py              Cheng is unique to modSimSample
test_phase6_stage_a.py                 both hand-written Phase-6 modules
```

**No numeric module-count authority was added** to PowerShell or to any test —
the Step-6 settlement removed the last one and Step 7 did not reintroduce it.
The historical Run-2 fixtures keep their 15: they prove the old failure mode.

One real defect surfaced from an accepted Phase-5 rule: `width` is a VBA
reserved identifier, and `test_86`/`test_87` of `test_phase5_vba_source.py`
caught it in two declarations. The local is now `span`; the two `detail` message
strings still read "no width", which is prose and not a declaration.

---

## 6. Verification

### 6.1 Python suite

```
2496 passed, 0 failed          (418.44s)
2496 collected
```

| Count | What |
|---|---|
| 61 | Step-7 conformance tests (`test_phase6_sim_sample_vba.py`) |
| 39 | Step-7 mutation controls (`test_phase6_sim_sample_vba_validation.py`) |
| 47 | Step-6 RNG conformance tests — **still green, unmodified** |
| 27 | Step-6 RNG mutation controls — **still green, unmodified** |
| 43 | D6-11 source tests across 16 files |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2396** (`2ec1844`) to **2496**: +100, being the 61
conformance tests and 39 mutation controls. **No test was deleted, skipped or
weakened**; the thirteen inventory tests the new module invalidated were pointed
at a named Phase-6 inventory.

### 6.2 Stage A

```
351 passed, 0 failed
Stage A build complete.
```

`build/PCCM_stageA.xlsx` — 35 736 bytes, `.xlsx`, **no `vbaProject.bin`**, no
`.bas` or `.bin` member of any kind, and no `modSim*` member.
`modSimSample.bas` is source VBA for Stage B; Step 7 does not embed it.

**No Windows and no Excel runtime ran.**

### 6.3 Module registry after Step 7

```
 1. modConstants        generated
 2. modWorkbook
 3. modAppState
 4. modTimeline
 5. modDrivers
 6. modProfiling
 7. modInflation
 8. modStructuralCheck
 9. modCalcContract     generated
10. modCalcFactors
11. modCalcAnalytical
12. modCalcFingerprint
13. modCalcResolve
14. modCalcCheck
15. modCalcReport
16. modSimContract      generated
17. modSimRng
18. modSimSample                       <- new
```

### 6.4 Artefact hashes

| Artefact | SHA-256 | Status |
|---|---|---|
| `src/vba/modSimSample.bas` | `57ed2ada91201b6e0d78648e989edac230d1f284b9f76c45c954c13a3264ef9a` | **new** |
| `build/stage_b_manifest.json` | `d1571ce16ffb815da77b4e18c66579338e11a1da16de2040c8cde1f420c32909` | **changed, registry only** |
| `src/vba/modSimRng.bas` | `a258b0d6628cd1d7bc8a40b712c8b4cc9968bfc96e7985040979dc62527024a9` | **byte-identical to the accepted Step-6 hash** |
| `build/vba/modSimContract.bas` | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` | unchanged |
| `build/phase6_cases.json` | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` | unchanged |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | byte-identical |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | byte-identical |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | byte-identical |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | byte-identical |

`git diff 2ec1844 -- src/vba/modSimRng.bas evidence/ builder/ bootstrap/` is
**empty**: no numerical authority, no evidence and no bootstrap file moved.

### 6.5 The manifest movement is the registry and nothing else

Rebuilding the manifest from the Step-6 contract and comparing leaf by leaf:

```
keys added      3     (.vba.modules[17].name / .generated / .responsibility)
keys removed    0
values changed  0
```

and the structured forbidden-rule projection is **bit-identical**:

```json
{"construct": "MRG32k3a",      "allowed_in": ["modSimRng"]},
{"construct": "RunSimulation", "allowed_in": []}
```

with every other rule still global. Step 7 required no new scoped exception and
took none.

### 6.6 Phase-5 fingerprint vectors

```
fingerprint("PCCM-FP")      6551C6F365DA7F3F
fingerprint_probe(A|B)      42E49DC715F06970
fingerprint_probe(AB|)      7558FD9248656EAD
canonical_number(1/3)       3.3333333333333331E-01
```

Unchanged.

---

## 7. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed and not touched; still
mandatory before Gate-B harness extension or Windows execution.

**Two stale descriptive strings**, carried forward unchanged and outside this
step's boundary:

```
bootstrap/windows/phase4_functional_test.ps1:134   "... inventory back to 15"
docs/phase5_gate_b_harness.md:1015                 "... inventory back to 15"
```

Neither is an executable gate; no check reads either. Both must be cleaned
before the next real Gate-B/Windows execution.

---

## 8. Step-7 acceptance gate — self-check

| Gate condition | Status |
|---|---|
| `modSimSample.bas` exists | yes, 13 procedures, 5 public |
| pure / worksheet-independent | `test_06`, `test_07`, `test_04` |
| no direct RNG implementation | `test_08` |
| all randomness through modSimRng | `test_08` — exactly two entry points |
| invalid state refused even on degenerate paths | `test_35`, `test_36`, `test_37` |
| Uniform ignores ML and degenerates on `a = b` | `test_19`, `test_47` |
| Uniform convex transform | `test_16`, `test_17`, `test_47` |
| Triangular conditioning and `<=` exact | `test_20`, `test_48` |
| Beta preparation before iteration sampling | `test_14`, `test_27`, `test_52` |
| degeneracy before parameterisation | `test_26`, `test_50` |
| BB/BC dispatch boundary exact | `test_25`, `test_50` |
| BB orientation min/max, BC max/min | `test_30`, `test_51` |
| all locked Cheng literals from modSimContract | `test_45`, `test_46` |
| exact logit form preserved | `test_54`, `test_55` |
| two uniforms per Cheng proposal | `test_28`, `test_53` |
| rejection never rewinds | `test_40`, `test_53` |
| consumption reproducible from source | `test_28`, `test_29` — all five families |
| Bernoulli strict `<` | `test_32`, `test_58` |
| `p = 0` and `p = 1` still cost one uniform | `test_33` |
| no Probability clamp | `test_34`, `test_58` |
| no simulation orchestration leaked in | `test_11`–`test_14` |
| `modSimRng` byte-identical | `a258b0d6…` |
| D6-11 scope unchanged | `test_09`, §6.5 |
| registry adds only `modSimSample` | `test_02`, §6.3, §6.5 |
| no Step 8 exists | no `modSimEngine`, `modSimStats`, `modSimFingerprint`, `modSimReport`, `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B temp-dir debt | **OPEN**, §7 |
