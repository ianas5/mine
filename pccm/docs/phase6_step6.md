# PCCM Phase 6 — Step 6 authority record

Step 6 adds the **first Phase-6 source VBA** and performs the **first real
D6-11 scoped-forbidden activation**, atomically, in one commit:

```
src/vba/modSimRng.bas           the pure generator backbone
spec/structure_contract.yaml    modSimContract + modSimRng enter the registry;
                                MRG32k3a becomes scoped to modSimRng
```

Those two changes are one change. A scoped grant landed before its owner would
make the construct legal in a module that does not exist; the owner landed
before the grant would make the module illegal the moment it was written. There
is no ordering of two commits that is correct, so there is one commit.

> **HARNESS SETTLEMENT ROUND.** Independent review of `45d40f2` confirmed the
> Step-6 core — 74 conformance/mutation tests passing, Stage A green, all three
> generated hashes as reported — and did **not** reopen `modSimRng`. It found one
> defect this record had misclassified: P5-M still carried an obsolete
> `$expected.Count -eq 15` gate beside the manifest-derived exact-set check, so a
> **correct** 17-module workbook would fail Gate B. The first submission called
> that cosmetic. It was not. It also corrected the test-count arithmetic. Both
> are fixed here and recorded in §12; `modSimRng.bas` is byte-identical.

**Not in this step.** No `modSimSample`, `modSimEngine`, `modSimStats`,
`modSimFingerprint`, `modSimReport` or `PCCM_RunSimulation`. No sampler, no
iteration loop, no statistic, no quantile, no contingency, no digest, no
`_SimData`, no Results publication, no Gate B, no Windows or Excel runtime.
Step 7 has not begun.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Step-1 contract authority | `35c2467c1f0852fd6cbe5285600c96baeedca2de` and its accepted hardening history |
| Accepted Step-4 oracle head | `614be1acf0f69c16443ace5381edf6157e0f57d3` |
| Accepted Step-5 Stage-A emission head | `fa30424bfe5e782756fc44b0912c4f113781606c` |
| Step-6 modSimRng + D6-11 activation | this commit |

`builder/pccm_builder/sim_rng.py` remains the single definition of these
semantics. `modSimRng.bas` is their VBA implementation and creates no authority:
every operational value it uses is read from `modSimContract`, which projects
`spec/sim_contract.yaml`, `spec/input_contract.yaml` and `spec/workbook.yaml`.

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

This distinction governs every claim in this record and every test added in it.

**SOURCE CONFORMANCE — proved now, on Linux.**
The module is read as text. Its purity, its public surface, its use of the
projected constants, the shape of every locked formula, and the *arithmetic
those formulas describe* are all checked, and the arithmetic is checked against
the accepted Step-0 vectors.

**VBA EXECUTION CONFORMANCE — NOT proved, deferred to Gate B.**
No VBA runtime exists in this step and none was authorised. Type coercion, the
numeric-literal parser, `Fix`, `ByRef` binding, `Long` overflow and array
lower bounds are the Windows interpreter's behaviour, not the source's.

`tests/test_phase6_sim_rng_vba.py` carries that distinction in its own banner
and it is repeated here because it is the easiest thing in this project to
overstate:

> No test in this file may be read as "VBA produced this number".

The retained Step-0 evidence and `build/phase6_cases.json` groups `A_rng`,
`B_jump` and `B_seed` remain the execution-vector authority for that later gate.

### 2.1 How the arithmetic is checked without a VBA runtime

`test_phase6_sim_rng_vba.py` contains a **transcription engine**: it parses the
statements of `modSimRng.bas` — signatures, `Dim`/`ReDim`, `If`/`ElseIf`/`Else`,
`Do While`, `For`/`Next`, `Exit Function`, assignments — and compiles them into
Python, modelling VBA's semantics where they differ from Python's:

| VBA | Modelled as |
|---|---|
| a scalar passed `ByRef` | a one-slot box, so an out-parameter reaches the caller |
| a scalar passed `ByVal` | re-boxed at entry, so the caller cannot be written |
| a UDT assignment (`state = candidate`) | a deep copy into the caller's storage |
| `ReDim` of a `ByRef` array | resized in place, not rebound |
| `Fix` | truncation toward zero |
| `StrComp(a, b, vbBinaryCompare)` | ordinal comparison of UTF-16 code units |
| `LBound(x) + i` | the caller's own lower bound |

Every number it evaluates comes from the projected constants, and every
expression it evaluates is read out of the `.bas` **at test time**. That is why
the mutation controls in §6 work: damage the source and the transcription
computes something else.

It is a transcription of the source, not an interpreter of VBA. It is evidence
about the algorithm the module writes down. It is not evidence about Excel.

---

## 3. `modSimRng.bas`

15 procedures, 7 public.

| Public | Does |
|---|---|
| `SimRngValidateState` | refuses an inadmissible or absorbing state |
| `SimRngStateFromFixedSeed` | D6-05(a): the scalar repeated into all six words |
| `SimRngAutoSeedFromNonce` | D6-03(b): `multiplier ^ nonce mod modulus`, square-and-multiply |
| `SimRngNextUniform` | one recurrence step and the `<=` combination |
| `SimRngJumpNextStream` | the canonical 2^127 jump |
| `SimRngStreamInitialState` | stream *k* by walking the ladder |
| `SimRngBuildComponentStreams` | D6-16 family A: canonical component assignment |

| Private | Does |
|---|---|
| `MRG32k3aStep` | one recurrence step — **the scoped D6-11 token** |
| `SimRngNorm` | the constructed normalisation (§4) |
| `SimRngReduce` | the `Fix`-based modular reduction |
| `SimRngMultModM` | the locked safe modular multiply |
| `SimRngJumpRow` | one matrix row through the safe multiply |
| `SimRngOrderIds` | ordinal insertion sort on an index permutation |
| `SimRngValidTriple`, `SimRngValidWord` | the state predicates |

### 3.1 Deliberate absences

No sampler. No iteration loop. No contribution arithmetic. No statistic. No
quantile. No contingency. No digest. No `_SimData`. No Results. No run state. No
user command. The module produces raw uniforms and stream identities; everything
downstream begins in a later step.

**No global mutable state.** No module-level generator, no `Static` local, no
hidden singleton. Every operational value is an explicit parameter or a returned
typed value, so two callers cannot interfere and a run cannot depend on what
happened before it.

**Worksheet-free by construction.** No `Range`, `Cells`, `Worksheet`,
`Workbook`, `ListObject`, `Application`, `Evaluate`, `MsgBox`, `CreateObject`,
`Environ`, `Shell`, `Date`, `Now`, `Timer` or `DoEvents`, and no public
procedure accepts `As Object`, `As Variant` or any COM type.

### 3.2 Failure semantics

The accepted pure-kernel pattern: a public operation returns `False` and names
the stage in `detail`, and a failing state-changing call leaves the caller's
state exactly as it found it. Results are computed into locals, the candidate is
validated, and only then is the output committed. Nothing displays anything.

---

## 4. THE NORMALISATION CONSTANT IS CONSTRUCTED, NOT SPELLED

**This is the most important implementation decision in Step 6 and the one most
likely to be silently undone by a later editor.**

`modSimContract` projects

```
Public Const SIM_RNG_NORM As Double = 2.328306549295728E-10
```

which is the accepted Double and needs **sixteen** significant digits to name.

**VBA converts a numeric literal at about fifteen.** Phase-5 Gate-B Runtime
Run 3 proved this from the other direction, on `MAX_DOUBLE`, and
`modCalcFactors` records it: a literal that needs more precision than the parser
keeps becomes a *different* Double. The fifteen-digit spelling here is
`2.32830654929573E-10`, which is **four ulp away**. Every uniform drawn through
it would be wrong in the last bits and every downstream vector would miss.

The accepted value is exactly `1 / (m1 + 1)` in binary64. `m1` is `4294967087`,
ten digits, so `SIM_RNG_M1` parses exactly; adding one is exact; IEEE division is
correctly rounded. So the module does not spell the constant — it builds it:

```vba
Private Function SimRngNorm() As Double
    SimRngNorm = 1# / (SIM_RNG_M1 + 1#)
End Function
```

This is the same discipline `modCalcFactors` applies to `MAX_DOUBLE`: build the
value from constants that *do* survive the parser rather than trust a spelling
the parser cannot keep. It is a `Function` rather than a `Const` because a
`Const` initialiser cannot compute.

**`SIM_RNG_NORM` remains the authority.** `test_13` binds the construction to the
projection: it asserts `SimRngNorm() == SIM_RNG_NORM`, that
`SIM_RNG_NORM == 1 / (SIM_RNG_M1 + 1)` exactly, that the fifteen-digit spelling
is a *different* Double exactly four ulp away, that sixteen digits do round-trip,
and that `SIM_RNG_NORM` does not appear in `modSimRng.bas` at all. The
construction cannot drift from the projection without that test failing.

---

## 5. The locked shapes

| Rule | Where | Why it is stated rather than assumed |
|---|---|---|
| **Reduction is `Fix`-based** | `SimRngReduce` | VBA's `Mod` coerces to an integer type and cannot take a value above `Long`. `Int` floors and `Round` rounds; either gives a different `k` for a negative `p`. |
| **The negative-remainder correction** | `SimRngReduce`, `SimRngMultModM` | The recurrence forms a *signed* difference and the mathematics requires the non-negative residue. |
| **One safe modular multiply** | `SimRngMultModM` | Shared by the jump product *and* the AUTO modular power. A second implementation is a second chance to get it wrong. The split constant is `SIM_JUMP_DECOMPOSITION_H`; a second copy of `2^17` would be a second authority. |
| **State is stored oldest-first** | `SimRngState` | `[s10, s11, s12, s20, s21, s22]`, as the contract orders it. `Double`, not `Long`: the words reach `4294967086`, which overflows a signed `Long`, and every value is an exact integer well inside 2^53. |
| **The jump reverses in and out** | `SimRngJumpNextStream` | The matrices operate on **newest-first** triples. A transpose is a different matrix and a dropped reversal is a different vector: both produce a plausible stream that is not the canonical one. |
| **`p1 <= p2`** | `SimRngNextUniform` | `<` would produce a different uniform whenever the two residues are equal. |
| **AUTO is a modular POWER** | `SimRngAutoSeedFromNonce` | Stepping the cycle `nonce` times gives the same answer and is unusable near the period. Stating it as a power is what tells an implementation to square and multiply. |
| **AUTO never wraps** | `SimRngAutoSeedFromNonce` | At `SIM_NONCE_EXHAUSTED` the mathematics returns the seed nonce 0 already issued. Reissuing it silently is the one outcome the lifecycle exists to prevent, so the module refuses. |
| **Stream *k* is algorithmic** | `SimRngStreamInitialState` | 400 is the design-target component count, not a contract cap. The accepted vectors include stream **401** precisely so a table masquerading as an algorithm cannot pass. No substreams: 2^76 is not used in Phase 6. |
| **Kind and role are separate sort keys** | `SimRngBuildComponentStreams` | The Risks interleave per driver — `R-099 occurrence, R-099 severity, R-100 occurrence, R-100 severity` — and not "every occurrence, then every severity". |
| **Ordering is ordinal** | `SimRngOrderIds` | `vbBinaryCompare` on UTF-16 code units. A locale collation, a case fold, a trim or a numeric-suffix reading would all put `CL-999` before `CL-1000`; the accepted order is lexical, so `CL-1000` comes first. |
| **Row order is not an input** | `SimRngOrderIds` | The caller's arrays are never reordered; ordering runs on a private index permutation, exactly as the accepted Phase-5 fingerprint ordering does. Caller arrays are read through `LBound`, so a 1-based array does not shift the assignment. |
| **A duplicate identity is refused** | `SimRngOrderIds` | Never silently deduplicated: two drivers claiming one identity would quietly share a stream. |
| **The ladder is walked once** | `SimRngBuildComponentStreams` | O(N) jumps for N components, not O(N²). |

---

## 6. D6-11 — the first real activation

### 6.1 The rule

`spec/structure_contract.yaml` previously listed `MRG32k3a` as a bare string,
which means *forbidden everywhere*. It is now:

```yaml
    - construct: "MRG32k3a"
      allowed_in:
        - "modSimRng"
```

The scoped `{construct, allowed_in}` shape was added in Phase-6 Step 1 and
deliberately granted to nothing until an owner existed. This is the first and
only grant.

**`RunSimulation` stays global.** Its owner is the later reporting module, and
scoping it now would license the endpoint before the step that authorises it.
Every other forbidden construct is unchanged and still global.

### 6.2 The grant is exercised, not decorative

`MRG32k3aStep` is a real private procedure name in executable code, not a
comment. The source scanners strip comments *and* string literals before
scanning, so a token that lived only in prose would satisfy the rule vacuously.
`test_08` asserts the token survives that stripping and is a member of
`module.procedures`.

### 6.3 The manifest representation

`stage_b_manifest.json` now carries **both**:

```json
"forbidden_constructs": ["Worksheet_Change", ..., "MRG32k3a", ...],
"forbidden_construct_rules": [
  {"construct": "Worksheet_Change", "allowed_in": []},
  ...
  {"construct": "MRG32k3a", "allowed_in": ["modSimRng"]},
  ...
  {"construct": "RunSimulation", "allowed_in": []}
]
```

The flattened list is **display only** now that a scoped rule exists: it cannot
represent `allowed_in`, so a consumer enforcing from it would read the one scoped
construct as globally forbidden and reject the module that owns it. It stays for
backward-compatible display and for an old harness reading an old manifest.
`forbidden_construct_rules` is the enforcement authority; global rules carry an
empty `allowed_in`, so a consumer needs no second source to decide whether a
construct is permitted in the module it is looking at.

### 6.4 Every static consumer is module-aware

Activating the scope broke 25 consumers that assumed "forbidden" meant "forbidden
everywhere". All of them now decide per module, and the P5-EV persisted-project
sweep in `bootstrap/windows/phase5_gate_b_scenarios.ps1` does the same: it
iterates the structured rules and asks `Test-ConstructForbiddenIn` for each
component, with an explicit check that `MRG32k3a` is scoped to `modSimRng` and to
nothing else, and that `RunSimulation` is still global. A manifest carrying only
the flattened list is treated as all-global, so the harness still reads an older
manifest correctly.

---

## 7. The module registry

`modSimContract` and `modSimRng` enter `vba.modules` in this step. The registry
is now, in order:

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
16. modSimContract      generated      <- new
17. modSimRng                          <- new
```

`GENERATED_MODULES` in `structure_loader.py` becomes
`("modConstants", "modCalcContract", "modSimContract")`. `modSimContract` was
emitted in Step 5 but did not enter the registry then: Step 6 is the first
implementation step that depends on it, and the registry is a statement about
what the Stage-B import must carry, not about what the emitter happens to write.

`build_stage_b.ps1` needed no **semantic** change: it iterates
`$manifest.vba.modules` and branches on each entry's own `generated` flag rather
than globbing a directory or holding a list of its own, so both new modules are
imported by the existing generic algorithm. Its **documentation** was stale and
is corrected in §12.3.

**Stage A is unaffected.** The workbook remains `.xlsx` with no
`vbaProject.bin` and no Phase-6 runtime content. Step 6 adds source VBA to the
repository; it does not embed it.

---

## 8. Tests

### 8.1 New files

| File | Contents |
|---|---|
| `tests/test_phase6_sim_rng_vba.py` | 47 Step-6 conformance tests |
| `tests/test_phase6_sim_rng_vba_validation.py` | 27 mutation controls (the 24 required, plus a baseline and two extras) |

**Group A — the module exists and is declared** (`test_01`–`test_04`): the
`VB_Name` attribute and `Option Explicit`; the structure-contract registration
with `generated: false`; the exact seven-name public surface with the numerical
internals private; and no public procedure accepting an object, a COM type or a
`Variant`.

**Group B — purity** (`test_05`–`test_07`): the workbook/environment sweep; no
module-level state, no `Static` local, and exactly the two public `Type`
declarations; and no sampler, engine, statistic, quantile, digest or Results
token leaking in early.

**Group C — D6-11** (`test_08`–`test_12`): the token is in executable code here
and in no other module and not in the generated projection; exactly one construct
is scoped and to exactly one owner; `RunSimulation` is not scoped; every other
forbidden construct still applies to `modSimRng` *and* is absent from it; and the
manifest carries the structured rules with the flattened list still covering the
same construct set.

**Group D — source arithmetic conformance** (`test_13`–`test_32`, `test_46`,
`test_47`): the transcription against the accepted Step-0 vectors — the
constructed normalisation, FIXED seeding and its domain, the AUTO mapping and
both ends of its domain, the first five uniforms for every accepted seed, twenty
draws and the state they leave, every accepted jump stream (states *and*
uniforms), stream *k* including **401**, the 400-component assignment against
`family_a_first_10` and `family_a_last_4`, per-component stream identity, Risk
interleaving, ordinal ordering, caller arrays untouched, row-order invariance,
duplicate and blank refusal, every inadmissible-word refusal, state left
untouched on a failed draw and on a failed jump, 2000 draws strictly inside
(0, 1) with no repeat, the AUTO call count bounded by `2·log2(nonce) + 2`, and
exactly `N − 1` jumps for `N` components.

**Group E — shape conformance** (`test_33`–`test_45`): no owned value restated as
a literal (the only numerals in the module are `0`, `1` and `2`); `Mod`, `Int`,
`Round` and `CInt` absent and the `Fix` reduction present; the negative-remainder
correction in both primitives; the split constant used exactly three times and
nowhere duplicated; each recurrence coefficient paired with its own word; the
`<=` boundary and both combination branches; all twelve reversal statements and
all eighteen matrix elements read from the projection; no table, `Array(`,
`Choose(`, `Switch(`, `Select Case`, substream or `400`; `vbBinaryCompare` and no
identity-parsing token; every `SIM_` identifier owned by the projection and no
local `Const`; every output committed after its last check; caller arrays read
through `LBound`; and the transcription having read every procedure and both
types.

### 8.2 Mutation controls

`test_00` first asserts the accepted source passes **every** detector, so the
controls below are not measuring noise. Each control then damages a copy in a
temporary directory, reruns the **whole** conformance battery against it, and
requires (a) at least one detector to refuse and (b) a **named** detector to be
among the refusers — a control cannot quietly degrade into "something, somewhere,
went red". A five-second per-test budget means a mutation that fails to terminate
is caught as a refusal rather than hanging the suite. Nothing is written to the
repository.

| # | Mutation | Named detector |
|---|---|---|
| 1 | `a12` changed | `test_37` |
| 2 | `a13n` changed | `test_37` |
| 3 | `m1`/`m2` swapped | `test_37` |
| 4 | `<=` changed to `<` | `test_38` |
| 5 | `Fix` reduction replaced by `Mod` | `test_34` |
| 6 | negative-remainder correction deleted (`SimRngReduce`) | `test_35` |
| 6a | negative-remainder correction deleted (`SimRngMultModM`) | `test_35` |
| 7 | seed mixer introduced | `test_14` |
| 8 | AUTO changed from modular power to linear stepping | `test_46` |
| 9 | unsafe multiplication used in the modular exponent | `test_16` |
| 10 | jump safe multiply replaced by a naive product | `test_39` |
| 11 | input matrix reversal removed | `test_39` |
| 12 | output reversal removed | `test_39` |
| 13 | matrix element transposed | `test_39` |
| 14 | precomputed stream table introduced (a 400 cap) | `test_21` |
| 15 | physical input order used directly | `test_25` |
| 16 | numeric Permanent-ID suffix sort introduced | `test_41` |
| 17 | Risk occurrence/severity not interleaved | `test_24` |
| 18 | duplicate component silently accepted | `test_28` |
| 19 | worksheet reference introduced | `test_05` |
| 20 | global mutable generator state introduced | `test_06` |
| 20a | `Static` local introduced | `test_06` |
| 21 | `Rnd` introduced | `test_05` |
| 22 | MRG token placed in another module | `test_09` |
| 23 | `allowed_in` widened | `test_10` |
| 24 | `RunSimulation` scoped prematurely | `test_10` |

Controls 22–24 damage the **authority**, not the source: 22 smuggles the token
into a copy of `modCalcFactors.bas`, 23 adds a second module to the grant, and 24
converts the bare `RunSimulation` entry into a scoped rule. All three are applied
to a temporary copy of `spec/` or `src/vba/`.

### 8.3 Consumers made module-aware

Activating the scope turned 25 previously-passing enforcement tests red, because
every one of them asserted that a forbidden construct is forbidden *everywhere*.
None was deleted, skipped or weakened; all now decide per module.

| File | Change |
|---|---|
| `test_phase4_stage_b_source.py` | module-aware forbidden sweep; `PHASE6_VBA_MODULES`; UDT members recognised as declarations; `emit_sim_artifacts` in the emitted-artefact set; new `test_15a` |
| `test_phase4_structure_contract_validation.py` | the scoped-rule mutation expectation |
| `test_phase5_gate_b_harness_source.py` | `_scan_production_against_rules()`; the P5-EV manifest module list |
| `test_phase6_sim_contract.py`, `test_phase6_sim_contract_validation.py` | the registry now contains `modSimContract` |
| `test_phase6_stage_a.py` | `test_21`/`test_22` inverted: the registry entry and the grant now *must* exist |
| `test_phase5_gate_b_harness_source.py` (settlement round) | `test_38` rebuilt, `test_38a`/`test_38b` added — see §12.4 |

### 8.4 Boundary excess, disclosed

Section 22 of the Step-6 authorisation enumerated the test files expected to
change and added "Do not assume this list is exhaustive." Six files outside that
enumeration also required edits, all of them for the same reason — they hold a
frozen inventory of VBA modules or a global forbidden-construct assumption, and
the activation must be atomic:

```
tests/test_phase5_check_source.py      PHASE6_HANDWRITTEN
tests/test_phase5_report_source.py     PHASE5_INVENTORY frozen by name, not count
tests/test_phase5_resolve_source.py    module-aware construct sweep
tests/test_phase5_vba_source.py        PHASE6_MODULES
tests/test_phase6_sim_rng.py           the projection's module registration
tests/test_phase6_sim_sample.py        the projection's module registration
```

`tests/test_phase4_structure.py` was allowed by §22 and needed no change.

### 8.5 The P5-M module-inventory gate — CORRECTED IN THE SETTLEMENT ROUND

The first submission of this step reported the P5-M "15 modules" wording as
**cosmetic**. That was wrong, and §12 records the correction: the wording was the
visible half of an executable `$expected.Count -eq 15` gate that would have failed
a correct 17-module workbook.

---

## 9. GATE-B TEMP-DIR CLEANUP DEBT — OPEN

Carried forward unchanged and **not** addressed in this step, as instructed. The
temp-build helper and its directory lifecycle were not altered. The debt remains
mandatory before Gate-B harness extension or Windows execution.

---

## 10. Verification

### 10.1 Python suite

```
2396 passed, 0 failed          (305.72s)
```

**The collection arithmetic, corrected.** The first submission said the parent
baseline was 2320. It was **2319**. Derived by collecting both revisions and
diffing the node IDs, not by subtraction:

| | Collected |
|---|---|
| Accepted Step-5 package `fa30424` | **2319** |
| Step-6 package `45d40f2` | **2394** |
| Step-6 after the settlement round | **2396** |

Composition of the +75 from `fa30424` to `45d40f2`:

```
+47  tests/test_phase6_sim_rng_vba.py             Step-6 conformance
+27  tests/test_phase6_sim_rng_vba_validation.py  Step-6 mutation controls
 +1  tests/test_phase4_stage_b_source.py::test_15a_the_one_scoped_grant_is_real_and_is_the_only_one
     an additional D6-11 scoped-source conformance test in an existing suite
---
+75
```

and the +2 the settlement round adds (§12.4):

```
+1  test_phase5_gate_b_harness_source.py::test_38a_p5_d8_returns_the_inventory_to_the_manifest_set_not_a_count
+1  test_phase5_gate_b_harness_source.py::test_38b_reintroducing_a_hard_coded_module_count_is_refused
---
+2      total from fa30424: +77
```

Nine further tests were **renamed in place**, not added: their subject inverted
from "no Phase-6 VBA exists" / "no D6-11 exception was granted" to "only the
authorised Phase-6 VBA exists" / "the grant exists and is the only one". Nine
node IDs left the collection and nine entered it, netting zero:

```
test_phase4_stage_b_source.py       test_15_no_forbidden_construct_appears_in_phase_4_vba
                                 -> test_15_no_forbidden_construct_appears_where_it_is_forbidden
test_phase5_report_source.py        test_02_the_final_phase_5_inventory_is_fifteen_modules
                                 -> test_02_the_phase_5_inventory_is_complete_and_nothing_of_phase_5_moved
test_phase6_sim_contract.py         test_50_no_phase6_vba_or_emission_exists
                                 -> test_50_only_the_authorised_phase6_vba_exists
test_phase6_sim_contract.py         test_67_d6_11_activation_precondition_is_recorded
                                 -> test_67_d6_11_activated_exactly_once_and_only_with_its_owner
test_phase6_sim_contract_validation.py  test_77_the_global_scalar_entry_still_works
                                 -> test_77_the_global_scalar_entry_still_works_beside_the_one_grant
test_phase6_sim_rng.py              test_36_no_phase6_vba_exists
                                 -> test_36_the_only_phase6_vba_is_the_generator_backbone
test_phase6_sim_sample.py           test_43_no_phase6_vba_exists
                                 -> test_43_no_sampler_vba_exists
test_phase6_stage_a.py              test_21_no_d6_11_exception_was_granted
                                 -> test_21_the_generated_module_needs_no_d6_11_exception
test_phase6_stage_a.py              test_22_the_module_is_not_in_the_stage_b_module_registry_yet
                                 -> test_22_the_module_is_in_the_stage_b_registry_as_a_generated_module
```

**No test was deleted, skipped or weakened.** The 25 enforcement tests the
activation turned red were made module-aware; the nine above changed their
subject because the fact they assert changed.

| Count | What |
|---|---|
| 47 | Step-6 conformance tests (`test_phase6_sim_rng_vba.py`) |
| 27 | Step-6 mutation controls (`test_phase6_sim_rng_vba_validation.py`) |
| 41 | D6-11 source tests across 14 files (see §10.6) |
| 351 | Stage-A builder/verifier checks, 0 failed |

### 10.2 Stage A

```
351 passed, 0 failed
Stage A build complete.
```

`build/PCCM_stageA.xlsx` — 35 734 bytes, `.xlsx`, **no `vbaProject.bin`**, no
`.bas` and no `.bin` member of any kind. Step 6 adds source VBA to the repository
and does not embed it.

**No Windows and no Excel runtime ran.** P5-EV was not executed.

### 10.3 Artefact hashes

| Artefact | SHA-256 | Status |
|---|---|---|
| `src/vba/modSimRng.bas` | `a258b0d6628cd1d7bc8a40b712c8b4cc9968bfc96e7985040979dc62527024a9` | **new** |
| `build/stage_b_manifest.json` | `ba64ce4647bb5c04a6d3f983ff5e4a38045107a4cdded942258fd9f08457be2e` | **changed, expected** |
| `build/vba/modSimContract.bas` | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` | **unchanged — matches the accepted Step-5 hash** |
| `build/phase6_cases.json` | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` | **unchanged — matches the accepted Step-5 hash** |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | byte-identical |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | byte-identical |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | byte-identical |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | byte-identical |

Byte-identity was checked with `cmp` against copies taken before Step 6 began,
not by re-deriving the hashes from the same run.

### 10.4 The manifest movement is additive only

`stage_b_manifest.json` moved from `c9d7a50dd0b56ff94ac462c99b90466e5ac6b062876004ad3a6b151f615939ce`
to `ba64ce46…`. Compared leaf by leaf against the pre-Step-6 manifest:

```
keys added      18
keys removed     0
values changed   0
added roots     .vba.modules  .vba.forbidden_construct_rules
```

Every pre-existing key kept its pre-existing value. The movement is exactly the
two new registry entries and the structured-rule projection, and nothing else.

### 10.5 The scoped rule as the manifest carries it

```json
{"construct": "Worksheet_Change",      "allowed_in": []},
{"construct": "Workbook_SheetChange",  "allowed_in": []},
{"construct": "Worksheet_Calculate",   "allowed_in": []},
{"construct": "Workbook_Open",         "allowed_in": []},
{"construct": "Rnd(",                  "allowed_in": []},
{"construct": "Randomize",             "allowed_in": []},
{"construct": "MRG32k3a",              "allowed_in": ["modSimRng"]},
{"construct": "NPV",                   "allowed_in": []},
{"construct": "Percentile",            "allowed_in": []},
{"construct": "RunSimulation",         "allowed_in": []},
{"construct": "FinalReleaseComObject", "allowed_in": []}
```

One scoped construct. One owner. `RunSimulation` still global.

### 10.6 D6-11 source-test distribution

```
 5  tests/test_phase4_stage_b_source.py
 1  tests/test_phase4_structure.py
 1  tests/test_phase4_structure_contract_validation.py
 4  tests/test_phase5_gate_b_harness_source.py
 1  tests/test_phase5_report_source.py
 1  tests/test_phase5_stage_a.py
 1  tests/test_phase5_vba_source.py
 2  tests/test_phase6_sim_contract.py
 9  tests/test_phase6_sim_contract_validation.py
 1  tests/test_phase6_sim_rng.py
 7  tests/test_phase6_sim_rng_vba.py
 3  tests/test_phase6_sim_rng_vba_validation.py
 1  tests/test_phase6_sim_sample.py
 4  tests/test_phase6_stage_a.py
41  TOTAL
```

### 10.7 Phase-5 fingerprint vectors

```
fingerprint("PCCM-FP")      6551C6F365DA7F3F
fingerprint_probe(A|B)      42E49DC715F06970
fingerprint_probe(AB|)      7558FD9248656EAD
canonical_number(1/3)       3.3333333333333331E-01
```

Unchanged.

---

## 11. Step-6 acceptance gate — self-check

| Gate condition | Status |
|---|---|
| `modSimRng.bas` exists | yes, 565 lines, 15 procedures |
| pure and worksheet-independent | `test_04`, `test_05`, `test_06` |
| state words `Double`, seed/nonce `Long` | `SimRngState`; `SimRngStateFromFixedSeed(seed As Long)`; `SimRngAutoSeedFromNonce(nonce As Long, seed As Long)` |
| FIXED repeated-scalar seeding | `test_14` against every accepted example |
| AUTO modular power, O(log nonce) | `test_16`, `test_46` |
| base recurrence uses `Fix`, not `Mod` | `test_34` |
| the exact `<=` combination | `test_38` |
| safe `MultModM` implemented once and reused | `test_36` |
| 2^127 jump, every element, correct orientation | `test_39`, `test_20` |
| no substreams | `test_40` |
| stream *k* algorithmic, including *k* > 400 | `test_21` against stream 401 |
| component assignment row-order invariant | `test_26`, `test_27` |
| Risk occurrence/severity interleaved per Risk | `test_24` |
| no numeric Permanent-ID parsing | `test_25`, `test_41` |
| no global mutable RNG state | `test_06` |
| no sampler | `test_07` |
| D6-11 MRG scope active only for `modSimRng` | `test_08`, `test_09`, `test_10` |
| `RunSimulation` remains globally forbidden | `test_10`, `test_11` |
| all static consumers module-aware | 2394 passed |
| P5-EV source module-aware | `test_phase5_gate_b_harness_source.py` |
| `stage_b_manifest` preserves scoped rules | `test_12`, §10.5 |
| `modSimContract` in the Stage-B registry | §7 |
| no Step 7 | no `modSimSample`, `modSimEngine`, `modSimStats`, `modSimFingerprint`, `modSimReport`, `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B temp-dir debt | **OPEN**, §9 and §12.6 |

### Settlement-round targets

| Target | Status |
|---|---|
| P5-M contains no fixed production-module count | §12.1, proved by `test_38` (D) |
| P5-M inventory manifest-driven and bidirectional | §12.1, proved by `test_38` (A)(B)(C) |
| P5-D8 current wording count-independent | §12.2, proved by `test_38a` |
| historical Run-2 `15` evidence preserved | §12.2 |
| `build_stage_b` documentation manifest-accurate | §12.3, algorithm unchanged |
| Step-6 test-count record arithmetically correct | §10.1 |
| `modSimRng.bas` byte-identical | `a258b0d6…`, §12.6 |
| D6-11 scope unchanged | §12.6 |
| all tests pass | 2396 passed, 0 failed |
| Stage A passes | 351 passed, 0 failed |
| no Step 7 exists | no `modSim{Sample,Engine,Stats,Fingerprint,Report}`, no `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |

---

## 12. Harness settlement round

Independent review of `45d40f2` did not reopen `modSimRng` and confirmed the
Step-6 core: 74 conformance/mutation tests passing, Stage A 351/0, and all three
generated hashes as reported. Four things needed correcting.

### 12.1 P5-M carried a second authority — and the first record misclassified it

The exact-set inventory helper `Add-Phase5ModuleInventoryChecks` **was** already
manifest-derived: it proves the persisted standard-module set equals
`manifest.vba.modules` by name, in both directions, with no tolerance.

But P5-M carried an **additional, executable** gate beside it:

```powershell
$expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
$null = Add-Check $list 'the manifest declares 15 production modules' `
    ($expected.Count -eq 15) ("declared " + $expected.Count)
```

After Step 6 the accepted manifest declares **17** production standard modules,
so that check is guaranteed to fail on a **correct** workbook. The harness would
have disagreed with the contract, and the contract would have been right.

The first submission of this record described the issue as "cosmetic staleness"
in the result *title*. It was not cosmetic; the title was the visible half of an
executable gate. **That misclassification is withdrawn.**

**The gate is removed, not renumbered.** `-eq 17` would be the identical defect
one module later, so no production-module count literal appears in the block at
all — the defect is the second authority, not the number 15. What P5-M now
states about the declared list is that it is well-formed (no blank name, no
duplicate) and what its members are; the **set** is proved by the helper.

```powershell
$expected    = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
$blankNames  = @($expected | Where-Object { [string]::IsNullOrWhiteSpace($_) })
$uniqueNames = @($expected | Sort-Object -Unique)
$null = Add-Check $list 'the manifest names a well-formed production module set' `
    (($blankNames.Count -eq 0) -and ($expected.Count -eq $uniqueNames.Count)) `
    ("declared: " + ($expected -join ', '))
```

A future legitimate module addition now requires **no** edit to the Gate-B
harness.

### 12.2 Current result text is count-independent; historical evidence is not touched

| | Before | After |
|---|---|---|
| P5-M title | `Persisted project: 15 modules by name, 5 buttons, 6 API procedures` | `Persisted project: manifest module set by name, 5 buttons, 6 API procedures` |
| P5-D8 title | `Transient diagnostic module removed; inventory back to 15` | `Transient diagnostic module removed; inventory back to the manifest module set` |
| P5-D8 comment | "must return to exactly the 15 manifest modules" | "must return to exactly the manifest-declared production modules" |
| inventory header | "Nothing is weakened to *at least 15*" | "Nothing is weakened to a count comparison of any kind" |

The button and API-procedure counts stay: they are five and six by contract and
were not part of this defect.

**Historical Run-2 evidence is preserved verbatim.** These are records of what
Run 2 actually reported and remain unchanged:

```
FAIL the inventory is exactly the 15 manifest modules again -- present 30 of 15
15 standard modules + 14 sheet documents + 1 ThisWorkbook = 30 components
All fifteen production modules were individually confirmed present ...
P5-M PASS fifteen modules present, and six API procedures reported callable
          - under the evidence model P5-M had at the time            (Run 7)
```

One Run-2 sentence gained a clarifier — "the manifest's vba.modules collection,
**which held fifteen entries at the time**" — so a reader cannot mistake the
historical figure for a current one. The tests that model the historical Run-2
topology (15 standard modules + 14 sheet documents + ThisWorkbook = 30
components) keep their 15: they are proving the old failure mode, not the current
inventory.

### 12.3 `build_stage_b.ps1` documentation

The import algorithm was already generic and is **unchanged**:

```powershell
foreach ($m in $manifest.vba.modules) {
    $dir = $srcDir
    if ($m.generated) { $dir = $genDir }
    ...
}
```

Only the comments, help block and one step label moved:

- "PCCM_stageA.xlsx plus **two generated inputs**" → the workbook, the manifest,
  and the generated VBA projections;
- `build/vba/modConstants.bas — the generated VBA constants module` →
  `build/vba/*.bas — the generated VBA projection modules`, with a note that a
  module is generated when the **manifest** says so, and that the current
  inventory — `modConstants`, `modCalcContract`, `modSimContract` — is an
  inventory, not a dependency;
- `Import the Phase-4 VBA modules` → `Import every manifest-declared VBA module`;
- step 5 of the ordered description → "import every manifest-declared VBA
  module, source and generated alike".

No execution dependency on exactly three generated modules was introduced.

### 12.4 The test that pinned the defect

`tests/test_phase5_gate_b_harness_source.py` asserted the literal
`'the manifest declares 15 production modules'` was present in the harness — it
was protecting the bug. It is replaced by four future-proof requirements plus a
mutation control:

| Test | Proves |
|---|---|
| `test_38` (A) | P5-M derives `$expected` from `$Manifest.vba.modules` |
| `test_38` (B) | P5-M calls `Add-Phase5ModuleInventoryChecks` with `-ExpectedModules $expected` |
| `test_38` (C) | the helper requires exact set equality by name and rejects stray standard modules, in both directions — **unweakened** |
| `test_38` (D) | the **active** P5-M block contains no production-module count literal, at any value |
| `test_38a` | P5-D8 goes through the same helper, carries no count literal, and both current result titles are count-independent |
| `test_38b` | the mutation control: reintroducing `-eq 17`, `-ge 17` **or** `-eq 15` is detected, and the accepted source is clean |

The rejection pattern matches `$expected.Count -<op> <integer>` for any
comparison operator and any integer, and `manifest declares <integer> production
modules`. It runs over `_executable(SCENARIOS)`, which strips comments, so the
explanatory comment naming the removed `-eq 15` does not satisfy or trip it.

### 12.5 Out of boundary, disclosed

Two further files still carry the number 15 in **current** display text and were
**not** touched, because §10 of the settlement authorisation does not allow them:

```
bootstrap/windows/phase4_functional_test.ps1:134
    "P5-D8  The diagnostic module REMOVED and the inventory back to 15"
docs/phase5_gate_b_harness.md:1015
    "| P5-D8 | The diagnostic module removed, inventory back to 15 |"
```

Both are descriptive text about P5-D8, not executable gates — no check reads
either. **Flagged for whichever step next opens those files.**

`tests/test_phase6_sim_contract_validation.py::test_78` had its stale docstring
corrected (§7 of the settlement authorisation, optional and non-semantic): the
fixture stays synthetic and unchanged, and the text no longer says `modSimRng`
does not exist.

### 12.6 Unchanged by this round

`src/vba/modSimRng.bas` is **byte-identical** to `45d40f2`. `spec/**`,
`evidence/**` and `builder/**` are untouched, so no contract or manifest
authority moved and every generated hash is unchanged. D6-11 remains exactly as
accepted: `MRG32k3a → [modSimRng]`, `RunSimulation` global, structured rules the
enforcement authority, P5-EV module-aware.

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** This round fixed an inventory
assertion. It did not touch the temp-directory lifecycle, and the debt must still
close before Gate-B harness extension or Windows execution.

---

## 13. Pre-Step-8 dependency correction — the empty component model

> This section was appended after Step 6 was accepted. It corrects an
> **inherited Step-6 defect** in `SimRngBuildComponentStreams` found while the
> Step-7 dependency boundary was being checked. No other Step-6 semantic moved:
> state validation, FIXED seeding, the AUTO modular power, the recurrence, the
> constructed `SimRngNorm`, the safe modular multiply, the jump, stream *k*,
> Permanent-ID ordering, duplicate refusal, Risk interleaving and the D6-11
> scope are all exactly as accepted, and every Step-6 vector is unchanged.

### 13.1 The defect: an invented business prerequisite

The module contained:

```vba
total = costCount + 2 * riskCount
If total = 0 Then
    detail = "components: the model declares no driver"
    Exit Function
End If
```

**No accepted contract requires a Cost Line or a Risk to exist.** The accepted
Phase-5 source tests pin that an empty driver set is *not* refused once the
model-level prerequisites resolve, and Phase 6 introduced no minimum of its own.
The accepted Python reference agrees: `components_for((), ())` is the empty
tuple, and `component_stream_states` validates the base state, produces nothing
and jumps nowhere.

So the refusal was a business rule the module invented, and it put the VBA at
odds with the accepted Python reference **before Step 8 had begun**.

### 13.2 The corrected semantics

For `costCount = 0` and `riskCount = 0`:

| | |
|---|---|
| negative counts | still refused, exactly as before |
| the base RNG state | **still validated** — an empty component set is not permission to accept a state the recurrence cannot legally be in |
| `costIds()` / `riskIds()` | **never touched**. No `LBound` is read, no ordering runs, no identity is inspected |
| components constructed | none |
| stream jumps | zero |
| the caller's base state | unchanged |
| the result | **True** |

The refusal was **removed, not renumbered**: there is no `total >= 1`, no "at
least one Cost Line", no "at least one Risk", and no minimum of any other
spelling.

### 13.3 The carrier convention, and why there is no dummy component

VBA has no zero-length dynamic array. The module follows the **accepted Phase-5
zero-count carrier convention** — the one `CalcFpSortedRecords` uses and
`SimRngOrderIds` already used inside this very module:

> the output is sized to one slot, and **the logical count** —
> `costCount + 2 * riskCount`, which the caller supplied — decides whether any
> element may be inspected. At zero, no element is semantically present.

The slot is left at its `Type` defaults and **nothing is written into it**. It
is not a component and cannot be mistaken for one: its `PermanentId` is the
empty string, which is exactly what `SimRngOrderIds` refuses. No stream index is
assigned to it — `StreamIndex` is zero because that is the `Long` default, not
because anything set it.

It is *assigned* rather than left alone so that a caller cannot retain an
earlier, longer result and read it back as this answer. The public API is
otherwise unchanged; no out-parameter was added for this correction.

### 13.4 Tests

`test_phase6_sim_rng_vba.py` grew from 47 to 51:

| Test | Proves |
|---|---|
| `test_29` | *renamed*: blank Cost Line id, blank Risk id and a negative count are still refused |
| `test_29a` | zero drivers succeeds — zero jumps, base state untouched, and the carrier slot holds no component |
| `test_29b` | the same empty model with an **invalid** base state is refused, and the source validates the state before taking the empty path |
| `test_29c` | the empty path reads no bound from either driver array and no minimum-driver spelling survives in the body |
| `test_29d` | the VBA transcription and `sim_rng.py` agree directly: empty component tuple, no stream state, base state returned as supplied, and both refuse the same inadmissible state |

`test_43` now compares the **last** commit against its guard, because the
procedure legitimately commits the carrier on two paths — the zero-component one
and the ladder — and asserts separately that the zero-component commit is
governed by the state check.

`test_phase6_sim_rng_vba_validation.py` grew from 27 to 32, with five controls
**scoped to `SimRngBuildComponentStreams`** so a comparison against zero
elsewhere in the module stays legal:

| # | Mutation | Named detector |
|---|---|---|
| 25 | the zero-component refusal reintroduced | `test_29a` |
| 26 | a minimum invented in four other spellings | `test_29a` |
| 27 | the empty path skips base-state validation | `test_29b` |
| 27a | the empty path orders the driver arrays anyway | `test_29c` |
| 27b | the empty path advances the base state | `test_29a` |

### 13.5 Verification of this correction

```
2530 passed, 0 failed          (389.17s)
2530 collected
Stage A: 351 passed, 0 failed
```

| Artefact | SHA-256 | |
|---|---|---|
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` | changed by this correction; the hashes recorded in §10.3 and §11 remain the record of commit `2ec1844`, where they were accurate |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` | **byte-identical** — no Step-7 semantic reopened |
| `build/vba/modSimContract.bas` | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` | unchanged |
| `build/phase6_cases.json` | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` | unchanged |
| `build/stage_b_manifest.json` | `d1571ce16ffb815da77b4e18c66579338e11a1da16de2040c8cde1f420c32909` | unchanged — neither the module registry nor the D6-11 rules moved |

`spec/`, `evidence/`, `builder/` and `bootstrap/` are identical to `f2f654e`.
The module is now 606 lines and still 15 procedures, 7 public: the correction
changed one branch inside one procedure and added no procedure and no
parameter.

**Every prior Step-6 RNG vector is unchanged** — the 51 conformance tests
include the FIXED seed, AUTO nonce, first-5, first-20, state-after-20, jump
streams 0/1/7/399/401 and 400-component assignment checks, and all pass
unmodified.

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN**, untouched, as are the two stale `"15"`
display strings.
