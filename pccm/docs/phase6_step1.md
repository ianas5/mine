# PCCM Phase 6 — Step 1 authority record

Step 1 is the **contract step**. It creates `spec/sim_contract.yaml`, discharges
the input contract's deferred seed note, extends the forbidden-construct schema
without granting anything, closes D6-08 from the `_SimData` layout, and adds the
loader, validator and static tests that make all of it enforceable.

**No simulation implementation exists.** No RNG, no sampler, no Monte Carlo, no
VBA, no Stage-A emission, no workbook change, no Windows or Excel runtime.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Phase-6 Step-0 authority / evidence head | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| First Step-1 contract | `1be66d3f74c62377d26d6ddaa0741fc66be60684` |
| Step-1 contract completion | this commit — reported by the delivery message and `PROVENANCE.txt` |

### What the completion round corrected

Independent review of `1be66d3` returned NOT YET ACCEPTED with no Step-0
decision reopened and no RNG or Cheng change requested. Thirteen
contract-completeness, fail-loud and authority-binding defects are corrected.

| # | Defect | Fix |
|---|---|---|
| 1 | the state predicates were **neither exclusive nor complete**, and read the attempt history | derivation rewritten, ordered and attempt-blind — §4 |
| 2 | unknown semantic keys were **silently accepted** in six demonstrated places | closed-world schema over all 77 mappings — §5 |
| 3 | nothing defined what a driver actually **contributes** | new `contribution` section — §6 |
| 4 | the inherited **hot-kernel boundary** was unencoded | new `kernel` section — §7 |
| 5 | only pieces of plan §4.6 reached the contract | new `numerical_domain` section — §7 |
| 6 | driver **independence** was never stated | new `dependence` section — §7 |
| 7 | only 4 percentiles were retained, not the full ladder | retained **by reference**, 11 today — §8 |
| 8 | the Run Stamp was missing **model version** | added; D6-08 **re-derived** — §9 |
| 9 | two authority bindings only "resolved" and were **false** | content-bound — §10 |
| 10 | `result_digest.tolerance: null` remained | key removed; the scanner exception removed — §11 |
| 11 | publication / command surface / cancellation unencoded | three new sections — §7 |
| 12 | D6-11 had no **activation precondition** | recorded — §12 |
| 13 | the Stage-A claim was **too broad** | corrected — §13 |

**Step 1 encodes authority; it creates no modelling semantics.** Every value in
`sim_contract.yaml` traces to Step 0 or to an accepted earlier contract. Where
the two could disagree, the loader holds the Step-0 value as a locked constant
and the contract is checked against it — the contract *encodes* the accepted
authority, it does not choose it. **No contradiction between sources was found**,
so nothing was resolved by inventing a Step-1 rule.

---

## 2. Files

| File | Change |
|---|---|
| `spec/sim_contract.yaml` | **NEW** — the sixth authority, 1072 lines, 30 sections |
| `spec/input_contract.yaml` | seed admissibility resolved; deferred note discharged |
| `builder/pccm_builder/sim_loader.py` | **NEW** — loader, validator, cross-validator |
| `builder/pccm_builder/structure_loader.py` | D6-11 mixed scalar-or-scoped schema |
| `builder/pccm_builder/__init__.py` | exports |
| `builder/build_stage_a.py` | loads and cross-validates the sim contract; **emits nothing** |
| `tests/test_phase6_sim_contract.py` | **NEW** — 67 positive / conformance tests |
| `tests/test_phase6_sim_contract_validation.py` | **NEW** — 114 mutation controls |
| `tests/test_phase2_inputs.py` | `test_19c` narrowed, `test_19d` added — see §9 |
| `docs/phase6_step1.md` | **NEW** — this record |
| `docs/phase6_plan.md` | §10.3 **post-acceptance authority correction** only — see §4 |

**`spec/structure_contract.yaml` is unchanged.** The D6-11 capability lives in
the loader; the contract's entries stay in their existing scalar shape, so
nothing was granted and no existing protection moved.

**`spec/workbook.yaml` is unchanged.** §18's authority conflict did not arise:
`sim_contract.yaml` owns the future `_SimData` machine layout until the
publication step materialises it. Column widths remain presentation and remain
`workbook.yaml`'s; they are adjusted when the layout is materialised, not now.

---

## 3. Contract sections and where each rule comes from

| Section | Owns | Source |
|---|---|---|
| `versions` | `RNG_VERSION`, `SIM_METHOD_VERSION`, the bump-ownership table | Step 0 §9 |
| `rng` | MRG32k3a constants, state orientation, recurrence, combination, output domain, arithmetic obligation | Step 0 §5.1, §5.6 |
| `seeding` | modes, scalar→state, AUTO mapping, nonce lifecycle | D6-03, D6-05 |
| `components` | `COST_SAMPLE`, `RISK_OCCURRENCE`, `RISK_SEVERITY`, `C + 2R` | D6-16 |
| `stream_assignment` | canonical sorted order, ordinal UTF-16, the preserved consequence | D6-16 |
| `jump` | `2^127`, both matrices, `H = 2^17`, no substreams | Step 0 §5.6 |
| `distributions` | three families, the degenerate rule, Uniform, Triangular, Beta-PERT | plan §4.0–§4.3 |
| `cheng` | the exact locked formulation | Step 0 §5.2a |
| `risk` | occurrence and D6-18b severity semantics | Step 0 §6.1 |
| `accumulation` | canonical driver order, both accumulators | plan §13 |
| **`contribution`** | **what a Cost Line and a Risk contribute to an iteration** | **plan §4.1–§4.3, Phase-5 analytical semantics** |
| **`kernel`** | **the zero-worksheet, zero-COM hot-loop boundary** | **plan §5, inherited** |
| **`numerical_domain`** | **the inherited Double domain and the safe-arithmetic disciplines** | **plan §4.6** |
| **`dependence`** | **drivers are sampled independently; no correlation authority exists** | **plan §13** |
| **`publication`** | **`_SimData` is the source of truth; commit-last; no partial publish** | **plan §11, §16** |
| **`command_surface`**, **`interruption`** | **automation endpoint only; no UI; no cancellation** | **plan §12, §16** |
| `request_fingerprint` | `HEADER·COST·RISK·SIM`, prefix semantics, SIM fields | plan §11 |
| `result_digest` | `PCCM-RD` framing | D6-17 |
| `iterations` | the **technical** ceiling only | D6-08 |
| `sim_data` | the future machine layout, and therefore `H` | D6-08, plan §11.1 |
| `label_sets`, `sim_state` | exactly three states, and the corrected derivation | plan §10.3 as corrected |
| `prerequisite` | Phase-5 `CURRENT` | D6-14 |
| `run_id` | success counter | D6-15 |
| `statistics` | mean, SD, Type-7 percentile, the full ladder by reference | plan §7, §8 |
| `contingency` | `Selected Px − A` | plan §9 |
| `results_minimum` | Run Stamp + Summary Statistics; everything else deferred | plan §12 |
| `authority_references` | **twelve** declared boundaries | this step |

### What the contract deliberately does NOT own

| Absent | Owner | Enforced by |
|---|---|---|
| the admissible Random Seed **range** | `input_contract.yaml` | `_forbid_seed_range` scans the whole document |
| the Monte Carlo business **minimum** | `input_contract.yaml` | `business_minimum` key refused |
| the selectable **ladder values** | `input_contract.yaml` | resolved by reference; copying refused |
| every oracle comparison **tolerance** | the Step-0 §10 evidence policy | `_forbid_tolerance` — **no exception, null included** |
| the fingerprint **hash mathematics** | `calc_fingerprint.py` | only version numbers appear |
| any executable **algorithm** | Step 2 onward | `test_47`, `test_49` |

---

## 4. The corrected simulation-state authority

Revision 6's predicates were **neither mutually exclusive nor collectively
complete**, and they read the attempt history — contradicting the inherited
Phase-5 orthogonality the plan's own authority matrix records.

**The hole.** Success `A` → invalid edit → `RunSimulation` REFUSED → the user
restores exactly `A`. The fingerprint matches and the last attempt is `REFUSED`,
so CURRENT was false, STALE was false, INVALID was false. **No state at all.**

**The overlap.** A stored success, valid but changed inputs, `FAILED` last
attempt → STALE **and** INVALID.

`docs/phase6_plan.md` §10.3 now carries a clearly marked **POST-ACCEPTANCE
AUTHORITY CORRECTION**. That is the only change to the plan in this round.

**The corrected derivation — ordered, first match wins:**

| # | Condition | Status |
|---|---|---|
| 1 | current prerequisites do not resolve | **INVALID** |
| 2 | prerequisites resolve, no successful snapshot exists | **BLANK** |
| 3 | prerequisites resolve, snapshot exists, fingerprints **equal** | **CURRENT** |
| 4 | prerequisites resolve, snapshot exists, fingerprints **differ** | **STALE** |

Ordering makes the rules **mutually exclusive by construction**; rules 1 and 2
make them **total**. The three labels are unchanged and **no fourth state
exists**.

**The attempt result never participates.** `derive_sim_status` in the loader
takes exactly three parameters — `prerequisites_resolve`,
`successful_snapshot_exists`, `request_fingerprint_matches` — and the attempt
result is not among them. `test_53` pins that signature and scans the source, so
a future branch on attempt history cannot be added without changing something
this test asserts. The loader additionally refuses any rule condition or state
definition containing `attempt`, `refused` or `failed`.

**BLANK is the absence of a comparison, not a fourth label.** A status can only
be CURRENT or STALE *relative to a stored success*. The field was renamed
`never_evaluated_status` → **`no_success_valid_status`**, because a blank status
also occurs after an unsuccessful attempt when no snapshot exists and the current
request is now valid. `status_evaluated_at` may be populated while the status is
blank, which distinguishes *never evaluated* from *evaluated, nothing to compare
against*.

**The required cases, all tested (`test_52`):**

| Case | Status | `last_attempt_result` |
|---|---|---|
| A. success `A` → REFUSED edit → restore `A` | **CURRENT** | `REFUSED` |
| B. success `A` → valid changed request `B` | **STALE** | any |
| C. success `A` → FAILED on `B`, rolled back; viewing `B` | **STALE** | `FAILED` |
| C. …restored to `A` | **CURRENT** | `FAILED` |
| D. prerequisites invalid | **INVALID** | any |
| E. no snapshot, request valid | **BLANK** | any |
| F. no snapshot, request invalid | **INVALID** | any |

`test_51` exhausts all 8 input combinations and asserts every one lands in
exactly one of the four outcomes.

---

## 5. The contract is now a closed world

Independent review demonstrated that `root.future_semantic`,
`rng.future_semantic`, `seeding.future_semantic`, `statistics.future_semantic`
and `result_digest.future_semantic` were all **silently accepted**. A key the
validator does not know is a semantic nobody enforces, and a field that looks
authoritative while governing nothing is worse than no field at all.

**Every mapping now has a declared shape** — **77 of them**, from the root to
each reserved-row record, each run-identity field and each iteration column. Two
things fail:

- an **unknown key** in a known mapping;
- a **mapping at a path the schema does not describe** — so a whole new block
  cannot be added either.

Lists whose membership is locked are closed separately by the section
validators, because closing the mappings alone would still let an invented
record be appended to a list.

`test_87` is systematic: it injects an unknown key into **every one of the 77
declared mappings** and asserts each is rejected. It also asserts the count is at
least 70, so deleting schema entries to make it pass is itself caught.

---

## 6. What a driver contributes — newly contracted

Without this, an engine could satisfy every other section and still implement the
most important arithmetic wrongly.

**Cost Line** — the sample is **unit-cost** uncertainty; Quantity is
deterministic and sits **outside** the distribution:

```
unit_cost       = sample(distribution, Min, MostLikely, Max)
contrib_nominal = unit_cost * Quantity * Knom
contrib_pv      = unit_cost * Quantity * Kpv
```

Quantity is applied **exactly once**. Total-cost uncertainty is **not** sampled.
Probability does not apply.

**Risk** — no Quantity; severity *is* the money amount:

```
occurred = occurrence_uniform < Probability
if occurred: nominal = severity * Knom ; pv = severity * Kpv
else:        nominal = 0               ; pv = 0
```

Probability is **never** folded into `Knom`/`Kpv`, and occurrence and severity
never share a stream.

**PV is an independent contribution**, computed with `Kpv` — never a discount
applied to the nominal total. `pv_derived_from_nominal: false`.

Mutation controls: sampling the total instead of the unit cost · omitting
Quantity · applying it twice · declaring it stochastic · moving it inside the
distribution · folding Probability into the K factors · multiplying a Risk by
Quantity · a non-occurring Risk still contributing · weakening the occurrence
comparison · deriving PV from nominal.

---

## 7. Four more inherited invariants, now written down

**Hot-kernel boundary.** Zero worksheet, `Range`, `ListObject`, `Application`,
`ThisWorkbook`/`ActiveWorkbook` or COM access inside the iteration loop; inputs
resolved once before it; `Knom`, `Kpv`, quantities, probabilities, distribution
parameters and stream initial states all resolved before the loop; no worksheet
inflation, FX or profile recomputation inside it. This is semantic authority: an
engine that touches a worksheet in the loop is wrong even if its numbers are
right.

**Numerical domain (plan §4.6, in full).** Finite correctly ordered negatives
remain legal; supports crossing zero remain legal; **no** positivity rule and
**no** magnitude restriction; a representable, valid result must **not** be
refused merely because a naive intermediate overflows; if no valid `Double`
result exists the refusal is explicit and **names the numerical stage**; silent
non-finite results are forbidden. Disciplines: safe product for contributions,
safe signed sum for accumulation, convex percentile interpolation, safe
subtraction for contingency, scale-safe statistics.

**Driver independence.** `inter_driver_dependence: independent`; no correlation
matrix, no copula, no shared or hidden dependence. The component-stream
architecture makes independence *achievable*; it does not *state* it, and an
unstated invariant is one a later phase can quietly break.

**Publication and command surface.** `_SimData` is the persisted source of truth;
Results derives from it and **never recomputes** a Monte Carlo; publish only
after the statistics complete; commit-last; a REFUSED or FAILED attempt publishes
no partial distribution and the prior publication survives. Automation endpoint
`PCCM_RunSimulation`; **no** Phase-6 user-facing button, MsgBox, UserForm or
ribbon; **no** user cancellation. Reporting read-accessor names are deliberately
**not** invented here — no accepted authority names them yet.

---

## 8. The full confidence ladder — retained by reference

Phase 6 stores the whole ladder a selection could ask for, so Selected Px is a
deterministic lookup rather than a computation over samples.

```
fixed_nonselectable_percentiles:     ["P10"]
include_all_selectable_ladder_values: true
selectable_ladder_locator:            config_tables.confidence_levels
headline_percentiles:                 ["P10", "P50", "P70", "P90"]
```

**The ladder's values are not copied.** `retained_percentiles()` resolves them
from the owner, and today that yields **11 distinct percentiles**:

```
P10  P50 P55 P60 P65 P70 P75 P80 P85 P90 P95
```

`test_63` resolves the owner, asserts every selectable value is retained, and
asserts the values do **not** appear in `sim_contract.yaml` — so a legitimate
future ladder change flows through one authority instead of requiring a
duplicate to be edited in step. P10 remains non-selectable; Selected CL remains
reporting-only.

---

## 9. D6-08 — re-derived from the final layout

`model_version` is an **accepted** Run Stamp field (`workbook.yaml`'s Results
placeholder names it) and it was missing. It is snapshot data — the model version
*at the time of the successful run*, not a live lookup when Results is displayed
— so it belongs with the other version fields at row 19, **not** appended after
the derived group. Row 29 was offered as a spacer to reuse; that would have put
`model_version` after `status_evaluated_at`, in the derived group, which is
semantically wrong. **The layout keeps its regular shape and every later row
shifts by one.**

`H` is therefore **re-derived and changed**, not preserved:

| Row(s) | Purpose |
|---|---|
| 1–5 | shell margin, title, subtitle, rule, spacer |
| 6–7 | run identity section heading, note |
| 8–29 | run identity fields (**22**, was 21) |
| 30 | spacer |
| 31–32 | iteration records section heading, note |
| **33** | **iteration table header** |
| 34 | **first iteration row** |

```
H                            = 33          (was 32)
first_iteration_row          = 34          (was 33)
footer_rows                  = 0
max_iterations_representable = 1048576 − 33
                             = 1048543     (was 1048544)
```

**Boundary vectors.**

| `n` | Last row needed | Outcome |
|---|---|---|
| `1048543` | `34 + 1048543 − 1 = 1048576` | **representable** — the last Excel row exactly |
| `1048544` | `1048577` | **pre-flight technical refusal** |

The refusal remains **technical, never business validation**, and still fires
before sample allocation, stream construction, AUTO seed allocation and any
random draw — so a storage refusal cannot consume an AUTO nonce.

**The run-identity block is now exact authority**: key, row, group and value type,
in order, for all 22 fields. "Contains these required fields" was not enough —
an invented field could be appended on the next free row and the loader accepted
it. `test_89` rejects an appended field, a swapped pair, and `model_version`
moved out of the snapshot group.

---

## 10. Authority bindings that check content, not only resolution

Two bindings were **false**: `_SimData` visibility could be changed from
`veryHidden` to `hidden` and the distribution master list could be changed
outright, and `validate_sim_against` accepted both. A locator that merely reaches
*some* node is not sufficient where this contract also depends on that node's
content.

| Boundary | Locator | Content check |
|---|---|---|
| `_SimData` | `sheets._SimData` | sheet exists **and** `visibility == veryHidden` |
| Results placeholder | `sheets.Results` | the `Run Stamp` and `Summary Statistics` sections still exist |
| model version | `model.model_version` | present and non-blank |
| distribution master | `config_tables.distributions` | **set membership** equals the simulation families; duplicates refused |
| selectable ladder | `config_tables.confidence_levels` | non-empty; no overlap with the fixed headline |
| Random Seed domain | `inputs.random_seed` | whole `between` `1`…`2147483646`, blank allowed, optional |

**Membership, not order.** `lstDistributions` is user-facing presentation order
and has no reason to become dispatch order, so `test_111` asserts a reordering is
**accepted** — a validator that refused it would be enforcing something no
authority says.

Where an upstream loader already guards a mutation at its own boundary — a blank
`model_version` is refused by `load_spec`, a resized locked list by
`load_contract` — that layering is recorded rather than duplicated
(`test_109b`), and the sim-level binding is tested directly against a raw
document.

---

## 11. Zero tolerance semantics

`result_digest.tolerance: null` is **removed**, and so is the special case that
allowed it in `_forbid_tolerance`. Null is not an approximate value, but the
field was still a tolerance semantic living in the file the declared boundary
says holds none. `equality: exact` already states the runtime rule.

Any tolerance-shaped key **anywhere**, null included, is now rejected. Comments
may explain where the tolerance lives; the parsed contract carries none.
`test_45` walks the parsed document and asserts zero tolerance keys and zero
tolerance-mentioning values.

---

## 12. D6-11 — still ungranted, with an activation precondition recorded

The mixed scalar/scoped parser capability is unchanged and **no scoped grant is
added**. `MRG32k3a` and `RunSimulation` remain globally forbidden strings.

**The activation precondition, recorded now because it is a future atomic
dependency:**

> Before the **first real scoped forbidden-construct entry** is committed, the
> **same implementation step** MUST update all actual consumers — the Phase-4
> source scan in `test_phase4_stage_b_source.py::test_15` and the Gate-B / P5-EV
> manifest enforcement — to use `ForbiddenConstruct.forbidden_in(module)`
> semantics.
>
> **No scoped grant may land while any consumer still interprets that construct
> as globally forbidden or silently drops `allowed_in`.**

Today both consumers read the flattened global string list, which is correct
precisely because nothing is scoped. `bootstrap` is **not** changed in this
round. `test_67` asserts no scoped rule exists and that this precondition is
recorded.

---

## 13. The Stage-A claim, corrected

The previous record said the Stage-A workbook is byte-identical with or without
Step 1. **That was too broad.** Precisely:

- the **simulation contract** emits no Phase-6 workbook block and no VBA — it is
  a load-time gate;
- the Stage-A **verification count remains 351**;
- the Step-1 **Random Seed validation change legitimately changes the Stage-A
  workbook**, by adding a data-validation rule to `C21`;
- **no claim of byte-identical whole-workbook output is made.**

---

## 14. The exact Cheng formulation — unchanged

Encoded field by field: both orientations (BB min-first, BC **max**-first), every
setup expression, every literal **as a literal**, the exact expression order, the
acceptance operators, and the locked logit form `log(u1 / (1.0 - u1))`.

`1.3862944` is those digits and is **not** evaluated as `log(4)`; `0.0138889`,
`0.0416667` and `0.777778` are not `1/72`, `3/72` and `7/9`. `test_13` checks
this **semantically** — `log(4)` may legitimately appear as documentation of what
the literal approximates, so the test asserts it never appears in an expression
the implementation would follow.

Drift is detectable: the contract carries the SHA-256 of the three retained
Step-0 function bodies (`d5ca71b8…eed90`) and both jump-matrix hashes, compared
against the evidence directly.

---

## 15. Version semantics — unchanged

`RNG_VERSION = 1`, `SIM_METHOD_VERSION = 1`, `FP_VERSION = 1` inherited and
untouched. The `PCCM-RD` version field **is** `SIM_METHOD_VERSION` — no third
version, and the field is not semantically ownerless. Eleven changes are
classified; `test_65` rejects removing any, `test_66` rejects moving one to the
wrong version.

**`FP_VERSION` is deliberately NOT added to the simulation snapshot.** No
accepted authority requires it there, and inventing it in a correction round
would be exactly the kind of unearned addition this process exists to prevent.

---

## 16. Test inventory

| Suite | Tests | Kind |
|---|---|---|
| `test_phase6_sim_contract.py` | **67** | positive, authority conformance, evidence binding, state truth table, scope discipline |
| `test_phase6_sim_contract_validation.py` | **114** | **mutation controls** |
| `test_phase2_inputs.py::test_19d` | **1** | the accepted seed domain, in contract and on the sheet |
| **Total new** | **182** | |

New in this round: the full state truth table and its exhaustion · the attempt
axis proven unable to reach the derivation · systematic unknown-key injection
into all 77 closed mappings · a mapping at an undeclared path · an appended,
reordered or regrouped run-identity field · a missing `model_version` ·
`_SimData` visibility drift · a broken Results placeholder binding · distribution
master disagreement in three directions, with reordering explicitly accepted · an
empty ladder · dropping the full-ladder rule · a null tolerance · ten contribution
mutations · eleven kernel mutations · four dependence mutations · five
numerical-domain mutations · five unsafe-arithmetic mutations · cancellation ·
six publication mutations.

Every mutation control from the first Step-1 round is retained; none was
weakened or deleted.

---

## 17. One existing test was narrowed, and why that is not a weakening

*(Accepted in review; recorded here unchanged.)*

`test_phase2_inputs.py::test_19c` asserted that Random Seed has **no**
validation. Its own docstring said "no business limits **yet**", and the contract
note said the domain "is fixed when the RNG is implemented". Both were explicit
deferrals, and D6-19 and D6-20 discharged them.

`random_seed` was removed from `test_19c`'s list and `test_19d` added, asserting
the accepted domain in the contract **and** as applied to the sheet. **Net
protection increases**: the seed moved from *unvalidated* to *validated against
an accepted authority*, and the new test also pins the blank-is-AUTO affordance,
which nothing previously checked.

---

## 18. Static verification

| Check | Result |
|---|---|
| Full Python suite (`python3 -m pytest pccm/tests -q`) | **1934 passed, 0 failed** (174.24 s) |
| Baseline before Step 1 | 1752 passed |
| New Step-1 tests | **+182** (67 + 114 + 1) |
| Mutation controls | **114** |
| Stage-A verifier / build | **351 passed, 0 failed**; "Stage A build complete." |

### Baseline preservation

Byte-identical to the first Step-1 commit `1be66d3`:

| Path | `git diff 1be66d3` |
|---|---|
| `pccm/src` | **EMPTY** |
| `pccm/bootstrap` | **EMPTY** |
| `pccm/spec/workbook.yaml` | **EMPTY** |
| `pccm/spec/calc_contract.yaml` | **EMPTY** |
| `pccm/spec/driver_contract.yaml` | **EMPTY** |
| `pccm/spec/structure_contract.yaml` | **EMPTY** |
| `pccm/spec/input_contract.yaml` | **EMPTY** — the accepted seed change was not touched again |
| `pccm/evidence` | **EMPTY** |
| `pccm/tests/test_phase2_inputs.py` | **EMPTY** — `test_19c`/`test_19d` not reverted |

Phase-5 fingerprint vectors, recomputed live at this commit:

```
fingerprint("PCCM-FP")     6551C6F365DA7F3F
fingerprint_probe(A|B)     42E49DC715F06970
fingerprint_probe(AB|)     7558FD9248656EAD
canonical_number(1/3)      3.3333333333333331E-01
FP_BASE / FP_MOD_1 / FP_MOD_2   131 / 2147483647 / 2147483629
```

`test_phase5_fingerprint.py` — **52 passed**.

---

## 19. Step-1 acceptance gate

| Requirement | Status |
|---|---|
| state derivation total, exclusive, attempt-orthogonal | yes — §4 |
| no fourth simulation state | yes — blank is an absence, `test_93` |
| no-success/valid represented by blank status | yes — `no_success_valid_status: null` |
| unknown semantic keys fail everywhere | yes — 77 mappings, `test_87` |
| run-identity layout exact, not extensible | yes — `test_89` |
| Cost Line and Risk contribution arithmetic contracted | yes — §6 |
| zero worksheet/COM hot-loop boundary contracted | yes — §7 |
| numerical-domain guarantees contracted | yes — §7 |
| driver independence contracted | yes — §7 |
| full selectable ladder retained by reference | yes — 11 today, §8 |
| `model_version` persisted in the Run Stamp | yes — §9 |
| D6-08 re-derived from the final layout | yes — `H = 33`, max `1048543` |
| borrowed-content bindings cannot silently drift | yes — §10 |
| **zero** tolerance field in `sim_contract` | yes — §11 |
| publication / cancellation / command surface encoded | yes — §7 |
| D6-11 ungranted, activation precondition recorded | yes — §12 |
| Phase-5 behaviour green | yes — §18 |
| no Step-2 implementation, no VBA | yes — `test_47`, `test_50` |
| no Windows/Excel runtime | yes — Linux only |

---

## 20. What Step 1 does NOT do

- **It implements nothing.** No function advances an MRG state, generates a
  uniform, executes a jump, samples a variate, runs a Bernoulli trial or executes
  an iteration. `test_47` scans the loader for exactly those. `derive_sim_status`
  is state *semantics*, not simulation: it draws nothing and reads no workbook.
- **It does not read evidence at run time.** Tests read the retained Step-0
  package for conformance; `test_48` asserts no production module references it.
- **It does not evaluate the contract's formulas.** They are text. `test_49`
  asserts the loader never calls `eval`, `exec` or `compile`.
- **It does not touch the workbook structure.** No sheet, table, cell or emission
  changed. The Random Seed data validation is the one Stage-A surface change, and
  it landed in the accepted first Step-1 commit — see §13.
- **It does not prove any of this works in Excel.** That is Gate-A and Gate-B.

---

**STEP 1 — ACCEPTANCE REQUESTED (CONTRACT COMPLETE)**
