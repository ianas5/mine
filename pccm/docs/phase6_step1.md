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
| Step-1 contract completion | `35c2467c1f0852fd6cbe5285600c96baeedca2de` |
| Step-1 authority hardening | this commit — reported by the delivery message and `PROVENANCE.txt` |

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

### What the hardening round corrected

Independent review of `35c2467` accepted the contract architecture in substance
and returned NOT YET ACCEPTED on **enforcement**: the contract claimed to lock
authority the loader did not actually check. One inherited-authority
contradiction and eleven enforcement gaps are closed.

| # | Defect | Fix |
|---|---|---|
| 1 | **Uniform degeneracy read the ignored Most Likely** | family-specific detection — §21 |
| 2 | `RNG_VERSION = 2` was accepted | pinned to `1` — §22 |
| 3 | `recurrence.advance = "banana"` was accepted | pinned — §22 |
| 4 | BC expressions and **both** returns were free text | pinned; evidence paths pinned — §22 |
| 5 | the whole digest grammar was free text | pinned token by token — §22 |
| 6 | conditioning scales and triangular boundaries were free text | pinned — §22 |
| 7 | run-identity initials, enum owners and block columns were free | complete record pinned — §23 |
| 8 | iteration column letters, headers and types were free | pinned — §23 |
| 9 | **55 of 647 key deletions were accepted** | required-key discipline + sweep — §24 |
| 10 | wrong *values* were unchecked | semantic-leaf sweep — §24 |
| 11 | a zeroed jump hash was accepted | hashes verified, not decorative — §22 |
| 14 | D6-08 | **re-derived unchanged** at `H = 33` — §9 |

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
| `spec/sim_contract.yaml` | **NEW** — the sixth authority, 1091 lines, 30 sections |
| `spec/input_contract.yaml` | seed admissibility resolved; deferred note discharged |
| `builder/pccm_builder/sim_loader.py` | **NEW** — loader, validator, cross-validator |
| `builder/pccm_builder/structure_loader.py` | D6-11 mixed scalar-or-scoped schema |
| `builder/pccm_builder/__init__.py` | exports |
| `builder/build_stage_a.py` | loads and cross-validates the sim contract; **emits nothing** |
| `tests/test_phase6_sim_contract.py` | **NEW** — 71 positive / conformance tests |
| `tests/test_phase6_sim_contract_validation.py` | **NEW** — 131 mutation controls |
| `tests/test_phase2_inputs.py` | `test_19c` narrowed, `test_19d` added — see §9 |
| `docs/phase6_step1.md` | **NEW** — this record |
| `docs/phase6_plan.md` | two **post-acceptance authority corrections**: §10.3 state derivation (§4) and §4.0/§4.1 Uniform degeneracy (§21) |

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

**Re-derived again after the hardening round and unchanged.** The corrections in
§21–§24 pin values and column letters; none of them adds or removes a row, so the
tiling still covers rows 1–33 and `1048576 − 33 = 1048543` still follows from it.
The loader recomputes it every load, so this is a derivation, not an assertion.

**The run-identity block is now exact authority**: key, row, group, label, value
type, enum owner and initial, in order, for all 22 fields. "Contains these required fields" was not enough —
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
| `test_phase6_sim_contract.py` | **71** | positive, authority conformance, evidence binding, state truth table, Uniform degeneracy cases, scope discipline |
| `test_phase6_sim_contract_validation.py` | **131** | **mutation controls**, including two systematic sweeps |
| `test_phase2_inputs.py::test_19d` | **1** | the accepted seed domain, in contract and on the sheet |
| **Total new** | **203** | |

**Sweep sizes:**

| Sweep | Shapes exercised | Instances covered |
|---|---|---|
| unknown-key injection (`test_87`) | **78** mappings | 78 |
| missing-key deletion (`test_126`) | **444** (mapping, key) shapes | 653 |
| semantic-value mutation (`test_128`) | **367** leaf paths | 714 |

New in this round: six Uniform-degeneracy controls · both version bumps · the
state transition · six Cheng BC/BB expression and return mutations · three
evidence-binding mutations · two zeroed jump hashes · five digest-grammar
mutations · six conditioning and boundary mutations · eleven run-identity initial
and enum mutations · three block-column mutations · five iteration-column
mutations · a reserved-row purpose · two definition mutations · and the three
sweeps above.

Every mutation control from both earlier Step-1 rounds is retained; none was
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
| Full Python suite (`python3 -m pytest pccm/tests -q`) | **1955 passed, 0 failed** (343.56 s) |
| Baseline before Step 1 | 1752 passed |
| New Step-1 tests | **+203** (71 + 131 + 1) |
| Mutation controls | **131** |
| Unknown-key sweep | **78** |
| Missing-key deletion sweep | **444** shapes / 653 instances |
| Semantic-value mutation sweep | **367** leaves / 714 instances |
| Stage-A verifier / build | **351 passed, 0 failed**; "Stage A build complete." |

### Baseline preservation

Byte-identical to the contract-completion commit `35c2467`:

| Path | `git diff 35c2467` |
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
| `pccm/builder/pccm_builder/structure_loader.py` | **EMPTY** |
| `pccm/builder/build_stage_a.py` | **EMPTY** |

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
| Uniform degeneracy no longer reads the ignored Most Likely | yes — §21 |
| ignored Uniform ML cannot change RNG consumption | yes — contracted and mutation-controlled |
| `RNG_VERSION` and `SIM_METHOD_VERSION` exactly `1` | yes — §22 |
| complete MRG state transition pinned | yes — §22 |
| complete BB/BC formulation **and returns** pinned | yes — §22 |
| Cheng evidence bindings enforceable | yes — pinned **and** followed by `test_48b` |
| complete result-digest grammar pinned | yes — §22 |
| conditioning and boundary semantics pinned | yes — §22 |
| run-identity initials, enums and columns exact | yes — §23 |
| iteration columns, types and headers exact | yes — §23 |
| missing required keys fail systematically | yes — 444 shapes, §24 |
| wrong semantic values fail systematically | yes — 367 leaves, §24 |
| jump hashes verified, not decorative | yes — option A, §22 |
| D6-08 re-derived | yes — **unchanged at `H = 33`**, §9 |
| state derivation total, exclusive, attempt-orthogonal | yes — §4 |
| no fourth simulation state | yes — `test_93` |
| unknown semantic keys fail everywhere | yes — 78 mappings, `test_87` |
| contribution, kernel, numerical domain, dependence contracted | yes — §6, §7 |
| full selectable ladder retained by reference | yes — 11 today, §8 |
| `model_version` persisted in the Run Stamp | yes — §9 |
| borrowed-content bindings cannot silently drift | yes — §10 |
| **zero** tolerance field in `sim_contract` | yes — §11 |
| D6-11 ungranted, activation precondition recorded | yes — §12 |
| all previous Step-1 protections remain | yes — none weakened or deleted |
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

## 21. Uniform degeneracy no longer reads the ignored Most Likely

This was a real **inherited-authority contradiction**, not a validator gap.

Accepted Phase-5 **D1**: Uniform's Most Likely is ignored numerically and
excluded from the calculation fingerprint. Accepted plan **§4.1**: a degenerate
Uniform `a = b` returns `a` and consumes nothing. But **§4.0 used one predicate
for all three families** — `a = m = b`.

**The consequence.** A legal Uniform with `Min = Max` and a populated, unrelated
Most Likely was **not** degenerate under that predicate. It would enter the
sampler and consume a uniform — so an input the model explicitly ignores changed
**RNG consumption**, the stream position, and every subsequent draw on that
component.

**Detection is now family-specific:**

| Family | Degenerate when |
|---|---|
| **Uniform** | **`a == b`** — Most Likely is not read |
| Triangular | `a == m == b` |
| Beta-PERT | `a == m == b` |

For Triangular and Beta-PERT the accepted ordering `a ≤ m ≤ b` already makes
`a = b` imply `m = a`, so the three-way condition states the semantic rather than
adding a restriction. The outcome for every degenerate family is unchanged:
return `a`, enter no sampler, consume **zero** uniforms, stream unchanged.

The contract also states plainly that Most Likely affects neither degeneracy nor
Uniform consumption. `phase6_plan.md` §4.0 and §4.1 carry a marked
**POST-ACCEPTANCE AUTHORITY CORRECTION** explaining why the common predicate was
wrong. Nothing else in the plan was reopened.

**Conformance cases** (`test_69`), each asserted against the contract:

| # | Case | Expected |
|---|---|---|
| A | Uniform `a = b`, Most Likely blank | degenerate, **0** uniforms |
| B | Uniform `a = b`, Most Likely populated and unrelated | degenerate, **0** uniforms |
| C | same `Min`/`Max`, two different ignored Most Likely values | identical semantics |
| D | Uniform `a < b` | non-degenerate, **1** uniform |

---

## 22. Values that are now pinned, not merely shaped

Each of these was demonstrated to be accepted before this round.

| Was accepted | Now |
|---|---|
| `RNG_VERSION = 2`, `SIM_METHOD_VERSION = 2` | pinned to **`1`** — a bump is an authority change, not a value a validator waves through because 2 is also a positive integer |
| `rng.recurrence.advance = "banana"` | the state shift is pinned exactly. A recurrence with the right `p1`/`p2` and the wrong word order is a plausible stream that is not MRG32k3a |
| `cheng.bc.per_driver = ["banana"]`, a BC expression replaced, `cheng.bc.return = "banana"`, `cheng.bb.return = "banana"` | BB **and** BC per-driver, per-attempt and return all pinned. BB and BC orient **oppositely**, so a free BC return is exactly how a mirrored distribution ships while every other check passes |
| the Cheng evidence paths pointed anywhere | both paths pinned, and `test_48b` **follows what the contract declares** and verifies the artefacts and hashes behind them |
| all three `result_digest.grammar` productions | pinned token by token, including nominal/PV order and the version source |
| `conditioning_scale = "banana"` on both families | pinned to `s = max(abs(a), abs(m), abs(b))` |
| triangular `m_equals_a` / `m_equals_b` | pinned — which branch a boundary shape takes is sampling semantics, not documentation |
| `a1_p127_sha256 = "0000…"` with the matrix still correct | both hashes verified against the accepted Step-0 values, and `test_48b` re-hashes the declared matrices |
| `stream_assignment.index_rule`, the state definitions, the reserved-row purposes | pinned — the tiling's purposes are the audit trail for `H` |

**Option A was taken on the jump hashes**, not removal: the hashes are now part
of the enforced authority. Authoritative-looking metadata a validator ignores is
worse than none, because it invites a reader to trust a binding that does not
exist.

---

## 23. `_SimData` is now exact in every field

The block was called "EXACT authority" while only key, row, group and type were
checked. Initials could be seeded and enum owners swapped — so a materialiser
could have written a **partial successful snapshot** into a workbook that had
never run.

**The locked record is now** `(key, row, group, label, value_type, enum,
initial)` for all 22 fields, plus the block columns `B` / `D` / `F`.

**Cross-semantic agreement is enforced**, because the same fact appears in two
sections — one is the rule, the other the cell that carries it:

| Field | Initial | Agrees with |
|---|---|---|
| `next_auto_nonce` | `0` | `seeding.nonce_lifecycle.initial` |
| `last_run_id` | `0` | `run_id.initial` |
| `last_attempt_result` | `NONE` | — |
| `simulation_status` | **blank** | a never-run workbook presents no derived status |
| `run_id`, `request_fingerprint`, `result_digest`, `effective_seed`, `last_successful_stamp`, `model_version` | **blank** | written only by a successful commit |

**Enum ownership is exact and conditional**: `seed_mode → seed_mode`,
`last_attempt_result → attempt_result`, `last_attempt_seed_mode → seed_mode`,
`simulation_status → sim_state`. The `enum` key is **required** when
`value_type == "enum"` and **refused** otherwise — both directions, because a
label set that governs nothing is a semantic nobody reads.

**Iteration columns are exact too**: `B` `Iteration` integer, `C` `Total Nominal`
double, `D` `Total PV` double, in that order. This is column-layout hardening and
**does not change `H`**.

---

## 24. Two systematic sweeps

**Unknown keys** were already closed. Two further sweeps close the other two ways
authority can be lost.

### Missing keys — the deletion sweep

An independent sweep deleted 647 keys and **55 were accepted**: a semantic could
be *removed* from the authority document and the validator still called it valid.

Every mapping now carries required keys as well as allowed keys, and
`test_126` deletes **every one of the 444 distinct (mapping, key) shapes** —
covering all 653 instances, since deleting a key removes it from every mapping at
that path — and requires rejection.

**`INTENTIONALLY_OPTIONAL` is deliberately empty.** Every key is required,
*including the ones whose canonical value is `null`*: `positivity_rule: null` is
the authority **saying** there is no positivity rule, and silent absence does not
say that. The test fails if that list grows.

**One conditional key**, with its rule written down and both directions enforced:
`sim_data.run_identity.fields[].enum`. `test_127` proves it is required where it
applies and refused where it does not — a conditional key is not an optional one.

### Wrong values — the semantic-leaf sweep

Shape closure said nothing about content. `test_128` mutates **each of the 367
distinct settled leaves** (714 instances) to a **type-compatible wrong value** —
`+1` for integers, the next float up, `not` for booleans, `"banana"` for strings,
a value where the canonical content is `null` — and requires rejection.

**`FLEXIBLE_LEAVES` holds exactly one entry**, with a reason:
`dependence.authority`, a prose citation whose enforced semantics are the four
booleans beside it. `test_129` proves that exemption is real by changing it and
loading successfully, so the allow-list is not decorative, and `test_128` fails
if it grows.

The goal, stated plainly: **no settled semantic leaf in `sim_contract.yaml` can
change while the loader still reports the accepted Step-1 authority as valid.**

---

**STEP 1 — ACCEPTANCE REQUESTED (AUTHORITY HARDENED)**
