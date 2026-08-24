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
| Step-1 contract | this commit — reported by the delivery message and `PROVENANCE.txt` |

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
| `spec/sim_contract.yaml` | **NEW** — the sixth authority, 869 lines |
| `spec/input_contract.yaml` | seed admissibility resolved; deferred note discharged |
| `builder/pccm_builder/sim_loader.py` | **NEW** — loader, validator, cross-validator |
| `builder/pccm_builder/structure_loader.py` | D6-11 mixed scalar-or-scoped schema |
| `builder/pccm_builder/__init__.py` | exports |
| `builder/build_stage_a.py` | loads and cross-validates the sim contract; **emits nothing** |
| `tests/test_phase6_sim_contract.py` | **NEW** — 50 positive / conformance tests |
| `tests/test_phase6_sim_contract_validation.py` | **NEW** — 86 mutation controls |
| `tests/test_phase2_inputs.py` | `test_19c` narrowed, `test_19d` added — see §9 |
| `docs/phase6_step1.md` | **NEW** — this record |

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
| `request_fingerprint` | `HEADER·COST·RISK·SIM`, prefix semantics, SIM fields | plan §11 |
| `result_digest` | `PCCM-RD` framing | D6-17 |
| `iterations` | the **technical** ceiling only | D6-08 |
| `sim_data` | the future machine layout, and therefore `H` | D6-08, plan §11.1 |
| `label_sets`, `sim_state` | exactly three states | plan §7 |
| `prerequisite` | Phase-5 `CURRENT` | D6-14 |
| `run_id` | success counter | D6-15 |
| `statistics` | mean, SD, Type-7 percentile | plan §8 |
| `contingency` | `Selected Px − A` | plan §9 |
| `results_minimum` | Run Stamp + Summary Statistics; everything else deferred | plan §12 |
| `authority_references` | ten declared boundaries | this step |

### What the contract deliberately does NOT own

| Absent | Owner | Enforced by |
|---|---|---|
| the admissible Random Seed **range** | `input_contract.yaml` | `_forbid_seed_range` scans the whole document |
| the Monte Carlo business **minimum** | `input_contract.yaml` | `business_minimum` key refused |
| every oracle comparison **tolerance** | the Step-0 §10 evidence policy | `_forbid_tolerance` scans keys and values |
| the fingerprint **hash mathematics** | `calc_fingerprint.py` | only version numbers appear |
| any executable **algorithm** | Step 2 onward | `test_47`, `test_49` |

---

## 4. D6-08 — CLOSED

The `_SimData` geometry is declared row by row, and **every row of the sheet is
accounted for**. The reserved ranges must tile rows `1 … H` with no gap and no
overlap — a gap leaves a row unaccounted for and an overlap counts one twice, and
either makes `H` unauditable — so the loader refuses both.

| Row(s) | Purpose |
|---|---|
| 1 | shell top margin |
| 2 | shell title |
| 3 | shell subtitle |
| 4 | shell rule |
| 5 | spacer |
| 6 | run identity section heading |
| 7 | run identity section note |
| 8 – 28 | run identity fields (21) |
| 29 | spacer |
| 30 | iteration records section heading |
| 31 | iteration records section note |
| **32** | **iteration table header** |
| 33 | **first iteration row** |

```
H                            = 32          (rows an iteration record cannot occupy)
first_iteration_row          = 33
footer_rows                  = 0           (nothing below the records)
max_iterations_representable = 1048576 − 32
                             = 1048544
```

**`H` is derived, not declared.** The loader recomputes it from the reserved-row
tiling and refuses a `reserved_rows_h` or `max_iterations_representable` that
disagrees. `test_57` proves it: adding one reserved row without moving the ceiling
is rejected, which is exactly the defect a free literal would hide.

**Boundary vectors.**

| `n` | Last row needed | Outcome |
|---|---|---|
| `1048544` | `33 + 1048544 − 1 = 1048576` | **representable** — the last Excel row exactly |
| `1048545` | `1048577` | **pre-flight technical refusal** |

The refusal is **technical, never business validation**, and it fires **before**
sample allocation, stream construction, AUTO seed allocation and any random draw.
That ordering is contract text rather than an implementation detail because it is
what guarantees a storage refusal **does not consume an AUTO nonce**.

---

## 5. The exact Cheng formulation

Encoded field by field: both orientations (BB min-first, BC **max**-first), every
setup expression, every literal **as a literal**, the exact expression order, the
acceptance operators, and the locked logit form `log(u1 / (1.0 - u1))`.

`1.3862944` is those digits and is **not** evaluated as `log(4)`; `0.0138889`,
`0.0416667` and `0.777778` are not `1/72`, `3/72` and `7/9`. `test_13` checks this
**semantically** — `log(4)` may legitimately appear as documentation of what the
literal approximates, so the test asserts it never appears in an expression the
implementation would follow.

**Drift is detectable.** The contract carries the SHA-256 of the three retained
Step-0 function bodies (`d5ca71b8…eed90`) and both jump-matrix hashes, and
`test_15`, `test_05` and `test_06` compare them against the evidence directly.

`vectors/cheng_vectors.json` is declared a **conformance authority**, with
`runtime_lookup_table: false`; `test_19` refuses the relabelling.

---

## 6. Version semantics

`RNG_VERSION = 1`, `SIM_METHOD_VERSION = 1`, `FP_VERSION = 1` inherited and
untouched. The `PCCM-RD` version field **is** `SIM_METHOD_VERSION` — no third
version, and the field is not semantically ownerless.

Eleven changes are classified, five to `RNG_VERSION` and six to
`SIM_METHOD_VERSION`. `test_65` mutation-checks the table: removing any entry
leaves a rule that can change retained output owned by nothing, and is rejected.
`test_66` rejects moving a change to the wrong version.

---

## 7. D6-11 — capability without premature authorisation

The schema now understands two shapes:

```yaml
- "Rnd("                          # globally forbidden - unchanged
- construct: "MRG32k3a"           # forbidden EXCEPT in the declared owners
  allowed_in: ["modSimRng"]
```

**Nothing was granted.** `MRG32k3a` and `RunSimulation` remain **globally
forbidden strings**, because `modSimRng` and `modSimReport` do not exist. The
loader refuses a scoped owner that is not a declared module (`test_79`), so a
construct cannot be granted to a module before that module arrives — the grant
and the module land together or not at all.

Refused: unknown owner · empty `allowed_in` · duplicate owner · wildcard (`*`,
`**`, `all`, `any`, `all_modules`, `*.bas`) · malformed entry · a mapping
carrying only one of the two keys · a duplicate construct.

`test_78` proves the scoped shape works, using a **synthetic** construct against
an already-declared module — capability demonstrated, nothing pre-authorised.
`Rnd(`, `Randomize`, `NPV` and `Percentile` stay globally forbidden and `test_85`
asserts none of them acquired an exception.

Existing consumers are unaffected: `StructureContract.forbidden_constructs`
remains a flat list of strings, so the Stage-B manifest and the Phase-4 static
scan read exactly what they read before.

---

## 8. Test inventory

| Suite | Tests | Kind |
|---|---|---|
| `test_phase6_sim_contract.py` | **50** | positive, authority conformance, evidence binding, scope discipline |
| `test_phase6_sim_contract_validation.py` | **86** | **mutation controls** |
| `test_phase2_inputs.py::test_19d` | **1** | the accepted seed domain, in contract and on the sheet |
| **Total new** | **137** | |

Mutation controls cover: RNG constants (including a one-ULP `norm` perturbation)
· state order · uniform combination and endpoints · every element of both jump
matrices (18 separate mutations) · jump exponent, substreams and `H` · `log(4)`
substituted for the literal · the `log1p` formulation · BB/BC boundary and
equality owner · inverted orientations · reordered expression lists · BC literal
fractions · seed range duplicated into the sim contract · input-contract seed
range widened, lowered, removed or reduced to a minimum · blank forbidden ·
mixer introduced · nonce lifecycle reordered, truncated, or made free after
allocation · exhaustion moved or wrapped · stepped multiplication as authority ·
cross-workbook uniqueness claimed · physical row order · numeric ID sorting ·
component kinds · the stream-shift consequence contradicted · conditional
severity · the withdrawn advancement wording · degenerate consuming a uniform ·
non-strict occurrence · digest tag, field order, index origin, sorting and
version ownership · approximate digest equality · `SIM` before `RISK` · Selected
CL or run-scoped identities in the request · a fourth sim state · a free ceiling
literal · a layout change that does not move the ceiling · reserved-row gap and
overlap · a footer · a business maximum · a refusal that consumes a nonce ·
tolerances in four places · authority references deleted, added, redirected and
unresolvable · percentile and SD methods · forbidden contingency baselines · a
fourth distribution family · every D6-11 malformed shape.

`test_00` asserts the unmutated contract loads — without it every rejection
below would prove nothing.

---

## 9. One existing test was narrowed, and why that is not a weakening

`test_phase2_inputs.py::test_19c` asserted that Random Seed has **no**
validation. Its own docstring said "no business limits **yet**", and the contract
note said "the admissible domain is fixed when the RNG is implemented". Both were
explicit deferrals. D6-19 and D6-20 discharged them.

`random_seed` was therefore removed from `test_19c`'s list and `test_19d` added,
asserting the accepted domain in the contract **and** as applied to the sheet —
`whole` · `between` · `1` · `2147483646` · `allow_blank` true · `errorStyle` stop.

**Net protection increases**: the seed moved from *unvalidated* to *validated
against an accepted authority*, and the new test also pins the blank-is-AUTO
affordance, which nothing previously checked. No other test was changed, none was
deleted, none was skipped, and no assertion was loosened.

---

## 10. Static verification

| Check | Result |
|---|---|
| Full Python suite (`python3 -m pytest pccm/tests -q`) | **1889 passed, 0 failed** (238.23 s) |
| Baseline before Step 1 | 1752 passed |
| New Step-1 tests | **+137** (50 + 86 + 1) |
| Mutation controls | **86** |
| Stage-A verifier / build | **351 passed, 0 failed**; "Stage A build complete." |

Stage-A count is **unchanged**: the simulation contract is a load-time gate, not
a workbook verification. Phase 6 emits nothing in Stage A, so the built workbook
is byte-identical with or without this step.

### Baseline preservation

Byte-identical to the accepted Step-0 head `49effe03`:

| Path | `git diff 49effe03` |
|---|---|
| `pccm/src` | **EMPTY** |
| `pccm/bootstrap` | **EMPTY** |
| `pccm/spec/workbook.yaml` | **EMPTY** |
| `pccm/spec/structure_contract.yaml` | **EMPTY** |
| `pccm/spec/calc_contract.yaml` | **EMPTY** |
| `pccm/evidence` | **EMPTY** |
| `pccm/docs/phase6_plan.md` | **EMPTY** |
| `pccm/docs/phase6_step0.md` | **EMPTY** |
| `builder/pccm_builder/calc_fingerprint.py` | **EMPTY** |
| `tests/test_phase5_fingerprint.py` | **EMPTY** |
| `tests/oracle`, `tests/fixtures` | **EMPTY** |

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

## 11. Step-1 acceptance gate

| Requirement | Status |
|---|---|
| `sim_contract.yaml` exists | yes |
| D6-08 CLOSED with exact `H` and maximum | yes — `H = 32`, max `1048544` |
| `input_contract` owns the seed range | yes |
| `sim_contract` carries no duplicated seed range | yes — scanned, not assumed |
| exact Cheng formulation encoded | yes — §5 |
| exact jump matrices encoded | yes — element-checked against evidence |
| exact version ownership encoded | yes — §6 |
| D6-18b including degenerate zero-consumption | yes |
| request-fingerprint `SIM` extension | yes |
| `result_digest` framing | yes |
| statistics / contingency / sim-state semantics | yes |
| comparison tolerance absent | yes — scanned |
| D6-11 capability without premature exceptions | yes — §7 |
| every malformed-shape test passes | yes — 86 controls |
| Phase-5 regressions green | yes — 1889 total, none altered |
| no Phase-6 implementation exists | yes — `test_47`, `test_50` |
| no Windows/Excel runtime ran | yes — Linux only |

---

## 12. What Step 1 does NOT do

- **It implements nothing.** No function advances an MRG state, generates a
  uniform, executes a jump, samples a variate, runs a Bernoulli trial or executes
  an iteration. `test_47` scans the loader for exactly those.
- **It does not read evidence at run time.** Tests read the retained Step-0
  package for conformance; `test_48` asserts no production module references it.
- **It does not evaluate the contract's formulas.** They are text. `test_49`
  asserts the loader never calls `eval`, `exec` or `compile`.
- **It does not touch the workbook.** No sheet, table, cell or emission changed.
- **It does not prove any of this works in Excel.** That is Gate-A and Gate-B.

---

**STEP 1 — ACCEPTANCE REQUESTED**
