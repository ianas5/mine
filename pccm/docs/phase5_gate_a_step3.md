# Phase 5 — Gate A — Step 3: Stage-A calculation workspace emission

**Status: READY FOR INDEPENDENT REVIEW.**

Step 2 is accepted and closed. This step makes `spec/calc_contract.yaml` a real
build input, puts the physical `_Calc` calculation workspace into the generated
workbook, and emits the two generated Phase-5 artifacts a later VBA implementation
will consume. It is Linux-only.

**The purpose of this step is REPRESENTATION, not calculation.** The accepted
numerical oracle remains the semantic authority; `calc_contract.yaml` remains the
physical workbook-representation authority. The workbook ships with no calculated
result of any kind.

---

## 0. The statements this step must make explicitly

> **NO EXECUTABLE PHASE-5 CALCULATION VBA WAS IMPLEMENTED.**
> **NO PHASE-5 NUMERICAL VBA WAS IMPLEMENTED.**
> **NO WINDOWS HARNESS WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **NO CALCULATION BUTTON WAS ADDED.**
> **NO RNG / MONTE CARLO IMPLEMENTATION WAS ADDED.**
> **STEP 4 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

`modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint`, `modCalcResolve`,
`modCalcCheck`, `modCalcReport` and `PCCM_Calculate` do not exist and are not
declared in the Stage-B manifest. A test asserts that.

---

## 1. Scope implemented

| # | Delivered | File |
|---|---|---|
| A | `calc_contract.yaml` as a required, cross-validated build input | `builder/build_stage_a.py` |
| B | the physical `_Calc` Phase-5 workspace renderer | `builder/pccm_builder/calc_render.py` |
| C | the generated VBA constants module and the acceptance corpus | `builder/pccm_builder/calc_emit.py` |
| D | the case fixtures the corpus is built from | `builder/pccm_builder/calc_cases.py` |
| E | extended post-build artifact verification | `builder/pccm_builder/verify.py` |
| F | Step-3 tests | `tests/test_phase5_stage_a.py` |
| G | this document | `docs/phase5_gate_a_step3.md` |

### Files changed

| File | Change |
|---|---|
| `builder/build_stage_a.py` | `--calc-contract`, load + `validate_calc_against` before any emission, pass to build and verify, emit the two Phase-5 artifacts |
| `builder/pccm_builder/calc_render.py` | **new** — the workspace renderer |
| `builder/pccm_builder/calc_emit.py` | **new** — `modCalcContract.bas` and `phase5_cases.json` |
| `builder/pccm_builder/calc_cases.py` | **new** — the case matrix and the C1/C2 regression corpus, as data |
| `builder/pccm_builder/verify.py` | `calc` parameter; `_verify_calc`; the table-inventory gate extended by five |
| `builder/pccm_builder/workbook_builder.py` | `calc` parameter; renders the workspace onto `_Calc`; `BUILDER_VERSION` 0.4.0 → 0.5.0 |
| `builder/pccm_builder/__init__.py` | exports `emit_calc_artifacts` and `validate_calc_against` |
| `spec/workbook.yaml` | `model_version` 0.4.0 → 0.5.0, `build_phase` → Phase 5 |
| `VERSION` | 0.4.0 → 0.5.0 |
| `tests/test_phase4_structure.py` | the version-convention test's literal moved with it; see §11 |
| `tests/test_phase5_stage_a.py` | **new** — 53 tests |

**Unchanged, and verified unchanged:** `spec/calc_contract.yaml`,
`spec/input_contract.yaml`, `spec/driver_contract.yaml`,
`spec/structure_contract.yaml`, `builder/pccm_builder/calc_numeric.py`,
`calc_oracle.py`, `calc_fingerprint.py`, `calc_loader.py`, `src/`, `bootstrap/`,
`readiness/`, every `.ps1`, and the Step-1/Step-2 test modules.

---

## 2. Build input flow

```
workbook.yaml ─┐
input_contract.yaml ─┤
driver_contract.yaml ─┼─► load ─► build_workbook ─► PCCM_stageA.xlsx
structure_contract.yaml ─┤              │
calc_contract.yaml ──────┘              ├─► emit_stage_b   ─► modConstants.bas
        │                               │                    stage_b_manifest.json
        │                               │                    phase4_scenarios.json
        └─ validate_calc_against ───────┼─► emit_calc_artifacts ─► modCalcContract.bas
           (BEFORE any emission)        │                          phase5_cases.json
                                        └─► verify_workbook(… , calc)
```

The calculation contract is loaded and cross-validated against all four other
authorities **before** a workbook is written. A malformed or cross-invalid
contract fails with `CALCULATION CONTRACT ERROR` and exit code **2**, and no
workbook is produced. The validation rules themselves are the accepted Step-1
loader's and are **not** restated in the builder.

---

## 3. Rendering architecture

`calc_render.py` consumes `CalcContract`, a `Worksheet` and the existing
`StyleBook`. It defines no font, colour or style literal of its own; the manifest
remains the single presentation authority. It never imports `calc_oracle`, and an
AST test proves it — read from the parse tree rather than the text, because the
module's own docstring says it does not import the oracle and a textual search
would find that sentence.

**Column widths.** None were added. The `_Calc` manifest entry still declares
`A:E` only, and the five new bands take Excel's default width. §11 of the Step-3
instruction permits adding presentation-only widths to `workbook.yaml`; nothing
required them, and inventing 39 width numbers would have been inventing business
semantics to fill a gap that does not exist.

---

## 4. Phase-4 territory

Rows **1–11** of `_Calc` are Phase-4 property: the permanent-ID counters at `C10`
and `C11` with their labels and notes. The renderer starts at row **13** — row 12
is a deliberate gap — and writes nothing above it. Post-build verification asserts
the counters still hold `0` with number format `0`, that their labels are intact,
and that no Phase-5 content appears anywhere in rows 1–12.

---

## 5. The physical empty-ListObject representation

**Each Phase-5 table is its header row plus exactly one physically blank data
row.** `tblCalcYears` is `H15:J16`, `tblCalcInflationFactors` is `M15:P16`,
`tblCalcFX` is `S15:U16`, `tblCalcDrivers` is `X15:AR16`, `tblCalcAnnual` is
`AU15:BB16`.

**Why not a header-only table.** A `ref` spanning only the header row describes a
ListObject with zero data rows, and Excel has no such object: deleting the last row
of a table leaves one blank row rather than none, and a single-row `ref` with
`headerRowCount=1` is the shape Excel treats as damaged and offers to repair. A
Stage-A workbook that prompts for repair on first open is not acceptable, so the
minimum PHYSICALLY valid body is used instead. It is also exactly what every
existing table in this workbook already does — the Phase-3 registers and the
Phase-4 grids all reserve blank data rows.

**The placeholder is semantically nothing.** No value, no formula, no identifier,
no zero, no Project Index 0, no SAR audit row, no Base-Year factor row. Those are
calculation outputs and do not exist until a calculation produces them.
Verification asserts every body cell is blank and reports the workspace as having
**zero semantic calculation rows**. The cell carries its column's number format so
a later runtime write lands in an already-correct cell and cannot change the audit
presentation by writing.

---

## 6. Initial state — blank is not zero

`calc_state` values ship as:

```
C13 blank    C14 blank    C15 blank    C16 blank
C17 "NONE"   C18 blank    C19 "NOT CALCULATED"   C20 blank
```

`C15` is blank **even though the contract declares `FP_VERSION = 1`**. The version
stamp records which algorithm produced the STORED digest, and there is no stored
digest until a successful commit writes both together. Seeding it would assert a
fingerprint that does not exist. Verification states this as its own named check.

All ten `calc_totals` cells (`C23:C32`) ship **blank**, with number format
`#,##0.00` and the accepted em-dash labels. Blank means no analytical calculation
has committed; zero would mean a calculated total of zero, and conflating the two
would be irreversible.

No formulas anywhere. No data validation on either block or on any of the five
tables — these are model-controlled audit cells, not user inputs.

---

## 7. `modCalcContract.bas` — projection rules

Generated source only. **Not executed in Step 3, not embedded in the Stage-A
`.xlsx`** (a test opens the archive and asserts no `.bas` and no `vba` part is
inside). 138 `Public Const` declarations, no `Sub`, no `Function`, no loop, no
`Mod`, no hash recurrence, no analytical formula — a test asserts each of those
absences against the code with comments stripped.

Projected, each from the authority that owns it:

| From `spec/calc_contract.yaml` | From `builder/pccm_builder/calc_fingerprint.py` |
|---|---|
| `FP_VERSION` | `FP_BASE`, `FP_MOD_1`, `FP_MOD_2`, `FP_INIT_1`, `FP_INIT_2` |
| `calc_state` / `calc_totals` columns, first/last row, value range, per-field row | `FP_STREAM_TAG`, `FP_SECTION_1..3`, `FP_TAG_TEXT/NUMBER/INTEGER` |
| the five table names, header rows, first/last columns, first-column index, column count, empty ref, first body row | |
| every per-column ordinal index | |
| the four derived-status and four attempt-result labels | |
| the four tolerance numbers | |
| the sheet name, its required visibility, the Phase-4 reserved rows | |

### How the hash constants reach the module without becoming a second authority

They are **imported**, not restated. `calc_emit.py` does `from . import
calc_fingerprint as fp` and writes `fp.FP_BASE`, `fp.FP_MOD_1` and the rest
directly into the literal. The YAML never contains them — a test asserts the
string `2147483647` appears **zero** times in `calc_contract.yaml`, so the modulus
cannot be edited into existence there.

That is asserted by moving the source rather than by inspection:
`test_mutating_the_hash_authority_changes_the_projection` patches
`calc_fingerprint.FP_BASE` to `137`, re-renders the module, and requires the
emitted constant to read `137`. If the `.bas` carried its own literal the test
would fail. The same technique proves the YAML-owned side:
`test_mutating_the_contract_changes_the_projected_geometry` moves
`tblCalcYears`'s header row to 99 in a replaced contract object and requires the
projection to follow.

**Not in this module:** the fingerprint recurrence, any analytical formula, any
calculation. It is a constants module.

---

## 8. `phase5_cases.json` — schema

Test data for the later Windows/VBA harness. Nothing in the workbook reads it, and
nothing in it defines the mathematics.

```
schema_version        1
purpose               a sentence saying it is test data
provenance            model_version, calc_contract_version, fingerprint_version,
                      and the three source modules by path
tolerances            the four contract numbers
fingerprint           constants, reference (case 26), collision_probes (case 27),
                      decimal_separator (case 35), reduction_vectors (case 36)
plan_cases[37]        id, kind, title, and per kind:
                        analytical   -> model + expected{resolved_fx,
                                        inflation_factors, discount_factors,
                                        drivers[], annual[], totals{}}
                        refusal      -> model + expected_refusal (the class name)
                        statistics   -> statistic, points, expected
                        fingerprint  -> a pointer into the fingerprint section
                        runtime_only -> why, and NOTHING else
regression_vectors    conditioning, convex_statistics, materialization, product,
                      row_order, signed_sum
```

**Cases 32, 33, 34 and 37 are `runtime_only`.** They are workbook-state and
rollback behaviours that no pure function can evidence, so they carry a `why` and
no expected numbers at all. A test asserts their record has exactly the four keys
`{id, kind, title, why}` — Python does not pretend to prove them.

### The C1/C2 regression corpus

The plan's case matrix predates the edges implementation found, so a compact named
section carries the load-bearing fixed vectors — not the 10,000-case property
sweeps, which stay in the Step-2 test modules where they belong:

* **signed_sum** — the `[1e16, 1, -1e16]` canonical-order case (expected `0.0`,
  because tier 1 owns it), the three cancellation vectors, the rounding-residual
  case (expected **`-1e292`**), the genuine out-of-range case, and `MAX + MAX`;
* **product** — the bad-order rescue, the `5e-324` result, the out-of-range-by-
  under-an-ulp case, and `1e308 * 10`;
* **convex_statistics** — the degenerate `MAX_DOUBLE` and subnormal cases, the
  exact zero, and the non-zero statistic that collapses to zero;
* **materialization** — Reproducers A, B and C evaluated end to end, plus the
  published-driver-boundary case that must still refuse;
* **conditioning** — the headline and within-year C1 cancellation models;
* **row_order** — the `1e16 / 1 / -1e16` three-driver model.

---

## 9. Golden independence

```
hand-derived literal  ->  Python oracle  ->  phase5_cases.json  ->  later VBA
```

The corpus is generated BY the accepted oracle, so comparing it back to that
oracle would prove nothing. Every emitted number is re-derived by a second,
independently written route:

1. **Plan §23's own literals**, restated in the test file:
   case 1 `Knom = 1`, `A = 1000`, `C = 1100`, `B = 100`; case 2 `Knom = 3.75`,
   `A = 1500`; case 3 `Knom = 1.1085375`, `A_nom = 1108.5375`; case 4
   `Kpv = 0.998150826446`, `A_pv = 998.150826446`, `C_pv = 1097.965909091`;
   case 5 `C_nom = 1219.39125`, `B_nom = 110.85375`; case 6 mean `105`; case 7
   central = mean = `115`, `B = 0`; case 8 mean severity `250`, `D = 75`; case 9
   `D_nom = 83.1403125`, `D_pv = 74.8613119835`; case 22 central = mean = `115`.
2. **`_exact_case`**, an exact `Fraction` evaluator written in the test from the
   plan's formulas, sharing no code with `calc_oracle.py`. It re-derives the
   inflation factors, the discount series, `Knom`, `Kpv`, every driver amount,
   every annual column and all ten headline totals from the case's INPUT DATA
   alone. **~598 emitted numbers** are checked this way, and the test asserts a
   floor of 400 so it cannot pass vacuously.
3. **The fingerprint literals**, restated in the test: the reference stream
   character for character, its length of **366** UTF-16 code units, its digest
   **`50B6EB0E26857EA7`**, the eight probe digests, and the four reduction
   remainders.

`test_the_independent_oracle_disagrees_when_the_corpus_is_wrong` and
`test_a_mutated_golden_expected_value_is_caught` prove the comparison can fail.

A field that does not apply to a driver kind is asserted **blank**, never zero —
`central_value` and the deterministic/mean-basis amounts on a Risk row, the
expected-risk amounts on a Cost Line row.

---

## 10. Determinism

`PCCM_BUILD_TIMESTAMP` fixed, two builds from the same source:

```
build/vba/modCalcContract.bas   byte-identical
build/phase5_cases.json         byte-identical
build/PCCM_stageA.xlsx          structurally identical (structural_digest)
```

No current time, no build-time random seed, no absolute machine path, no
environment-specific locale formatting inside either generated file — tests assert
the absence of `/home/`, `/tmp/`, `C:\` and the repository root, and that the
document has no `generated_at` or `timestamp` key. JSON is emitted with
`allow_nan=False`, so a non-finite value would be a build failure rather than the
non-standard `NaN`/`Infinity` tokens. Numbers use Python's shortest
round-tripping `repr`, which is stable across supported CPython versions; VBA
`Double` literals get an explicit `.0` where integral so VBA cannot type them as
`Long`.

---

## 11. The phase-version advance

The repository's established convention is that `VERSION`, the manifest's
`model_version` and `BUILDER_VERSION` move together, once per phase, and that a
Phase-4 test asserts all three agree. All three advanced **0.4.0 → 0.5.0**, and
`build_phase` became `"Phase 5 - Calculation Workspace (Gate A: source)"`, because
the generated workbook now physically contains Phase-5 blocks and must not keep
describing itself as a Phase-4-only artifact.

`test_36_version_file_matches_the_model_version` therefore moved its literal. What
it proves — that the three agree and none drifts alone — is unchanged, and
`BUILDER_VERSION` is now checked there explicitly rather than only implied. **This
is the one pre-existing test whose literal moved**, and it is called out here so
review can confirm the change is the convention rather than a weakening.

---

## 12. Post-build verification

`verify_workbook(…, calc)` now checks, against the GENERATED ARTIFACT read back
from disk:

* `_Calc` exists and is `hidden`; `_SimData` unchanged
* the Phase-4 counters and their number formats intact; no Phase-5 content in rows 1–12
* every `calc_state` and `calc_totals` label, number format, note and initial value
* `C15` blank rather than seeded with `FP_VERSION`
* no `calc_totals` cell seeded with zero
* all five tables exist on `_Calc`, with the exact `ref`, header row, first and
  last column, column count and header spelling/order
* every table column's body number format
* zero semantic rows: the one physical body row is blank
* no formula and no data validation inside any calculation table
* no data validation on either scalar block's value column
* no two calculation tables overlap; none intersects the Phase-4 reservation
* the workbook Table inventory is the prior accepted set **plus exactly the five
  Calc tables** — the "only contract-declared Tables exist" gate is **extended,
  not relaxed**

---

## 13. Negative controls

`_verify_mutated` builds the workbook, mutates the saved artifact, and re-verifies.
Each of these produces a verification failure for the intended reason:

| Mutation | Caught by |
|---|---|
| `tblCalcYears` moved one column (`I15:K16`) | table ref / column check |
| `tblCalcFX` widened one column (`S15:V16`) | ref and column-count check |
| `H15` header changed to `Project Idx` | header spelling/order check |
| `tblCalcAnnual` renamed | the Table-inventory gate |
| a `calc_state` row moved from `C17` to `C21` | initial-state check |
| `C15` seeded with `1` | the dedicated fingerprint-version check |
| `C27` seeded with `0` | the "no calc_totals cell is seeded with zero" check |
| `=SUM(AU16:AU20)` inserted at `C23` | the formula gate |
| `C10` overwritten with `7` | Phase-4 counter check |
| a fabricated `H16`/`I16` semantic row | the zero-semantic-rows check |

Three further controls act on the generated files rather than the workbook: a
mutated projected tolerance, a mutated projected hash modulus and a mutated golden
expected value are each shown to change or fail the corresponding assertion. And
`test_a_malformed_calculation_contract_fails_the_build_with_exit_code_two` moves
`tblCalcYears`'s header row in a copied contract, runs the real entry point, and
requires exit code 2, a `CALCULATION CONTRACT ERROR`, and **no workbook on disk**.

---

## 14. Consumed, but not calculated

`test_stage_a_renders_the_workspace_without_calling_the_oracle` replaces
`calc_oracle.calculate` with a function that raises, builds the workbook and runs
full verification — which passes. Stage-A workbook emission therefore calls the
oracle **nowhere**: it does not compute A/B/C/D/E, does not resolve FX or inflation
from workbook inputs, and does not populate an annual output.

The CASES generator is a separate artifact and IS allowed to call the oracle; that
is what produces the expected values, and §9 is what keeps it honest.

**No simulation leakage.** `calc_render.py`, `calc_emit.py` and `calc_cases.py`
contain no reference to `inpMonteCarloIterations`, `inpRandomSeed`,
`inpSelectedConfidenceLevel`, `MRG32k3a`, `Randomize`, `random` or `seed`, and
`_SimData` is byte-for-byte the same whether the calculation contract is supplied
or not.

---

## 15. Exact test counts

Run from a clean extraction, Linux, Python 3.11.

```
python -m pytest pccm/tests/ -q        962 passed, 0 failed
python pccm/builder/build_stage_a.py   351 passed, 0 failed
```

Standalone:

```
python pccm/tests/test_phase5_stage_a.py                  53 passed, 0 failed
python pccm/tests/test_phase5_numeric.py                  94 passed, 0 failed
python pccm/tests/test_phase5_oracle.py                  111 passed, 0 failed
python pccm/tests/test_phase5_calc_contract_validation.py 151 passed, 0 failed
python pccm/tests/test_phase5_fingerprint.py              52 passed, 0 failed
```

| Module | Step 2 accepted | Step 3 |
|---|---|---|
| `test_phase1_manifest_validation.py` | 10 | **10** |
| `test_phase1_structure.py` | 21 | **21** |
| `test_phase2_contract_validation.py` | 42 | **42** |
| `test_phase2_inputs.py` | 40 | **40** |
| `test_phase3_driver_contract_validation.py` | 31 | **31** |
| `test_phase3_drivers.py` | 28 | **28** |
| `test_phase3_verifier_intersection.py` | 12 | **12** |
| `test_phase4_oracle.py` | 68 | **68** |
| `test_phase4_stage_b_source.py` | 155 | **155** |
| `test_phase4_structure.py` | 43 | **43** |
| `test_phase4_structure_contract_validation.py` | 52 | **52** |
| `test_phase5_calc_contract_validation.py` | 151 | **151** |
| `test_phase5_fingerprint.py` | 52 | **52** |
| `test_phase5_numeric.py` | 94 | **94** |
| `test_phase5_oracle.py` | 111 | **111** |
| `test_phase5_stage_a.py` | — | **53** |
| **total** | **910** | **962** |

**Every existing count is unchanged.** The 910 → 962 delta is exactly the 53 new
Step-3 tests. No Step-2 test was weakened, added to or removed; the numeric and
oracle modules are byte-identical to the accepted Step-2 package.

### Stage-A verification count

```
old checks retained = 181
new checks added    = 170
final checks        = 351      (0 failures)
```

`test_the_stage_a_verification_is_extended_and_nothing_was_dropped` proves this by
matching check DESCRIPTIONS, not by comparing totals: every one of the 181 prior
checks is still present in the extended run, and no old check was replaced by a
broader one to keep the number small. Run without the calculation contract, the
same artifact yields 180 passes and **1 legitimate failure** — the
"only contract-declared Excel Tables exist" gate, because the artifact now carries
five tables the caller did not declare. That is the gate working.

---

## 16. Deliberately unimplemented

* the Phase-5 VBA module set — `modCalcFactors`, `modCalcAnalytical`,
  `modCalcFingerprint`, `modCalcResolve`, `modCalcCheck`, `modCalcReport`,
  `PCCM_Calculate`. None exists, and none was added to the Stage-B manifest;
* any Calculate button or user-facing calculation entry point;
* runtime row creation, per-kind N/A styling and the not-applicable-is-blank rule
  in practice — Step 3 preserves the schema that will enable it;
* the Windows harness, and any Windows run;
* RNG, sampling, simulation output, and anything on `_SimData`.

---

## 17. Next step — NOT started

Step 4, whatever review scopes it to be.

> **NO EXECUTABLE PHASE-5 CALCULATION VBA WAS IMPLEMENTED.**
> **NO PHASE-5 NUMERICAL VBA WAS IMPLEMENTED.**
> **NO WINDOWS HARNESS WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **NO CALCULATION BUTTON WAS ADDED.**
> **NO RNG / MONTE CARLO IMPLEMENTATION WAS ADDED.**
> **STEP 4 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

**PHASE 5 GATE A STEP 3 READY FOR INDEPENDENT REVIEW**
