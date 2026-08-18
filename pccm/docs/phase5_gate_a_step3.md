# Phase 5 — Gate A — Step 3: Stage-A calculation workspace emission

**Status: CORRECTED ONCE after independent review — ready for re-review.**

Step 2 is accepted and closed. This step makes `spec/calc_contract.yaml` a real
build input, puts the physical `_Calc` calculation workspace into the generated
workbook, and emits the two generated Phase-5 artifacts a later VBA implementation
will consume. It is Linux-only.

**The purpose of this step is REPRESENTATION, not calculation.** The accepted
numerical oracle remains the semantic authority; `calc_contract.yaml` remains the
physical workbook-representation authority. The workbook ships with no calculated
result of any kind.

**Round 1** (§18) — the workbook emission, the verifier, the constants projection
and the Stage-A architecture were confirmed sound and are unchanged. Two
acceptance-corpus gaps were closed: the Gate-B fingerprint vector set was
incomplete, and golden independence was a spot check rather than exhaustive. One
documented count was wrong. **The workbook and `modCalcContract.bas` are
byte-for-byte unchanged by the correction; only `phase5_cases.json` expanded.**

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
| `tests/test_phase5_stage_a.py` | **new** — 54 tests |

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
fingerprint           constants
                      reference          (case 26) stream, code_units, digest
                      collision_probes   (case 27) eight probes and digests
                      numeric_encodings  (case 26) the TEN locked canonical
                                         encodings: label, value, expected
                      utf16_vectors      (case 26) key, text, code_point_count,
                                         utf16_length, code_units, signed_ascw,
                                         canonical_text_field
                      decimal_separator  (case 35) label, value, expected, point,
                                         comma - eleven vectors
                      reduction_vectors  (case 36) modulus_name, modulus, h, u, x,
                                         remainder, double_only_remainder
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

### The complete Gate-B fingerprint vector set

Plan §24.1 requires real Windows/VBA to exercise the canonical encoder and the
reducer directly, and locks *"every expected value comes from
build/phase5_cases.json"*. Round 1 found two of the six required families missing,
which would have forced the Windows harness to hardcode expectations outside the
corpus. Both are now carried.

**A — the ten locked canonical numeric encodings** (`fingerprint.numeric_encodings`):

| label | value | expected |
|---|---|---|
| `0` | `0.0` | `0.0000000000000000E+00` |
| `-0` | `-0.0` | `0.0000000000000000E+00` |
| `1` | `1.0` | `1.0000000000000000E+00` |
| `-1` | `-1.0` | `-1.0000000000000000E+00` |
| `0.1` | `0.1` | `1.0000000000000001E-01` |
| `1e-20` | `1e-20` | `9.9999999999999995E-21` |
| `1e+20` | `1e+20` | `1.0000000000000000E+20` |
| `0.1 + 0.2` | `0.30000000000000004` | `3.0000000000000004E-01` |
| `MAX_DOUBLE` | `1.7976931348623157e+308` | `1.7976931348623157E+308` |
| `minimum subnormal` | `5e-324` | `4.9406564584124654E-324` |

Each record carries a textual `label`, so `-0` stays distinguishable from `0` even
though both encode identically and a JSON consumer may render the two values the
same way. A test asserts the two labels differ and that the second value really
carries a negative sign bit.

**D and E — direct UTF-16 / `AscW` vectors** (`fingerprint.utf16_vectors`):

| key | text | code points | UTF-16 units | `code_units` | `signed_ascw` | canonical field |
|---|---|---|---|---|---|---|
| `bmp_above_7fff` | `高` | 1 | 1 | `[39640]` | `[-25896]` | `S1:高` |
| `non_bmp` | `😀` | 1 | 2 | `[55357, 56832]` | `[-10179, -8704]` | `S2:😀` |
| `mixed_length_prefix` | `A😀` | 2 | 3 | `[65, 55357, 56832]` | `[65, -10179, -8704]` | `S3:A😀` |

`signed_ascw` is what VBA's `AscW` returns — a SIGNED 16-bit `Integer`, so every
unit above `U+7FFF` comes back negative; `code_units` is the same value after the
`+ 65536` normalisation. The generator derives the signed form and asserts the
round trip through `calc_fingerprint.normalise_code_unit`, so a drift between the
two fails the build rather than shipping a wrong expectation. The third vector
exists to prove the length prefix counts CODE UNITS: `A😀` is two code points and
three units, so a VBA implementation using `Len()` would emit `S2:` and be caught.

**B — separator injection** (`fingerprint.decimal_separator`) now covers the same
ten labelled vectors plus the hostile `-9.87e-5` that drove the Step-1 positional
normalisation fix, and each record carries an `expected` literal. The test asserts
`point == LOCKED` **and** `comma == LOCKED`, not merely `point == comma`: the
canonical form always uses `.`, so a host separator of `,` is normalised away, and
proving invariance alone would say nothing about correctness.

**C and F** — the four reduction vectors and the complete 366-code-unit reference
stream were already carried and are unchanged.

**No new fingerprint authority.** `calc_cases.py` derives every emitted value from
`calc_fingerprint.py`; the locked literal copies live in the test file only, and
none of them was added to `calc_contract.yaml`.

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

## 9. Golden independence — exhaustive, and accounted path by path

```
hand-derived literal  ->  Python oracle  ->  phase5_cases.json  ->  later VBA
```

The corpus is generated BY the accepted oracle, so comparing it back to that
oracle would prove nothing. Round 1 found the first attempt insufficient: it
covered only `plan_cases` of kind `analytical`, only a subset of the fields inside
them, and only spot checks of the regression corpus, behind a `checked > 400`
floor that could pass while hundreds of values went untested.

**The floor is gone. Coverage is now proved by a LEDGER.**

`_expectation_paths(document)` enumerates every expectation leaf in the JSON —
every leaf under `plan_cases[*].expected`, every
`plan_cases[*].statistics[*].expected`, every `plan_cases[*].expected_refusal`,
every leaf under `regression_vectors[*][*].expected`, every
`regression_vectors[*][*].expected_refusal`, and the whole `fingerprint` subtree.
Each independent check records the path it validated. The test then requires the
two SETS to be identical, in both directions:

```
expectation leaf paths  : 1979
independently validated : 1979
missing                 : 0        (nothing entered the corpus unproved)
extra                   : 0        (no check claimed a path the corpus lacks)
```

| root | leaves |
|---|---|
| `plan_cases` | 1268 |
| `regression_vectors` | 516 |
| `fingerprint` | 195 |

Model inputs (`model`, `terms`, `factors`, `points`), titles and prose are not
expectations and are excluded; every leaf under `fingerprint` is included, which
is why the corpus carries data and no explanatory `note` strings.

### The independent reference

`_reference_payload(model)` rebuilds the COMPLETE expected payload of any
analytical model — the resolved FX map, every inflation factor row, the discount
series, all twenty-one driver fields plus the weights, all eight annual fields and
all ten headline totals. It calls **neither** `calc_oracle.calculate`, **nor**
`calc_cases.evaluate`, **nor** any production rescue helper; it is written from
the plan's rules using `Fraction`, plain Double arithmetic and locked literals.

It models the accepted TWO TIERS rather than pure rational arithmetic, because
tier 1 owns its result:

* a product is left-to-right Double, refusing overflow and a non-zero collapse to
  zero; only if that fails does the exact rational value, correctly rounded, take
  over;
* a series is the staged per-contribution path, then one exact compound
  expression — the materialization rule of Erratum C2, written out from the
  specification;
* PV's tier 1 forms `nominal × discount` from the MATERIALIZED nominal, exactly as
  the accepted staging does;
* headline totals accumulate in canonical driver order.

That distinction is load-bearing. `[1e16, 1, -1e16]` sums to `0.0` in canonical
order and to `1.0` in exact arithmetic, and `0.0` is the required answer — a
reference built on exact rationals alone would have declared the corpus wrong.

Comparison is **exact equality wherever it holds**, which is almost everywhere,
including every zero, every subnormal and every `MAX_DOUBLE / 2`. A relative
tolerance of `1e-11` covers only the places where the plan's mandated stable forms
legitimately differ from exact rational arithmetic, and a per-payload `scale`
covers a value that cancelled to near-zero.

### The regression corpus is not exempt

* **model-based** (`conditioning`, `materialization` ×4, `row_order`) — the same
  `_reference_payload` validates the COMPLETE emitted payload, not selected
  fields;
* **helper-level** (`signed_sum`, `product`, `convex_statistics`) — exact
  `Fraction` arithmetic derives the exact value, classifies its range with
  `abs(exact) <= Fraction(MAX_DOUBLE)` tested BEFORE `float(exact)`, and yields
  either the correctly rounded Double or a refusal. The production helper is never
  asked what its own answer should be. A zero-uncertainty distribution is held to
  the stronger statement — it must return its point EXACTLY, not merely correctly
  rounded;
* **the materialized-boundary refusal** — `_check_materialization_refusal`
  independently proves from exact arithmetic that CL-001's own Mean-Basis Nominal
  is `2 × MAX_DOUBLE`, which exceeds `MAX_DOUBLE`, **and** that the headline it
  feeds cancels to exactly zero. Both halves matter: the first is why the refusal
  is correct, the second is why it is not obvious.

### The plan's own literals

Plan §23's hand-derived numbers are restated in the test file and checked
separately: case 1 `Knom = 1`, `A = 1000`, `C = 1100`, `B = 100`; case 2
`Knom = 3.75`, `A = 1500`; case 3 `Knom = 1.1085375`, `A_nom = 1108.5375`; case 4
`Kpv = 0.998150826446`, `A_pv = 998.150826446`, `C_pv = 1097.965909091`; case 5
`C_nom = 1219.39125`, `B_nom = 110.85375`; case 6 mean `105`; case 7 central =
mean = `115`, `B = 0`; case 8 mean severity `250`, `D = 75`; case 9
`D_nom = 83.1403125`, `D_pv = 74.8613119835`; case 22 central = mean = `115`.

The fingerprint half is entirely hand-written: the reference stream character for
character, its 366-code-unit length, its digest `50B6EB0E26857EA7`, the eight
probe inputs and digests, the ten canonical encodings, the three UTF-16 vectors
and the four reduction vectors — the last checked twice, once against the locked
remainder and once by independently recomputing `x = h × 131 + u` and `x mod m`.

A field that does not apply to a driver kind is asserted **blank**, never zero.

## 10. Determinism

**This correction changed neither the workbook nor the constants module.**
Rebuilding before and after the patch with a fixed timestamp:

```
build/vba/modCalcContract.bas   byte-identical before and after
build/PCCM_stageA.xlsx          structurally identical before and after
build/phase5_cases.json         expanded, as intended
```

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

### Against the generated workbook

`_verify_mutated` builds the workbook, mutates the saved artifact, and re-verifies.
Each produces a verification failure for the intended reason:

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

### Against the acceptance corpus

Each was applied to `calc_cases.py` on a working copy, the suite was run, and the
source restored. All seven fail against the corpus as it stood at `d82471e`:

| Sabotage | Result |
|---|---|
| one of the ten numeric canonical encodings removed | **fails** — the path ledger reports a missing expectation |
| `1e-20` expected string changed to `1.0000000000000000E-20` | **fails** — the hand-written literal disagrees |
| the `U+9AD8` vector removed | **fails** — the ledger reports the missing vector |
| `U+9AD8` normalised unit changed from `39640` | **fails** — the locked UTF-16 literal disagrees |
| `😀` made to contribute one unit instead of two (`len()` instead of `utf16_length`) | **fails** — the locked length and the `utf16_length == len(code_units)` property both break |
| a NON-spot-checked value tampered deep inside a materialization payload (`annual[2].total_pv`) | **fails** — `test_a_tampered_non_spot_checked_regression_value_is_caught` |
| a new expected field added with no matching check | **fails** — `test_a_new_expected_field_cannot_escape_independent_validation` |

The last one is the important one: it proves future corpus growth cannot silently
escape independent validation, which a numeric floor could never do.

Two further controls act on the generated constants: a mutated projected tolerance
and a mutated projected hash modulus are each shown to change the emitted `.bas`.
And `test_a_malformed_calculation_contract_fails_the_build_with_exit_code_two`
moves `tblCalcYears`'s header row in a copied contract, runs the real entry point,
and requires exit code 2, a `CALCULATION CONTRACT ERROR`, and **no workbook on
disk**.

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
python -m pytest pccm/tests/ -q        964 passed, 0 failed
python pccm/builder/build_stage_a.py   351 passed, 0 failed
```

Standalone:

```
python pccm/tests/test_phase5_stage_a.py                  54 passed, 0 failed
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
| `test_phase5_stage_a.py` | — | **54** |
| **total** | **910** | **964** |

**Every existing count is unchanged.** The 910 → 964 delta is exactly the 54 new
Step-3 tests. No Step-2 test was weakened, added to or removed; the numeric and
oracle modules are byte-identical to the accepted Step-2 package.

**COUNT CORRECTION.** The first Step-3 submission documented 962. That was a stale
reading taken before the last test of the round was added, and it was
arithmetically impossible against its own "910 + 53" statement — review was right
to reject it. The correct pre-correction baseline was **963**; this correction
replaced three spot-check tests with four exhaustive ones, so the current total is
**964**. No test was removed to make a number true.

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
