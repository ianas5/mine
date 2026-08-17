# Phase 5 — Gate A — Step 2: pure analytical oracle and safe numerical semantics

**Status: CORRECTED THREE TIMES after independent review — ready for re-review.**

Step 1 is accepted and closed. This step locks and tests the pure numerical
semantics that later Stage-A emission and later VBA must implement. It is
Linux-only, in-memory, and side-effect free.

**Round 1** (§16) — four blocking correctness issues and one test-portability
defect. One of them, the conditioning definition, proved that the accepted plan
§15 did not satisfy its own stated objective, and is the narrow, justified
exception to Step 1 being closed. Erratum C1 is now **accepted**.

**Round 2** (§17) — three remaining blockers, all in the resolution and metadata
layers rather than the mathematics: conditioning metadata could refuse a
representable model, driver reference fields leaked raw `AttributeError`, and
`resolved_fx` did not match the locked `tblCalcFX` row rule.

**Round 3** (§18) — one remaining numerical correctness class: a mathematically
representable final Double was still sometimes refused because an *intermediate*
operation overflowed or underflowed, which is precisely what plan §19.2 exists to
prevent. Repaired by a two-tier signed sum and a three-tier convex statistic, and
recorded as narrow plan **Erratum C2**. The algorithms both later VBA and any
independent reimplementation must follow are specified in §18.4 and §18.5.

**No business rule, formula or tolerance NUMBER changed in any round.**
Round 2 changed **no design at all** — `docs/phase5_plan.md`,
`spec/calc_contract.yaml`, `calc_loader.py`, `calc_fingerprint.py`,
`build_stage_a.py`, `src/` and `bootstrap/` are byte-identical to the round-1
package. Erratum C1, the canonical ordering, the A/B/C/D/E and annual definitions,
the safe-product strategy and the stable distribution formulas are untouched.
Round 3 changed two production modules and the two Step-2 test modules, added
plan Erratum C2, and changed **no contract**: `spec/calc_contract.yaml`,
`calc_loader.py`, `calc_fingerprint.py`, `build_stage_a.py`, `src/`, `bootstrap/`
and `tools/` are byte-identical to the round-2 package.

---

## 0. The statements this step must make explicitly

> **NO VBA WAS IMPLEMENTED.**
> **NO PHASE-5 WORKBOOK BLOCK WAS EMITTED.**
> **NO STAGE-A EMITTER WAS CHANGED.**
> **NO `phase5_cases.json` WAS EMITTED.**
> **NO WINDOWS HARNESS WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **STEP 3 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

`builder/build_stage_a.py` is byte-identical, does not read `calc_contract.yaml`,
and Stage-A verification is unchanged at **181 / 181**. Nothing under `src/`,
`bootstrap/` or `readiness/` was touched, and no `.ps1` file was edited.

---

## 1. Scope implemented

| # | Delivered | File |
|---|---|---|
| A | pure analytical calculation oracle | `builder/pccm_builder/calc_oracle.py` |
| B | pure numerical safety and factor primitives | `builder/pccm_builder/calc_numeric.py` |
| C | plain-data resolution semantics | inside `calc_oracle.py`, layer 1 |
| D | analytical / refusal / reconciliation golden tests | two new test modules |
| E | this document | `docs/phase5_gate_a_step2.md` |

**The one supporting module is justified, not decorative.** `calc_numeric.py` is
the layer later VBA `modCalcFactors` maps to: safe arithmetic, stable statistics,
iterative factor series and the cancellation-aware tolerance. Keeping it separate
is what lets a static test prove the kernel has no Excel dependency, and lets the
primitives be tested against the plan's own literals without an analytical model
in the way. Nothing else was fragmented out.

## 1.1 Deliberately NOT implemented

`phase5_cases.json` · `_Calc` workbook emission · Stage-A Phase-5 blocks · any
`build_stage_a.py` change · `modCalcContract.bas` · `modCalcFactors` ·
`modCalcAnalytical` · `modCalcResolve` · `modCalcCheck` · `modCalcReport` ·
`PCCM_Calculate` · the five status accessors · transactional write-back · the
attempt-state machine · any Windows harness change · the transient Windows
vector-coverage diagnostic of plan §24.1 · anything in Phase 6.

Step 2 preserves the failure *classification* the attempt-state machine will need;
it does not implement the machine.

---

## 2. Files changed

| File | Change |
|---|---|
| `builder/pccm_builder/calc_oracle.py` | **new** — data model, resolution/validation, canonical ordering, analytical kernel, contribution-conditioned reconciliation |
| `builder/pccm_builder/calc_numeric.py` | **new** — safe arithmetic, stable statistics, factor series, tolerance, failure hierarchy |
| `tests/test_phase5_numeric.py` | **new** — 52 tests |
| `tests/test_phase5_oracle.py` | **new** — 96 tests |
| `docs/phase5_gate_a_step2.md` | **new** — this document |
| `docs/phase5_plan.md` | **modified in round 1 only** — §15 erratum C1 plus its §0 register entry. Unchanged in round 2 |
| `spec/calc_contract.yaml` | **modified in round 1 only** — `conditioning_terms` names (erratum C1). Unchanged in round 2 |
| `builder/pccm_builder/calc_loader.py` | **modified in round 1 only** — `LOCKED_CONDITIONING_TERMS`. Unchanged in round 2 |
| `tests/test_phase5_calc_contract_validation.py` | **modified in round 1 only** — the conditioning-term expectations; count unchanged at 151 |
| `tools/package_review.py` | **modified in round 2** — writes `PROVENANCE.txt` into the archive so a review package identifies its own commit |

**Unchanged, and verified unchanged:** `builder/pccm_builder/calc_fingerprint.py`,
`builder/pccm_builder/__init__.py`, `builder/build_stage_a.py`, all four earlier
specifications, every Phase 1–4 source and test, and every `.ps1`.

The oracle is **not exported from `pccm_builder/__init__.py`**. That is
deliberate: exporting it would suggest the build consumes it, and the build does
not. Tests import `pccm_builder.calc_oracle` directly.

---

## 3. The pure data model

Immutable dataclasses, plain values only.

```
AppliedTimeline    base_year · start_year · duration
                   -> last_year, project_years() = [(index from 1, calendar year)]

FxRow              currency · rate            rate is loosely typed on purpose:
                                              blank and non-numeric are real states
CostDriver         permanent_id · distribution · currency · inflation_profile
                   min_value · most_likely · max_value · profile_weights · quantity
RiskDriver         same, with probability instead of quantity

CalculationModel   timeline · discount_rate · fx_rows · inflation_rates
                   cost_drivers · risk_drivers
Tolerances         the four locked constants, PASSED IN

InflationFactorRow profile · calendar_year · annual_rate|None · cumulative_factor
DriverFactors      the 21-field audit record
AnnualRow          project_index · calendar_year · six series values
AnalyticalTotals   A/B/C/D/E, nominal and PV
CalculationResult  totals · annual · drivers · inflation_factors
                   · discount_factors · resolved_fx
IdentityCheck      name · left · right · difference · allowance · holds
```

**Load-bearing semantics, not naming choices:**

- **identity is the permanent ID** — `profile_weights` travel *with* their driver;
- **row position is not data** — it is not an input to anything, and every
  accumulation runs in **canonical order**: ascending permanent ID, ordinal on
  UTF-16 code units, using the comparison imported from `calc_fingerprint`. All
  six permutations of an order-sensitive fixture are asserted to give an identical
  complete `CalculationResult` (§16.2);
- **applied timeline only** — entered-but-not-applied values never reach here;
- **`fx_rows` is a sequence, not a mapping** — a duplicated currency has to be
  representable before it can be refused;
- **a missing year and a `None` rate are different facts**, and both are refused
  for a referenced profile.

---

## 4. Layering

```
layer 1  resolution / validation   plain data in, resolved numbers out
layer 2  numerical kernel          resolved numbers in, results out
```

The split is not decoration. Layer 1 is the only place that knows a currency can
be missing or a profile incomplete; layer 2 receives numbers already proven
usable, so **its failures can only ever be representability failures**. That is
what makes the refusal classification below honest rather than a label.

The later mapping is meant to be obvious:

| Step-2 code | later VBA |
|---|---|
| `resolve_fx`, `resolve_inflation`, `_resolve_weights`, `_resolve_three_point` | `modCalcResolve` |
| `calc_numeric.py` in full | `modCalcFactors` |
| `precomputed_factors`, `_accumulate_totals`, `_annual_series`, `reconcile` | `modCalcAnalytical` |

No VBA is designed here.

---

## 5. Refusal hierarchy

```
OracleError
├── CalculationRefusal              -> the future PCCM_Calculate REFUSED path
│   ├── ModelInputRefusal           invalid user or model input
│   └── NumericalRangeRefusal       a representability limit
└── OracleInvariantError            the implementation disagrees with itself
```

`OracleInvariantError` is deliberately **not** under `CalculationRefusal`, and a
test asserts it. A reconciliation identity that fails means the inputs were
accepted, the calculation ran, and two independently accumulated quantities that
must agree do not. Reporting that to a user as "your model is invalid" would be
wrong, and would hide a defect behind a business message.

Every input refusal names its subject — permanent ID, currency, inflation
profile, calendar year, or project year — and tests assert the subject appears in
the message rather than merely that a refusal occurred.

---

## 6. Safe arithmetic design

`safe_add` · `safe_subtract` · `safe_multiply` · `safe_divide` ·
`safe_accumulate` · `safe_sum` · `is_usable_double`.

Python does not raise on float overflow — it produces `inf` — so each primitive
checks its result explicitly. The **observable contract** is what must match VBA:

| Situation | Behaviour |
|---|---|
| finite representable result | value returned |
| genuinely unrepresentable result | `NumericalRangeRefusal` |
| zero divisor | `NumericalRangeRefusal`, refused outright |
| product or quotient of non-zero operands collapsing to exactly zero | `NumericalRangeRefusal` |
| NaN or infinity | never returned, never accepted |

IEEE-754 `float` **is** the Double semantic reference. `Decimal` is deliberately
not the calculation engine: computing more accurately than the target would hide
the representability failures this design exists to surface. `Decimal` and
`Fraction` appear only in independent test oracles.

Accumulation is checked **at the term that fails**, not at the end, so a refusal
can name the driver, profile or year responsible instead of reporting that a total
came out infinite.

### Stable statistics — mandatory forms

```
Triangular mean   Min/3 + ML/3 + Max/3
Beta-PERT mean    Min/6 + ML*(2/3) + Max/6         (never (4*ML) first)
Uniform midpoint  Min/2 + Max/2
```

Deterministic central: `ML` for Triangular and Beta-PERT, midpoint for Uniform.
Central Basis label: `ML` / `ML` / `Midpoint`.

### Iterative factor series

```
infl(BaseYear) = 1        infl(Y) = infl(Y-1) * (1 + rate_Y)
disc(1)        = 1        disc(t) = disc(t-1) / (1 + r)
```

Never a power. `(1+r)**(t-1)` can overflow as an intermediate where the reciprocal
is representable, and it cannot say which year failed. Project year 1 is period 0.

---

## 7. The stable product strategy

Plan item 12 requires that a representable product not be lost to a bad
multiplication order. `safe_product` is **two-tiered**:

1. **Left to right first.** Every ordinary model takes this path, and the result
   is bit-for-bit what a naive implementation produces. **No existing value
   moves.**
2. **Only if tier 1 fails**, re-evaluate in a magnitude-balanced order: start at
   `1.0`, take the smallest remaining magnitude while the running product is
   `>= 1`, the largest otherwise.

`1e308 * 10 * 0.01` overflows at the first step left to right; tier 2 evaluates it
as `1e308 * 0.01 * 10 = 1e307`, which is the correct, representable answer. A
genuinely unrepresentable product such as `1e308 * 10` is still refused — no
ordering rescues it, and none should.

**Numerical edge, stated rather than hidden.** Floating-point multiplication is
commutative but **not associative**, so a reordered evaluation can differ from the
left-to-right one in the last unit in the last place. Tier 2 therefore runs only
where tier 1 produced *no value at all*, so the alternative is refusing a valid
calculation. Its ordering is fully deterministic, and **a later VBA implementation
must reproduce both tiers, in this order, to match.**

No arbitrary business cap was invented anywhere.

---

## 8. Cancellation-aware tolerance, computed stably

```
allowance = max(absolute_floor, coefficient * scale_floor + Sum_i coefficient*|term_i|)
```

The coefficient is **distributed over the terms** rather than multiplied by their
sum. Both are the same number; only the distributed form avoids an intermediate
that overflows. With three terms near `1.5e308` the raw scale is `inf` while the
allowance `4.5e296` is perfectly representable, and a test asserts exactly that.

Neither locked constant is loosened: the floor stays `1e-6`, the coefficient
stays `1e-12`.

---

## 9. Golden cases implemented

| Cases | Coverage |
|---|---|
| **1–13** | the hand-derived arithmetic: Knom/Kpv, FX, compounded inflation, PV, the three distributions, EMV, `Base = Start`, `Base < Start`, zero and negative inflation |
| **14–25** | refusals *and* the acceptances a naive implementation over-blocks |
| **28** | stable means where the naive numerator overflows |
| **29** | discount underflow refusing at project year 34 |
| **30** | cancellation-heavy reconciliation |
| **31** | the explicit Base-Year inflation row |

**Left in Step 1, correctly:** 26 (fingerprint reference), 27 (collision vectors),
35 (locale reference semantics), 36 (modular reduction).

**Left for later, correctly:** 32 (status revert), 33 (mid-write rollback), 34
(status axes), 37 (commit-boundary failure). These need workbook state and a
running VBA transaction. **Python does not prove them, and nothing here pretends
it does.**

### Hand-derived independence

Every expected value is a literal transcribed from plan §23 or derived by an
independent exact-rational calculation in the test. None is obtained by calling
another `calc_oracle.py` function. A separate test re-derives case 3/4 with
`Fraction` arithmetic as a **second** independent check — supplementing the
accepted literals, not replacing them.

**Golden comparisons use a relative tolerance of `1e-12`, not exact equality**, and
§12 explains why that is required rather than convenient.

---

## 10. Additional required coverage

Uniform ML populated has no influence (five different ML values, identical
totals) · numeric zero weight accepted · blank weight refused and explicitly not
zero · weights follow the permanent ID under driver reordering · wrong profile
length refused · non-numeric weight refused · unreferenced bad FX accepted ·
unreferenced incomplete inflation accepted · referenced bad/missing/duplicate FX
refused · referenced missing/incomplete profile refused · Quantity scales the
contribution but never `Knom`/`Kpv` · Probability scales the EMV but never
`Knom`/`Kpv` · probability 0 valid with zero EMV · probability 1 valid ·
probability outside `[0,1]` refused · negative but ordered cost values allowed ·
negative but ordered risk impacts allowed · out-of-order three-point sets refused
· Uniform `Min > Max` refused · invalid distribution refused · **empty driver set
not refused** · SAR invariant enforced even with no drivers and no foreign
currency · `Base Year > Start Year` refused · blank/non-numeric discount refused.

The annual series is asserted to use the **mean** basis: its sum reconciles to
`C`, and is asserted to differ from `A` by more than 1 SAR so the test cannot pass
by coincidence.

---

## 11. Architecture boundary tests

Three complementary checks, none of them a vocabulary ban:

1. **Executable imports** — the module source is AST-parsed and every
   `import` / `from … import` top-level name is checked against
   `openpyxl`, `win32com`, `pythoncom`, `xlwings`, `random`, `secrets`, `numpy`,
   `scipy`.
2. **Referenced identifiers** — AST `Name` and `Attribute` nodes only, checked
   against `Workbook`, `Worksheet(s)`, `Range`, `Cells`, `ListObject(s)`,
   `ThisWorkbook`, `ActiveWorkbook`, `Application`, `Rnd`, `Randomize`,
   `percentile`, `quantile`. Comments never enter the AST and docstrings are
   string constants, so **prose about later phases cannot trip it** — and a
   further test asserts the modules *do* discuss `ListObject` and `Monte Carlo` in
   their documentation while still passing.
3. **A fresh interpreter** loads both modules *without* the `pccm_builder`
   package — whose `__init__` legitimately imports openpyxl — runs golden case 1,
   asserts `A_nom == 1000.0`, and asserts no forbidden module was ever loaded into
   `sys.modules`. This is the strongest form: the oracle is proven usable with no
   Excel library present at all.

A fourth test asserts neither module references `open`, `read_text`, `write_text`,
`safe_load`, `dump` or `mkdir`: the pure layer performs no file I/O, and reads no
YAML.

The distribution adapter is tested as **non-authoritative**: its three keys are
asserted to equal the `config_tables.distributions` values in
`input_contract.yaml`, so an upstream change fails loudly instead of silently
widening or narrowing what the kernel accepts.

---

## 12. Numerical edges discovered

Four, all reported rather than worked around.

### 12.1 The mandated stable forms are not bit-identical to the naive forms

`(80 + 4*100 + 150)/6` is exactly `105.0`. The **mandated**
`Min/6 + ML*(2/3) + Max/6` gives `104.99999999999999` — one ulp away, a relative
deviation of `1.4e-16`.

This is inherent to §19.2, which requires the stable form, and is not a defect.
Its consequence is procedural: **hand-derived exact literals cannot be asserted by
exact equality**, and VBA must use the same stable form to reproduce Python. A
test pins the deviation so it cannot drift unnoticed.

### 12.2 Ordinary Double rounding separates literals from results

`100 * 1.1085375` is `110.85374999999999`, because `1.1085375` has no exact binary
representation. Case 5's hand-derived `B_nom = 110.85375` is therefore the exact
*mathematical* value and the computed Double differs in the last place. Same
category, different cause from 12.1 — nothing to fix, everything to state.

### 12.3 Reordering a product can move the last ulp

Tier 2 of `safe_product` is not associative-equivalent to tier 1. Mitigated by
running tier 2 only after tier 1 has produced no value at all, and by making the
order deterministic. **Carried forward as a VBA requirement** (§7).

### 12.4 The locked conditioning scale did not meet its own objective

Reported in the first Step-2 submission as a finding about the accepted plan
rather than about this code, and **now corrected by review as plan §15 erratum
C1**. Review additionally proved the same defect in the annual identities, which
the first submission had believed immune. §16.1 carries the full account.

---

## 13. Exact test counts

Run from a clean extraction, Linux, Python 3.11.

```
python -m pytest pccm/tests/ -q        879 passed, 0 failed
python pccm/builder/build_stage_a.py   181 passed, 0 failed
```

Standalone:

```
python pccm/tests/test_phase5_numeric.py                  73 passed, 0 failed
python pccm/tests/test_phase5_oracle.py                  101 passed, 0 failed
python pccm/tests/test_phase5_calc_contract_validation.py 151 passed, 0 failed
python pccm/tests/test_phase5_fingerprint.py              52 passed, 0 failed
```

| Module | Step 1 final | Step 2 | Round-1 | Round-2 | Round-3 |
|---|---|---|---|---|---|
| `test_phase1_manifest_validation.py` | 10 | 10 | 10 | 10 | **10** |
| `test_phase1_structure.py` | 21 | 21 | 21 | 21 | **21** |
| `test_phase2_contract_validation.py` | 42 | 42 | 42 | 42 | **42** |
| `test_phase2_inputs.py` | 40 | 40 | 40 | 40 | **40** |
| `test_phase3_driver_contract_validation.py` | 31 | 31 | 31 | 31 | **31** |
| `test_phase3_drivers.py` | 28 | 28 | 28 | 28 | **28** |
| `test_phase3_verifier_intersection.py` | 12 | 12 | 12 | 12 | **12** |
| `test_phase4_oracle.py` | 68 | 68 | 68 | 68 | **68** |
| `test_phase4_stage_b_source.py` | 155 | 155 | 155 | 155 | **155** |
| `test_phase4_structure.py` | 43 | 43 | 43 | 43 | **43** |
| `test_phase4_structure_contract_validation.py` | 52 | 52 | 52 | 52 | **52** |
| `test_phase5_calc_contract_validation.py` | 151 | 151 | 151 | 151 | **151** |
| `test_phase5_fingerprint.py` | 52 | 52 | 52 | 52 | **52** |
| `test_phase5_numeric.py` | — | 43 | 48 | 52 | **73** |
| `test_phase5_oracle.py` | — | 75 | 85 | 96 | **101** |
| **total** | **705** | **823** | **838** | **853** | **879** |

**Every existing count is unchanged.** No Step-1 test was weakened or removed;
the 705 → 879 delta is exactly the 174 new Step-2 tests. The 151 Step-1 contract
tests are unchanged in count; only the conditioning-term expectations inside them
moved, for erratum C1. Round 3 added 21 numerical and 5 oracle tests and removed
none. Stage-A verification is unchanged at 181/181 because Step 2 emits nothing.

**Two round-2 tests were rewritten, not weakened.** The structural guards
`test_e_is_accumulated_independently_and_not_derived_from_c_and_d` and
`test_b_is_accumulated_independently_and_not_derived_from_c_and_a` inspect the
shape of `_accumulate_totals`, and Erratum C2 changed that shape from a running
`SafeAccumulate` total to a per-measure contribution list summed by
`SafeSignedSum`. Each guard now checks **both halves** — that the sum is over the
measure's own list *and* that the list is appended to from the driver pass — so
each is strictly stronger than before, not looser. Both still fail when the
derivation `E = C + D` is substituted.

---

## 14. Negative controls

Each sabotage was applied to a working copy, the suite was run, and the source
restored.

| Sabotage | Result |
|---|---|
| Beta-PERT mean written in the forbidden naive form | **3 failed** |
| `safe_product` tier 2 removed (left-to-right only) | **3 failed** |
| discount factors built as `1/(1+r)**(t-1)` instead of iteratively | **2 failed** |
| underflow-to-zero check removed from `safe_divide` | **1 failed** |
| referenced-only FX widened to validate every row | **3 failed** |
| blank profiling weight treated as zero | **2 failed** |
| `E` derived as `C + D` instead of accumulated independently | **1 failed** *(see below)* |
| `B` derived as `C - A` instead of accumulated independently | **13 failed** |

### One control initially escaped, and the gap is closed

Substituting `E = C + D` left the **entire suite passing**. That is not an
accident of coverage: the derivation produces the same number as independent
accumulation for every ordinary model — which is exactly why I2 is worth checking
— so **no golden value can distinguish the two implementations.** What the
derivation destroys is the *meaning* of I2: an identity computed by definition can
never fail, and stops being a check at all.

Two structural guards were added, in the style of Step 1's reducer-shape test:
`_accumulate_totals` must build `e_nom` / `e_pv` with `safe_accumulate` over the
drivers and must never assign them from `c_*` or `d_*`, and the same rule is
enforced for `b_*` against `a_*` and `c_*`. A behavioural companion asserts each
driver's audit column carries that driver's own `(mean - central) * Qty * K`, so a
totals-level derivation would leave it wrong.

Both sabotages are now caught. The `B = C - A` case was caught anyway, by thirteen
tests, because B is also an audit column — only E was invisible.

---

## 15. What review should look at

1. `calc_numeric.py` — is the observable contract of the primitives the one VBA
   can honour, and is the two-tier product strategy acceptable given §12.3?
2. `calc_oracle.py` layer 1 — is the referenced-only rule genuinely built from the
   drivers first, and is any Config row validated that should not be?
3. `_accumulate_totals` — are A, B, C, D and E really five independent passes?
4. **§12.4** — the locked I1/I2 conditioning finding. This is the one item that may
   need a decision rather than an acknowledgement.
5. The golden tests — is any expected value derived from the code under test?

---

## 16. Independent review round 1 — four blockers and a portability defect

All five reproduced before fixing, all five fixed.

### 16.1 BLOCKER — conditioning must use underlying contributions (plan erratum C1)

The first submission reported that the locked I1 scale `max(1, |A|+|B|+|C|)` does
not cover cross-driver cancellation, and pinned the behaviour. **Review went
further and proved the annual scales fail the same way** — which the first
submission had explicitly claimed they did not.

| | Reproducer | Residue | Scale collapsed to | Verdict before |
|---|---|---|---|---|
| headline | three cost lines, two exact mirrors, ML `1e17` | 16 SAR | `1e-6` floor | I1 **false failure** |
| annual | `1e16` / `1` / `-1e16`, profiles `100/0`, `0/100`, `100/0` | 1 SAR | `1e-6` floor | I3a **false failure** |

Both are already-cancelled numbers: the headline totals cancel across drivers,
and the annual row aggregate cancels *within* the year — `Σ_y |annual aggregate|`
is `1` where the annual arithmetic processed `2e16`.

**Corrected.** Every scale now sums the scaled absolute magnitudes of the
**underlying per-driver and per-driver-per-year contributions**, captured during
accumulation by `ReconciliationMagnitudes` so they cannot describe a different
calculation from the one being checked. Nominal and PV each condition on their own
basis. `Σ_y |annual aggregate_y|` is not used anywhere.

The magnitudes are stored already multiplied by the relative coefficient, so the
raw contribution sum — which can exceed Double where the tolerance cannot — is
never formed. `reconcile` refuses magnitudes captured at a different coefficient
rather than silently mixing scales.

**No tolerance number changed**: `1e-9`, `1e-6`, `1e-12`, `1`. Plan §15 and the
`conditioning_terms` names in `spec/calc_contract.yaml` carry the correction;
Step-1 §14 records the narrow reopening.

### 16.2 BLOCKER — row order changed numerical results

`row position is not data` was true of the inputs and false of the arithmetic:
the accumulators ran in incoming tuple order, and floating-point addition is not
associative. Reproduced exactly as reported —

```
CL-001 = 1e16, CL-002 = 1, CL-003 = -1e16
order (001, 002, 003)  ->  A = C = E = 0
order (001, 003, 002)  ->  A = C = E = 1
```

Row order is **excluded from the calculation fingerprint**, so two workbooks with
the same fingerprint could disagree on the answer because someone sorted a
ListObject.

**Corrected.** `canonical_order` sorts drivers by ascending permanent ID, ordinal
on UTF-16 code units, using `utf16_sort_key` **imported from `calc_fingerprint`**
rather than reimplemented — two copies of an ordering rule are two chances to
disagree. It is applied to cost and risk calculation, A–E accumulation, annual
contribution accumulation, audit output order, and reference-set discovery (whose
order is observable through the inflation-factor audit rows).

This is **not** magnitude ordering. Magnitude-balanced evaluation remains only
`safe_product`'s tier-2 rescue for a single short product.

Tested across **all six permutations**, comparing the **complete**
`CalculationResult` — totals, annual rows, driver audit tuple, inflation-factor
tuple, resolved FX, discount factors and conditioning magnitudes. A companion test
asserts the fixture really is order-sensitive when summed naively, so the
permutation test cannot pass vacuously.

### 16.3 BLOCKER — `identity_allowance` did not implement its own formula

It computed `c*scale_floor + Σ(c*|term|)` where the contract says
`max(floor, c * max(scale_floor, Σ|terms|))`. Reproduced: `[1e6]` returned
`1.000001e-6` where the locked value is `1e-6`.

**Corrected** to `allowance_from_scaled`: `max(scaled_floor, scaled_terms)`, then
`max(absolute_floor, …)`. Still overflow-safe — the coefficient is still
distributed across the terms, never applied to their sum. The test that pinned the
additive form is replaced, and boundary tests were added around `scale = 1e6`,
where the relative allowance and the absolute floor coincide exactly.

### 16.4 BLOCKER — a huge Python integer leaked a raw `OverflowError`

`is_usable_double(10**400)` raised `OverflowError`, because `float()` of an
arbitrary-precision `int` raises rather than returning `inf`. A predicate that
raises is not a predicate, and the escape bypassed the whole failure contract.

**Corrected**: the conversion is guarded, `is_usable_double(10**400)` is `False`,
and a huge integer anywhere in the model — Quantity, FX rate, discount rate, Min,
a profile weight, an inflation rate, Probability — produces a structured refusal
naming the subject. Seven such inputs are asserted.

### 16.5 Reconciliation now uses the safe primitives

`A + B` and `left - right` were raw expressions, a second unchecked arithmetic
path beside the kernel's. They now go through `safe_add` / `safe_subtract` /
`safe_accumulate`, so an unrepresentable reconciliation surfaces through the
established controlled mechanism instead of putting `inf` into an `IdentityCheck`.

### 16.6 Test-portability defect — the fresh-interpreter boundary test

Review's environment preloaded `numpy`, `random` and `secrets` through site
startup *before* the oracle was imported, so the test failed on an environment
fact rather than an oracle dependency. It conflated "present before the import"
with "loaded because of the import".

**Corrected** by launching the subprocess with **`python -S`**, which disables
site processing entirely. Verified both ways: with a `sitecustomize.py` that
imports all three, a plain interpreter reports **85 preloaded modules** and `-S`
reports **none**, and the full suite passes under that environment. `-S` is also
the stronger claim — the pure modules run on a bare interpreter.

The AST import and reference tests are unchanged and were not weakened; a sabotage
adding `import random` to the oracle still fails both tests.

### 16.7 The pinned false-failure test is gone

`test_finding_headline_cancellation_can_exceed_the_locked_i1_allowance` is
removed. It asserted `valid model -> I1 fails`, which was honest evidence before
review decided the correction and is wrong to keep afterwards. Two regressions
replace it — `test_erratum_c1_headline_cross_driver_cancellation_reconciles` and
`test_erratum_c1_annual_within_year_cancellation_reconciles` — each asserting that
the same model now calculates, that the rounding residue is still real and
non-zero, that the allowance exceeds it, and that `assert_reconciled` passes.

### 16.8 Regression controls

Each of these fails against the previous Step-2 package:

| Control | Before | After |
|---|---|---|
| adversarial reorder `1e16 / 1 / -1e16`, six permutations | two distinct results | one result |
| headline cross-driver cancellation | I1 false failure | reconciles |
| annual within-year cancellation | I3a false failure, 1 SAR | reconciles |
| `identity_allowance([1e6])` | `1.000001e-6` | `1e-6` |
| `is_usable_double(10**400)` | raw `OverflowError` | `False` |
| boundary test under a preloading environment | fails | passes |

---

## 17. Independent review round 2 — three blockers

All three reproduced before fixing, all three fixed. **No design changed**: these
are implementation corrections to already-locked semantics, and neither
`docs/phase5_plan.md` nor `spec/calc_contract.yaml` needed to move.

### 17.1 BLOCKER — conditioning metadata could refuse a representable model

Erratum C1's contribution-level conditioning multiplies each contribution by the
relative coefficient before accumulating, which correctly avoids overflow. It did
so through the general-purpose `safe_multiply`, whose contract refuses a non-zero
product that rounds to exactly zero.

That rule is right for economic values and factors, and **too strong for internal
tolerance-scaling metadata.** Reproduced:

```
identity_allowance([2e-312], 1e-6, 1e-12, 1)
  -> NumericalRangeRefusal: multiplication of two non-zero values
     underflowed to exactly zero
```

and, on a fully valid model — SAR, FX 1, no inflation, zero discount, Uniform
`Min = Max = 2e-312`, Quantity 1, profile 100% — the whole calculation was refused
at `totals, driver 'CL-001': |A| nominal`. The economic outputs are perfectly
representable; only the bookkeeping was not.

The locked allowance for that input is unambiguous:

```
conditioning scale = max(1, 2e-312) = 1
relative allowance = 1e-12
final allowance    = max(1e-6, 1e-12) = 1e-6
```

**Fixed** by giving the conditioning accumulator its own underflow policy, scoped
to it alone. `safe_multiply` is **not** weakened. The justification is quantitative
rather than a convenience:

- `coefficient × |term|` only underflows when `|term|` is below about `5e-312`;
- the conditioning scale has a floor of 1, so the relative allowance is at least
  `1e-12`;
- a dropped term is therefore at most about `5e-324` against a quantity of at
  least `1e-12` — over three hundred orders of magnitude below what it would have
  to move. Even mixing huge and vanishing terms, the dropped amount is far under
  one ulp of the sum.

Overflow is still refused: a conditioning scale beyond Double makes the allowance
itself unrepresentable, and comparing against it would be meaningless.

Tested: `identity_allowance([2e-312]) == 1e-6`; the full `2e-312` model calculates
and reconciles; a mixed `[1e18, 2e-312, 5e-324]` scale equals the `[1e18]` scale
exactly; conditioning overflow still refuses; and **ordinary economic underflow is
still refused**, both in the primitives and end-to-end in the oracle, so the
exception cannot silently widen.

### 17.2 BLOCKER — driver reference fields leaked raw `AttributeError`

Canonical reference-set discovery and canonical driver ordering both compare
UTF-16 code units, so they call `.encode` — and they ran *before* the resolution
layer could refuse a non-text field. Reproduced for `currency` and
`inflation_profile` with `None`, `123` and `True`, plus a raw `TypeError` for an
unhashable `distribution`. A required user-editable field must never surface as a
Python implementation error.

**Fixed** by validating the text reference fields at the plain-data boundary,
*before* any canonicalisation. `Currency`, `Inflation Profile` and `Distribution`
must be non-blank text; `None`, booleans, numbers, empty and whitespace-only
strings all produce a `ModelInputRefusal` naming the permanent ID, the field and
the offending value or blank condition.

**Identifiers are used exactly as entered and never repaired.** `" USD "` is not
trimmed into `"USD"`: it resolves only if the FX table carries that exact key, and
is otherwise refused. Rewriting user data to make a lookup succeed would be
inventing an answer.

**One deliberate addition, flagged rather than hidden.** The permanent ID is also
checked for being non-blank text, because the canonical ordering cannot compare a
`None`. That is **not** Phase-4 permanent-ID validation: the `CL-`/`R-` prefixes,
the pattern and the counter rules remain owned by Phase 4 and are neither
re-implemented nor re-checked here. Without it, a non-text ID would still escape
as `AttributeError`.

Tested across Cost Lines and Risks, three fields × nine invalid forms, asserting a
structured refusal every time and that no raw `AttributeError`, `TypeError` or
`KeyError` escapes.

### 17.3 BLOCKER — `resolved_fx` did not match the `tblCalcFX` row rule

`resolve_fx` seeded its result with `{"SAR": 1.0}` unconditionally, conflating
*validated globally* with *belongs in the resolved referenced set*. Reproduced:

| Model | Was | Locked row rule |
|---|---|---|
| USD-only | `{"SAR": 1.0, "USD": 3.75}` | `{"USD": 3.75}` |
| empty drivers | `{"SAR": 1.0}` | `{}` |

**Fixed** by separating the two questions explicitly. The global SAR invariant is
checked exactly as before — missing, duplicated or `≠ 1` all still refuse, in
every model including one with no drivers and one referencing no foreign currency
— and the returned mapping is then built from the referenced set alone, in
canonical order.

No per-driver FX value and no analytical result changed. What changed is that the
resolution output now matches `"one row per referenced currency"`, before Step 3
starts consuming the oracle as an emission authority.

Tested: exact key sets for empty, SAR-only, USD-only, SAR+USD and AED+SAR+USD;
the rates themselves; **canonical key iteration order**, asserted as a list and
proven independent of driver arrival order; and every existing
referenced/unreferenced refusal test retained.

### 17.4 Provenance

The review archive carried no git metadata, so a reviewer could not tell which
commit it came from. `tools/package_review.py` now writes a `PROVENANCE.txt` into
the archive recording the commit, the tree hash and the file count. It is the only
entry that is not a repository blob; every other entry remains byte-identical to
`git cat-file blob`.

### 17.5 Negative controls

All ten new tests were run against the previous package, commit `e947fcd`:

| Control | vs `e947fcd` |
|---|---|
| `identity_allowance([2e-312]) == 1e-6` | **fails** |
| a vanishing term cannot move a scale that matters | **fails** |
| the `2e-312` model calculates and reconciles | **fails** |
| every invalid reference field is a structured refusal | **fails** |
| the refusal names the field and the value | **fails** |
| a blank reference field says so | **fails** |
| an identifier is used exactly as entered | **fails** |
| a non-text permanent ID is refused before ordering | **fails** |
| `resolved_fx` is exactly the referenced set | **fails** |
| `resolved_fx` iterates in canonical order | **fails** |

**10 failed, 138 passed** against `e947fcd`; all pass here. The controls proving
the exception did not widen — economic underflow still refused, conditioning
overflow still refused — pass against both, which is correct: they assert
behaviour that must not have changed.

---

## 18. Independent review round 3 — one numerical correctness class

Round 3 confirmed the three round-2 blockers fixed and package provenance
correct, and raised **one remaining class**:

> A mathematically representable final Double is still sometimes **REFUSED**
> because an intermediate operation overflows or underflows.

That is exactly the objective plan §19.2 states, so the finding is not a new
requirement — it is the accepted requirement not being met. It is recorded as
narrow plan **Erratum C2** (§0 of `docs/phase5_plan.md`), and **no broad revision
was created**.

### 18.1 The three reproducers, before the patch

All three were reproduced against the round-2 commit `c0eb409` before any code
was changed.

| | Model | Refusal at `c0eb409` | The answer that exists |
|---|---|---|---|
| **A** | three cost lines at `+MAX`, `+MAX`, `−MAX`; degenerate Uniform, SAR, FX 1, inflation 1, discount 0, Qty 1, profile 100% | `NumericalRangeRefusal: totals, driver 'CL-002': A nom` | `A_nom = C_nom = E_nom = MAX_DOUBLE` |
| **B** | one cost line, five-year profile `[MAX, MAX, −MAX, −MAX, 1]` — sum exactly `1` | `NumericalRangeRefusal: profiling for driver 'CL-001', project year 2` | the profile **is** 100%; `Knom = Kpv = 1` |
| **C** | `triangular(MAX, MAX, MAX)`; `beta_pert(MAX, MAX, MAX)`; `midpoint(5e-324, 1e-323)` | refusal; `1.7976931348623155e308` (two ulps low); refusal | `MAX`; `MAX`; `1e-323` |

Reproducer B is **not** solved by inventing a profile-weight positivity rule. A
negative weight remains legal; what changed is that the sum being validated is
computed so it can be validated at all.

### 18.2 What did NOT change

`spec/calc_contract.yaml` is byte-identical, and **no contract change was
required**. So are `calc_loader.py`, `calc_fingerprint.py`, `build_stage_a.py`,
`src/`, `bootstrap/`, `readiness/` and `tools/`. Erratum C1 and its
contribution-level conditioning, the conditioning terms and coefficients, every
tolerance number, canonical permanent-ID ordering, referenced-only FX and
inflation resolution, `resolved_fx` semantics, reference-field validation, the
two-tier `safe_product`, the discount and inflation underflow rules, D1–D6, the
A/B/C/D/E and annual definitions, the fingerprint and the Stage-A emitter are all
untouched.

**Ordinary models are bit-for-bit unchanged.** That is asserted directly, not
assumed: `test_ordinary_models_are_bit_for_bit_unchanged_by_the_signed_sum`
recomputes a five-driver mixed model's A, C, D and E totals with plain
left-to-right `+` over the same contributions in the same order and requires
exact equality.

### 18.3 Where the rules apply

| Sum | Site |
|---|---|
| profile-weight sum validation | `_resolve_weights` |
| `Knom`, `Kpv` | `precomputed_factors` |
| A, B, C, D, E | `_accumulate_totals` |
| annual Base Cost, Expected Risk, Total (nominal and PV) | `_annual_series` |
| annual-to-headline reconciliation (I3a–I3c, I4a–I4c), I5 profile sums | `reconcile` |

| Statistic | Site |
|---|---|
| Triangular mean, Beta-PERT mean, Uniform midpoint | `triangular_mean`, `beta_pert_mean`, `midpoint` |

The normal calculation is **not** magnitude-sorted anywhere.

### 18.4 `SafeSignedSum` — the algorithm later VBA must implement

Two tiers. Deterministic, integer-and-Double only, no library call.

**Tier 1 — the canonical order, unchanged.**

```
Total = 0
For i = 0 To n-1
    Total = SafeAccumulate(Total, Term(i))    ' refuses, naming Term(i)
Next
Return Total
```

If tier 1 returns, **that is the result, bit for bit**. A sum that already works
is never reordered. Tier 2 is entered **only** when tier 1 raised a
`NumericalRangeRefusal` — that is, only on addition overflow. Underflow never
enters tier 2: a sum has no underflow-to-zero rule, because a sum reaching zero is
genuine cancellation and a real answer.

**Tier 2 — deterministic opposite-sign cancellation.**

1. Validate every term as a usable finite Double (already done by tier 1).
2. Split the terms into two lists of **magnitudes**, `Pos` and `Neg`, each entry
   carrying its **original canonical index**. Exact zeros go in neither list.
3. Sort each list ascending by `(magnitude, canonical index)`. The index is the
   tie-breaker, so equal magnitudes have exactly one ordering and the result does
   not depend on sort stability.
4. While both lists are non-empty, remove the **largest** entry from each:
   * equal magnitudes → they annihilate exactly, discard both;
   * otherwise → the residual `larger − smaller` is re-inserted, in sorted
     position, on the side that was larger, keeping that side's canonical index.

   This step **cannot re-create the overflow being rescued**: it subtracts two
   non-negative magnitudes, and the residual never exceeds the larger operand. It
   terminates because each pass removes at least one entry.
5. One list is now empty. Accumulate the survivors **smallest magnitude first**
   with `SafeAccumulate`. If *that* overflows, the true signed total genuinely
   exceeds Double range and the refusal is correct.
6. Apply the surviving sign. If both lists emptied, return `+0`, never `−0`.

**The locked vectors.**

```
SafeSignedSum(MAX, MAX, −MAX)                 = MAX
SafeSignedSum(MAX, MAX, −MAX, −MAX, 1)        = 1
SafeSignedSum(MAX, MAX, −MAX, −MAX, 5e−324)   = 5e−324
SafeSignedSum(MAX, MAX)                       -> NumericalRangeRefusal
```

`math.fsum` is **not** the production contract and is not used: it has no
equivalent deterministic VBA form.

### 18.5 The convex statistics — three tiers

Each of the three statistics is a **convex combination**: positive weights summing
to exactly 1, so the true value always lies between `Min` and `Max` and always has
a representable answer.

**Tier 0 — the degenerate invariant, EXACT.** If every point is the same number,
return that number with no arithmetic at all.

```
If Min = ML And ML = Max Then Return Min      ' Triangular, Beta-PERT
If Min = Max Then Return Min                  ' Uniform
```

This is not an optimisation. `x/3 + x/3 + x/3 ≠ x` for many subnormal `x`, and
`x/2` is `0` for the smallest one, so the stable form drifts or refuses on a
distribution that has no uncertainty at all. Held across `0`, negatives,
`±MAX_DOUBLE`, `1e-320` and `5e-324`.

**Tier 1 — the accepted stable form of §19.2, unchanged.**
`Min/3 + ML/3 + Max/3`, `Min/6 + ML·(2/3) + Max/6`, `Min/2 + Max/2`, each division
before accumulation and each step through the safe primitives. Every ordinary
model lands here and its bits do not move.

**Tier 2 — binade rescue**, entered only when tier 1 raised.

1. `Biggest = Max(|point|)` over the points.
2. Count the power-of-two shift that brings `Biggest` into `[1, 2)`:
   ```
   Shifts = 0
   Do While Biggest >= 2 : Biggest = Biggest / 2 : Shifts = Shifts + 1 : Loop
   Do While Biggest <  1 : Biggest = Biggest * 2 : Shifts = Shifts - 1 : Loop
   ```
3. Apply the same shift to every point, one halving or doubling at a time.
   Halving and doubling by 2 are **exact** in IEEE-754 — they move the exponent
   and leave the significand alone — so the scaling introduces no error. Scaling
   down can flush a point more than `2^1074` times smaller than the largest to
   zero; that point cannot change any bit of a convex combination of them.
4. Evaluate the **straightforward** form on the scaled points — `(Min + ML +
   Max)/3`, `(Min + 4·ML + Max)/6`, `(Min + Max)/2` — with plain operators.
   Divide-first exists only to keep a numerator inside Double range, which in this
   binade is not in question (the numerator is bounded by 12), and dividing first
   here would cost real accuracy on nearly-cancelling points. Plain operators
   rather than the refusing primitives, because a scaled point can be subnormal
   and `subnormal / 3` may round to zero while contributing far below the last bit
   of the answer.
5. Scale the result back by the same shift.
   * **Doubling** (the overflow direction) is exact; if it overflows, the true
     statistic really does exceed Double range and the refusal is correct.
   * **Halving** (the underflow direction) is exact only while the value stays
     normal. Halve one step at a time while `|result| / 2 >= MIN_NORMAL_DOUBLE`,
     then perform **every remaining step as one division by `2^remaining`**, which
     rounds once. The exact loop cannot stop above `2^-1021`, so at most ~53 steps
     can remain before the true answer is below half the smallest subnormal, and
     the single divisor is always a small exact power of two. **Repeated halving
     through the subnormal range rounds twice and lands a unit low** — enough to
     turn `5e-324` into `0` and a value into a spurious refusal. If the single
     division reaches exactly zero, the statistic genuinely has no usable non-zero
     Double and the refusal is correct.

No `frexp`, `ldexp`, `fsum`, `Decimal`, `Fraction` or `nextafter` appears in
`calc_numeric.py`; a static AST test enforces that, because each would be correct
in Python and untranslatable to VBA.

### 18.6 Accuracy actually achieved

Swept over the §11 boundary corpus — `MAX_DOUBLE`, `nextafter(MAX_DOUBLE, 0)`,
`1e308`, `1`, `1e-300`, `1e-320`, `5e-324`, `0` and the negatives of each — against
an exact `Fraction` oracle used **only in the tests**:

| | Inputs reaching the rescue | Spurious refusals | Worst deviation from correctly rounded |
|---|---|---|---|
| Uniform midpoint | 50 | 0 | **0 ulp** |
| Triangular mean | 1090 | 0 | **0 ulp** |
| Beta-PERT mean | 826 | 0 | **1 ulp** |

The Beta-PERT ulp is the same class the accepted stable form already carries —
`(80 + 4·100 + 150)/6` is exactly `105.0` while the mandated
`Min/6 + ML·(2/3) + Max/6` is `104.99999999999999` — and is recorded, not hidden.
Its numerator needs two roundings where the other two need one.

Every refusal that remains on the corpus is a case whose exact statistic rounds to
zero — `midpoint(5e-324, 0)` is `2.47e-324`, below half the smallest subnormal —
which is the existing §19.3 underflow contract, not a C2 case.

### 18.7 Round-3 negative controls

Each sabotage was applied to a working copy, the suite was run, and the source
restored.

| Sabotage | Result |
|---|---|
| tier 2 of `safe_signed_sum` replaced by a plain left-to-right re-run | **8 failed** — all three locked vectors, the `Fraction` agreement test, and reproducers A and B |
| tier 0 (degenerate invariant) removed | **2 failed** — `triangular_mean(1e-320, 1e-320, 1e-320)` returns `1.0005e-320` |
| the binade rescue removed | **4 failed** — every subnormal and `MAX_DOUBLE` statistic refuses again |
| single-rounding scale-back replaced by repeated halving | **3 failed** — `midpoint(5e-324, 1e-323)` refuses with "underflowed to exactly zero" |

The last of these is the one worth noting: the naive scale-back passes every
*overflow* vector and fails only in the subnormal tail, which is why the corpus
sweep against an exact oracle is in the suite rather than a handful of literals.

Positive controls in the suite guard the other direction — that the rescue does
not fire when it should not, and does not change answers that already existed:

* `test_a_sum_that_already_succeeds_is_never_reordered` — over a corpus including
  `[1e16, 1.0, -1e16]`, where left-to-right gives `0.0` and a reordering would
  give `1.0`. `0.0` is the required answer.
* `test_a_signed_sum_that_genuinely_exceeds_double_range_is_still_refused`.
* `test_a_statistic_with_no_usable_non_zero_double_is_still_refused`.
* `test_the_cancelling_profile_is_still_refused_when_a_year_is_unrepresentable` —
  the same five-year profile with a unit cost of 100 makes project year 1 cost
  `100 × MAX`, and that is refused, naming the year.

---

## 19. Next step — NOT started

Step 3, whatever review scopes it to be. Nothing beyond §1 of this document has
been written.

> **NO VBA WAS IMPLEMENTED.**
> **NO PHASE-5 WORKBOOK BLOCK WAS EMITTED.**
> **NO STAGE-A EMITTER WAS CHANGED.**
> **NO `phase5_cases.json` WAS EMITTED.**
> **NO WINDOWS HARNESS WAS MODIFIED.**
> **NO WINDOWS TEST WAS RUN.**
> **STEP 3 HAS NOT BEGUN.**
> **PHASE 6 HAS NOT BEGUN.**

**PHASE 5 GATE A STEP 2 REPRESENTABLE-RESULT PATCH READY FOR INDEPENDENT REVIEW**
