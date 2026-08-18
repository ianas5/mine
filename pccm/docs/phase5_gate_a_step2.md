# Phase 5 — Gate A — Step 2: pure analytical oracle and safe numerical semantics

**Status: CORRECTED FOUR TIMES after independent review — ready for re-review.**

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
independent reimplementation must follow are specified in §19.

**Round 4** (§19) — three blocking numerical defects, all of them in the ROUND-3
RESCUES rather than in tier 1. Each round-3 rescue re-associated Double
operations, which is a heuristic and not a proof: the signed-sum rescue discarded
the rounding residual of its own cancellation and got a 100% error, the product
rescue neither detected an out-of-range exact product nor recovered a
representable one, and a convex statistic could evaluate to zero without raising
while its exact value was non-zero. All three are replaced by one exact
`Double`-and-limb kernel that computes the exact mathematical value of the
already-converted inputs, classifies its range exactly, and rounds once. §19.3
specifies it for VBA.

**No business rule, formula or tolerance NUMBER changed in any round.**
Round 2 changed **no design at all** — `docs/phase5_plan.md`,
`spec/calc_contract.yaml`, `calc_loader.py`, `calc_fingerprint.py`,
`build_stage_a.py`, `src/` and `bootstrap/` are byte-identical to the round-1
package. Erratum C1, the canonical ordering, the A/B/C/D/E and annual definitions,
the safe-product strategy and the stable distribution formulas are untouched.
Round 3 changed two production modules and the two Step-2 test modules, added
plan Erratum C2, and changed **no contract**: `spec/calc_contract.yaml`,
`calc_loader.py`, `calc_fingerprint.py`, `build_stage_a.py`, `src/`, `bootstrap/`
and `tools/` are byte-identical to the round-2 package. Round 4 changed
`calc_numeric.py` and the two test modules only — `calc_oracle.py` needed no
integration change, and the same contract and Phase-4 files remain byte-identical.
**Tier 1 of every path is untouched in round 4**, so no ordinary model's numbers
move.

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
2. **Only if tier 1 fails**, the EXACT product is formed and rounded once.

> **Superseded by round 4 (§19).** Tier 2 was originally a magnitude-balanced
> reordering: start at `1.0`, take the smallest remaining magnitude while the
> running product is `>= 1`, the largest otherwise. That is a re-association, and
> §19.1 shows it proves nothing — it returned `MAX_DOUBLE` for a product that
> genuinely exceeds `MAX_DOUBLE`, and refused one whose exact product is
> `5e-324`. Tier 2 is now the exact kernel of §19.3.

`1e308 * 10 * 0.01` overflows at the first step left to right; tier 2 answers
`1e307`, which is the correct, representable result. A genuinely unrepresentable
product such as `1e308 * 10` is still refused — and now so is one that exceeds
`MAX_DOUBLE` by less than half an ulp, which the reordering could not detect.

**A later VBA implementation must reproduce both tiers, in this order, to match.**

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

### 12.3 Reordering a product can move the last ulp — and worse

Tier 2 of `safe_product` was originally a reordering, which is not
associative-equivalent to tier 1. Round 4 proved the problem is larger than a last
ulp: a reordering also mis-classifies the range, accepting an out-of-range product
and refusing a representable one. Tier 2 is now the exact kernel of §19.3, which
has no ordering to get wrong. Tier 1 is unchanged and still runs first.
**Carried forward as a VBA requirement** (§7, §19.3).

### 12.4 The locked conditioning scale did not meet its own objective

Reported in the first Step-2 submission as a finding about the accepted plan
rather than about this code, and **now corrected by review as plan §15 erratum
C1**. Review additionally proved the same defect in the annual identities, which
the first submission had believed immune. §16.1 carries the full account.

---

## 13. Exact test counts

Run from a clean extraction, Linux, Python 3.11.

```
python -m pytest pccm/tests/ -q        897 passed, 0 failed
python pccm/builder/build_stage_a.py   181 passed, 0 failed
```

Standalone:

```
python pccm/tests/test_phase5_numeric.py                  88 passed, 0 failed
python pccm/tests/test_phase5_oracle.py                  104 passed, 0 failed
python pccm/tests/test_phase5_calc_contract_validation.py 151 passed, 0 failed
python pccm/tests/test_phase5_fingerprint.py              52 passed, 0 failed
```

| Module | Step 1 final | Step 2 | Round-1 | Round-2 | Round-3 | Round-4 |
|---|---|---|---|---|---|---|
| `test_phase1_manifest_validation.py` | 10 | 10 | 10 | 10 | 10 | **10** |
| `test_phase1_structure.py` | 21 | 21 | 21 | 21 | 21 | **21** |
| `test_phase2_contract_validation.py` | 42 | 42 | 42 | 42 | 42 | **42** |
| `test_phase2_inputs.py` | 40 | 40 | 40 | 40 | 40 | **40** |
| `test_phase3_driver_contract_validation.py` | 31 | 31 | 31 | 31 | 31 | **31** |
| `test_phase3_drivers.py` | 28 | 28 | 28 | 28 | 28 | **28** |
| `test_phase3_verifier_intersection.py` | 12 | 12 | 12 | 12 | 12 | **12** |
| `test_phase4_oracle.py` | 68 | 68 | 68 | 68 | 68 | **68** |
| `test_phase4_stage_b_source.py` | 155 | 155 | 155 | 155 | 155 | **155** |
| `test_phase4_structure.py` | 43 | 43 | 43 | 43 | 43 | **43** |
| `test_phase4_structure_contract_validation.py` | 52 | 52 | 52 | 52 | 52 | **52** |
| `test_phase5_calc_contract_validation.py` | 151 | 151 | 151 | 151 | 151 | **151** |
| `test_phase5_fingerprint.py` | 52 | 52 | 52 | 52 | 52 | **52** |
| `test_phase5_numeric.py` | — | 43 | 48 | 52 | 73 | **88** |
| `test_phase5_oracle.py` | — | 75 | 85 | 96 | 101 | **104** |
| **total** | **705** | **823** | **838** | **853** | **879** | **897** |

**Every existing count is unchanged.** No Step-1 test was weakened or removed;
the 705 → 879 delta is exactly the 174 new Step-2 tests. The 151 Step-1 contract
tests are unchanged in count; only the conditioning-term expectations inside them
moved, for erratum C1. Round 3 added 21 numerical and 5 oracle tests and removed
none. Stage-A verification is unchanged at 181/181 because Step 2 emits nothing.

Round 4 added 15 numerical and 3 oracle tests and removed none. Four round-3
numerical tests were RETARGETED rather than weakened, because the mechanism they
named no longer exists: `test_the_balanced_order_is_deterministic` and
`test_product_signs_are_preserved_through_the_balanced_order` became
`..._the_product_rescue_...` and each gained an assertion the old heuristic fails;
`test_the_binade_rescue_uses_only_powers_of_two` became
`test_the_exact_rescues_use_only_vba_translatable_operations` with a longer ban
list and a new ban on calling Python's `int`; and the boundary-corpus rounding
sweep tightened from "Beta-PERT may be one ulp out" to exact equality for all
three statistics. Every one of the four is strictly stronger than the test it
replaced.

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

### 18.4 The round-3 rescues — SUPERSEDED by §19

Round 3 replaced the refusals with two rescues that were **not faithful**, and
round 4 replaced both. They are described here only so the review trail is
readable; **§19 is the specification**, and nothing in this subsection is
implemented any more.

* **Signed sum**: split the terms into positive and negative magnitudes, then
  repeatedly cancel the largest opposite-signed pair, re-inserting the residual
  `p - n` as one rounded Double subtraction.
* **Convex statistics**: scale the points by a shared power of two into `[1, 2)`,
  evaluate there, and scale the result back.
* **Product** (from round 2, and unchanged by round 3): re-evaluate in a
  magnitude-balanced order.

All three are re-associations of Double operations. §19.1 explains why that class
of rescue cannot be correct, and §19.7 keeps each of them alive as a negative
control so the module cannot quietly return to one.

### 18.5 Accuracy the round-3 rescue achieved — superseded

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

### 18.6 Round-3 negative controls

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

## 19. Independent review round 4 — faithful rescues

Round 4 confirmed the package scope, provenance, every previously requested
correction and all 879 tests, and then proved that the round-3 **rescues
themselves were not faithful**. Three blocking numerical defects, all in rescue
paths, none in tier 1.

### 19.1 Root cause — a re-association is not a proof

Every round-2/round-3 rescue re-associated Double operations: a different order,
a different grouping, the same `+` and `*`. That is a heuristic. It gets the
right answer often enough to pass friendly vectors and gives no guarantee at all.

**Signed sum.** Cancelling the largest opposite-signed pair computes `p - n` as
one *rounded* Double subtraction, and the rounding residual it discards is gone
for good. When the large terms then cancel, that residual **was** the answer:

```
[6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]

exact sum        -1e292                    (in range, representable)
round-3 rescue   -1.99792015476736e292     (about 100% wrong)
```

"the residual never exceeds the larger operand" proves only that the subtraction
does not overflow. It says nothing about preserving the signed sum, and that
distinction is what the defect turns on.

**Product.** A magnitude-balanced order decides nothing about the exact product's
range. It returned `MAX_DOUBLE` for a product that genuinely exceeds `MAX_DOUBLE`
by 0.887 ulp, and refused one whose exact product rounds to `5e-324`.

**Convex zero.** A statistic can evaluate to exactly `0` in the stable form
*without raising*, while its exact value is non-zero and has no non-zero Double.
`midpoint(-20s, 19s)` with `s = 5e-324` is `-10s + fl(9.5s)` = `-10s + 10s` = `0`;
the exact midpoint is `-0.5s`. Accepting any successful tier-1 result reports zero
for a value the model does not have — the silent deletion §19.3 of the plan exists
to prevent.

### 19.2 The criterion, stated exactly

Tier 1 is untouched in all three cases. **If the ordinary evaluation produces a
value, that value is the result, bit for bit** — canonical left-to-right for sums,
left-to-right for products, the stable form for statistics. This patch does not
reinterpret ordinary Double expressions as exact rational arithmetic; doing so
would move accepted calculations.

Only a **failed** tier 1 — plus the one extra case of a *non-degenerate statistic
whose tier-1 result is exactly zero* — enters the rescue, and the rescue is judged
against the exact mathematical value of the already-converted IEEE-754 inputs:

```
|exact| > MAX_DOUBLE                 ->  NumericalRangeRefusal
exact != 0 and it rounds to zero     ->  NumericalRangeRefusal
otherwise                            ->  the correctly rounded Double
```

The achieved result is **correct rounding**, not merely faithful rounding: round
to nearest, ties to even, one rounding, on every rescue path and every statistic.
No weaker bound is claimed and none is needed.

**The range test is on the exact value, not on the rounded one.** `MAX_DOUBLE`'s
ulp is `2**971` and the floating overflow threshold sits half an ulp above
`MAX_DOUBLE`, so any exact value in `(MAX_DOUBLE, MAX_DOUBLE + 2**970)` is
mathematically out of range while a `float()` of it is `MAX_DOUBLE` and reports no
overflow. Classifying with `math.isfinite(float(exact))` would call that
representable and hide the defect; the tests classify with
`abs(exact) <= Fraction(MAX_DOUBLE)` first, and
`test_the_exact_range_test_is_not_the_same_as_float_overflow` demonstrates the
band rather than asserting it.

### 19.3 The exact kernel — what later VBA must implement

One kernel serves both rescues and the zero classification. **Every operation in
it is a `Double` operation**: addition, subtraction, multiplication, division by
an exact power of two, comparison and truncation (`Fix`). There is no `Decimal`,
no `Fraction`, no `fsum`, no `frexp`/`ldexp`, no arbitrary-precision integer and
no platform extended precision — a static AST test
(`test_the_exact_rescues_use_only_vba_translatable_operations`) enforces every one
of those bans on `calc_numeric.py` itself.

**Representation.** A value is `(sign, magnitude, shift)`, where `magnitude` is an
array of **limbs** in base `2**24` — each limb a `Double` holding an exact integer
in `[0, 2**24)` — and `shift` is a `Long` binary exponent. The value is
`sign * (Σ limb[i] * 2**(24i)) * 2**shift`. Base `2**24` is what keeps every
intermediate exact: a limb-by-limb product is under `2**48`, and with carries every
running total stays under `2**49`, comfortably inside the `2**53` exact-integer
range of a `Double`.

**`Fix`.** The kernel's one truncation is written in `Double` operations —
`(v + 2**52) - 2**52` rounds to the nearest integer, and one comparison turns that
into truncation — so VBA's native `Fix` is a direct substitute and Python's `int`
is never involved.

**Decomposition.** `_decompose(x)` returns `(sign, mantissa, exponent)` with
`|x| = mantissa * 2**exponent` and `mantissa` an integer in `[2**52, 2**53)`. It is
a counting loop over a fixed table of `2**512, 2**256, … , 2**1`: multiply while
below `1` (which lifts subnormals, exactly), then divide while at or above each
table power. Every step is an exact power-of-two scaling, so the decomposition is
exact for every finite Double including subnormals. No `frexp`.

**Exact sum.** Every Double is an exact integer multiple of `2**smallest`, where
`smallest` is the least of the terms' own exponents, so aligning all terms there is
exact and the sum is an exact integer in that unit. Positive magnitudes accumulate
into one array and negative magnitudes into another; the two are compared once and
the smaller subtracted from the larger. That needs no signed carries and no
tie-breaking rule at all — **addition is associative and commutative in this
representation**, so unlike a re-association the result does not depend on order.

**Exact product.** The mantissas multiply as integers (schoolbook, `O(k**2)` limb
products) and the exponents add. The product is exact however far outside Double
range it lands, which is what makes the range classification a fact rather than an
artefact of evaluation order.

**Rounding — once.** Given `(sign, magnitude, shift)`:

1. find the index `B` of the most significant set bit; an all-zero magnitude is an
   exact zero and returns `+0`;
2. `E = B + shift`. If `E > 1023` the value is at least `2**1024`: **out of
   range**;
3. the target ulp exponent is `q = max(E - 52, -1074)`, and `k = q - shift` bits
   must be dropped. If `k <= 0` the value already has at most 53 significant bits
   and no bit below `2**-1074`, so it **is** a Double and nothing is rounded;
4. otherwise `Q = floor(magnitude / 2**k)` (at most four limbs contribute, so `Q`
   is built exactly), `roundBit` is bit `k-1`, and `sticky` is "any bit below
   `k-1`";
5. **the exact overflow test.** `q = 971` is the only exponent where the answer can
   straddle `MAX_DOUBLE = (2**53 - 1) * 2**971`. There, `Q > 2**53 - 1`, or
   `Q = 2**53 - 1` with `roundBit` or `sticky` set, means the exact value exceeds
   `MAX_DOUBLE`: **out of range**, even though it would round to `MAX_DOUBLE`;
6. round to nearest, ties to even: add one to `Q` when `roundBit` is set and
   (`sticky` or `Q` is odd);
7. `Q = 0` after that means a non-zero exact value has no non-zero Double:
   **underflow refusal**;
8. the result is `sign * Q * 2**q`, applied by the same exact power-of-two table.
   `|q| <= 1074` and `Q < 2**54`, so every intermediate of that scaling is exactly
   representable.

**Dividing by 3, 6 or 2.** A convex statistic's numerator is dyadic but the
statistic is not. The numerator is shifted left by `_GUARD_BITS = 64`, divided by
the small integer with a limb-by-limb `divmod`, and the division's remainder
becomes an extra sticky flag. A non-zero remainder means the true value is
strictly above the quotient — exactly what a sticky bit encodes — so ties are
still resolved correctly and the rounding is exact.

**Determinism and ordering.** There is no ordering rule to get wrong: the exact
sum is order-independent by construction, the exact product is order-independent
by construction, and the rounding is a function of the exact value alone. Given
the same input Doubles, every implementation of the above must produce the same
bits.

**Cost.** `O(limbs)` per term for a sum and `O(limbs**2)` for a product, on a path
that only runs when the ordinary evaluation produced no value at all. Tier 1
remains a plain `O(n)` loop. Ten thousand adversarial rescue cases run in well
under a second.

### 19.4 The convex statistics — three tiers, and a classified zero

* **Tier 0 — degenerate invariant, EXACT.** Every point equal: return that point
  with no arithmetic. Unchanged from round 3, and still necessary —
  `x/3 + x/3 + x/3 != x` for many subnormal `x`.
* **Tier 1 — the accepted stable form of plan §19.2, unchanged.** A **non-zero**
  result is returned bit for bit.
* **Exact tier — reached when tier 1 raised, and when tier 1 returned exactly
  zero on a non-degenerate distribution.** The exact numerator is formed with the
  kernel (`Min + ML + Max`, `Min + 4·ML + Max` supplied as four copies of `ML` so
  it can be formed even where `4·ML` has no Double, `Min + Max`), divided by 3, 6
  or 2, and rounded once. An exactly zero numerator returns `0`; a non-zero one
  returns the correctly rounded tiny Double or refuses.

This replaces the round-3 binade rescue, and it is strictly better: Beta-PERT was
one ulp out under the binade rescue and is exact under this one.

### 19.5 The five reproducers, before and after

| | Input | Round 3 | Round 4 | Exact |
|---|---|---|---|---|
| §1.1 | `[6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292]` | `-1.99792015476736e292` | **`-1e292`** | `-1e292` |
| §1.3 | `[-8e307, -7e307, -1.78e308, 5e307, -1e292, 1e308, -MAX, 1.78e308]` | `-MAX_DOUBLE` | **`NumericalRangeRefusal`** | `|exact|` exceeds `MAX_DOUBLE` by ~0.501 ulp |
| §5.1 | `safe_product([1e50, MAX, 1e-150, 1e100])` | `MAX_DOUBLE` | **`NumericalRangeRefusal`** | exceeds `MAX_DOUBLE` by ~0.887 ulp |
| §5.2 | `safe_product([1e100, 0.5, 1e150, 5e-324, 1e-250])` | `NumericalRangeRefusal` | **`5e-324`** | `5e-324` |
| §8 | `midpoint(-20s, 19s)`, `s = 5e-324` | `0.0` | **`NumericalRangeRefusal`** | `-0.5s`, non-zero, no usable Double |

§1.2 is the same seven terms as a seven-cost-line model
(`CL-001 … CL-007`, degenerate Uniform, SAR, FX 1, inflation 1, discount 0,
Qty 1, 100% profile). `A_nom = C_nom = E_nom = -1e292`, the annual row matches,
and `assert_reconciled` passes.

### 19.6 The adversarial oracle

`Fraction.from_float` is exact and appears in **test code only**. It is not
production semantics; it is independent Gate-A evidence about what the production
algorithm should have produced. Production may not use it, and a static test
enforces that.

| Sweep | Cases | Shape | Corpus |
|---|---|---|---|
| `safe_signed_sum` | 10,000 | lengths 2–20 | `MAX_DOUBLE`, `nextafter(MAX_DOUBLE, 0)`, `1e308`, `1.78e308`, `1.7e308`, `1e292`, `1`, `1e-292`, `1e-308`, `1e-320`, `5e-324`, and negatives |
| `safe_product` | 10,000 | 2–6 factors | `MAX_DOUBLE`, `1e308`, `1e250`, `1e150`, `1e100`, `1e50`, `10`, `2`, `1`, `0.5`, `0.1`, `1e-50`, `1e-100`, `1e-150`, `1e-250`, `1e-300`, `1e-320`, `5e-324`, and negatives |
| convex statistics | full boundary corpus, plus every small subnormal multiple | 2–3 points | the §11 corpus and `k * 5e-324` for `|k| <= 8` |

The generator is a fixed linear congruential stream written into the test, so the
corpus does not depend on the stdlib RNG staying stable across Python versions and
there is **no unseeded randomness anywhere in the suite**. Only cases where tier 1
FAILED belong to the rescue oracle; each sweep asserts a floor on how many cases
actually reached the rescue, so it cannot pass vacuously (5,000+ and 3,000+
respectively).

Result: **zero spurious refusals, zero fabricated values, zero cases off the
correctly rounded Double** across all three sweeps, plus the boundary corpus of
§11 and the exhaustive subnormal grid.

### 19.7 Round-4 negative controls

Each sabotage was applied to a working copy, the suite was run, and the source
restored.

| Sabotage | Result |
|---|---|
| signed-sum rescue → the round-3 rounded `p - n` cancellation | **6 failed**, including both signed-sum reproducers, the 10,000-case sweep, and the end-to-end seven-cost-line model |
| product rescue → the round-2 magnitude-balanced heuristic | **5 failed**, including both product reproducers and the 10,000-case sweep |
| convex zero classification → blindly accept any tier-1 result | **4 failed**, including the §8 reproducer and the both-directions sweep |

Each superseded algorithm is ALSO kept inside the suite as a positive control —
`test_the_rounded_pair_cancellation_would_fail_the_residual_reproducer`,
`test_the_magnitude_balanced_order_would_fail_both_product_reproducers`,
`test_blindly_accepting_a_tier_one_zero_would_fail_the_convex_reproducer` — each
computing the old answer inline and asserting it differs from the faithful one. A
regression to any of the three cannot pass silently even if the sabotage is never
re-run by hand.

### 19.8 What reconciliation does NOT prove

`test_reconciliation_cannot_catch_a_consistently_wrong_rescue` substitutes the
round-3 rounded-pair rescue into the live oracle and calculates the seven-cost-line
model. A, C, E and the annual series all come out wrong — **and all wrong the same
way**, because they are the same algorithm applied to the same contributions.
Every identity holds. `assert_reconciled` passes. The model reports a result.

Reconciliation therefore verifies **consistency between calculation paths**. It is
not an independent numerical-accuracy oracle and cannot be made into one. Only a
fixture that compares against the exact mathematical value catches this class,
which is why the end-to-end assertion is on the calculated value and not on the
identities.

### 19.9 What did NOT change

Tier 1 of every path. `spec/calc_contract.yaml` is byte-identical and **no
contract change was required**, as are `calc_loader.py`, `calc_fingerprint.py`,
`build_stage_a.py`, `src/`, `bootstrap/`, `readiness/` and `tools/`. Erratum C1
and its contribution-level conditioning, the conditioning terms, every tolerance
number, canonical permanent-ID ordering, referenced-only FX and inflation
resolution, `resolved_fx` semantics, reference-field validation, the discount and
inflation underflow rules, D1–D6, the A/B/C/D/E and annual definitions, the
fingerprint and the Stage-A emitter are untouched. `calc_oracle.py` needed **no
integration change at all** in round 4 — the rescues live entirely below it.

---

## 20. Next step — NOT started

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

**PHASE 5 GATE A STEP 2 FAITHFUL-RESCUE PATCH READY FOR INDEPENDENT REVIEW**
