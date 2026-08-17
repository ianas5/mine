# Phase 5 — Gate A — Step 2: pure analytical oracle and safe numerical semantics

**Status: ready for independent review.**

Step 1 is accepted and closed. This step locks and tests the pure numerical
semantics that later Stage-A emission and later VBA must implement. It is
Linux-only, in-memory, and side-effect free.

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
| `builder/pccm_builder/calc_oracle.py` | **new** — data model, resolution/validation, analytical kernel, reconciliation |
| `builder/pccm_builder/calc_numeric.py` | **new** — safe arithmetic, stable statistics, factor series, tolerance, failure hierarchy |
| `tests/test_phase5_numeric.py` | **new** — 43 tests |
| `tests/test_phase5_oracle.py` | **new** — 75 tests |
| `docs/phase5_gate_a_step2.md` | **new** — this document |

**Unchanged, and verified unchanged:** `spec/calc_contract.yaml`,
`builder/pccm_builder/calc_loader.py`, `builder/pccm_builder/calc_fingerprint.py`,
`builder/pccm_builder/__init__.py`, `builder/build_stage_a.py`, all four earlier
specifications, every Phase 1–4 source and test, every `.ps1`, and
`docs/phase5_plan.md`.

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
- **row position is not data** — it is not an input to anything, and a test proves
  reordering the driver sequence changes no value at all;
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

### 12.4 The locked I1 / I2 conditioning scale does not cover cross-driver cancellation

**This one is a finding about the accepted plan, not about Step-2 code, and it is
reported rather than corrected.**

Plan §15 states the objective:

> "The scale must reflect the **magnitude of the arithmetic performed**, not the
> magnitude of its net result."

The locked annual scales achieve it — `Sum_y |annual| + |headline|` keeps growing
when signed annual values cancel. The locked I1 and I2 scales,
`max(1, |A| + |B| + |C|)` and `max(1, |C| + |D| + |E|)`, are sums of the
**headline totals**, so when the totals themselves cancel across drivers the scale
collapses to the `1e-6` floor while the accumulation error does not.

Reproducer, all inputs valid — ordered three-point sets, positive Quantity,
weights summing to 1:

```
CL-001   Min 0      ML 1e17   Max 4e17      (central 1e17, mean ~1.667e17)
CL-002   Min 10     ML 30     Max 110       (central 30,   mean 50)
CL-003   Min -4e17  ML -1e17  Max 0         (exact mirror of CL-001)
```

```
A = 32.0    B = 16.0    C = 64.0
A + B - C = -16.0        allowance = 1e-06        I1 REPORTED AS FAILING
```

One ulp at a partial sum of `1e17` is already 16 SAR, so the residue is
arithmetic, not a defect in the model or the calculation. The locked scale sizes
the tolerance by the near-zero *result* instead.

**Step 2 implements the accepted definition exactly and does not alter it.** The
behaviour is captured by
`test_finding_headline_cancellation_can_exceed_the_locked_i1_allowance`, which
asserts the current outcome so it cannot change unnoticed and **must be updated if
review amends §15**. Whether to amend it — for example by conditioning I1/I2 on
the summed absolute driver contributions, as the annual identities already do — is
a design decision for review, not for this step.

---

## 13. Exact test counts

Run from a clean extraction, Linux, Python 3.11.

```
python -m pytest pccm/tests/ -q        823 passed, 0 failed
python pccm/builder/build_stage_a.py   181 passed, 0 failed
```

Standalone:

```
python pccm/tests/test_phase5_numeric.py    43 passed, 0 failed
python pccm/tests/test_phase5_oracle.py     75 passed, 0 failed
```

| Module | Step 1 final | Step 2 |
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
| `test_phase5_numeric.py` | — | **43** |
| `test_phase5_oracle.py` | — | **75** |
| **total** | **705** | **823** |

**Every existing count is unchanged.** No Step-1 test was weakened, repointed or
removed; the 705 → 823 delta is exactly the 118 new Step-2 tests. Stage-A
verification is unchanged at 181/181 because Step 2 emits nothing.

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

## 16. Next step — NOT started

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

**PHASE 5 GATE A STEP 2 READY FOR INDEPENDENT REVIEW**
