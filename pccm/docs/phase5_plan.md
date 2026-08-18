# PCCM — Phase 5 plan: deterministic and analytical calculation engine

**Revision E.** The fingerprint mathematics is **not** reopened: the constants,
the canonical encoding, every numeric and collision vector, and the reference
digest `50B6EB0E26857EA7` are unchanged, and the review reproduced them again.
This revision fixes one **VBA-language implementation blocker** and four
consistency points: the recurrence may not be reduced with VBA's `Mod` or `\`,
whose operands are `Long`-typed while the intermediate reaches `2.8 × 10¹¹`, so a
`Double`-only reduction is locked and verified; the final commit assignment is
brought inside the transaction rather than assumed infallible; the post-failure
scalar state is described accurately; the public entry-point count is corrected to
six; and the initial `calc_state` values and the accessors' empty-case behaviour
are locked.

**DESIGN GATE. No code, no VBA, no workbook change, no build artifacts.** Phase 4
is accepted and closed; nothing in `src/`, `spec/`, `builder/`, `bootstrap/` or
`tests/` is touched by this document.

---

## 0. Revision E errata — editorial, applied in place

Revision E is **accepted** and the design gate is **closed**. Four errata were
raised during acceptance. They are **editorial and test-definition corrections**;
they change no locked design decision, no constant, no vector, no anchor, no
schema and no expected value. They are applied **in place** in the sections named
below, and recorded here so a reviewer can see exactly what moved.

| # | Correction | Applied in |
|---|---|---|
| **E1** | **Case 33 post-failure state.** Immediately after rollback and **before** failure metadata is written, `C13:C20` is restored exactly. The *final* observable state is `C13:C16` restored exactly, `C17:C20` carrying **new** failed-attempt and derived-status metadata, `C23:C32` restored exactly and all five analytical ListObjects restored exactly. The final acceptance comparison is therefore **`C13:C16` + `C23:C32` + the five analytical ListObjects**, never all of `C13:C20`. | §12.5, §23 case 33, §25.7 |
| **E2** | **Case 36 `Long` wording.** `281,320,423,161` is **not** equal to `131 × Long.MaxValue`. The exact statements are `131 × 2,147,483,647 = 281,320,357,757` and `281,320,423,161 = 131 × Long.MaxValue + 65,404`; prose may otherwise say "approximately 131 times the signed-`Long` maximum". The reduction vectors and their expected remainders are **unchanged**. | §11.5, §23 case 36 |
| **E3** | **No executable VBA proof exists on Linux.** No Gate-A acceptance wording may claim that VBA executes, or produces the fingerprint, on Linux. The proof split is locked in §21.0. Gate-A static validation of VBA source is **not** weakened by this erratum — it remains generated/source/static conformance, with **no execution**. | §11.6, §21.0, §21 |
| **E4** | **Callable terminology.** "six public Phase-5 callables in total" reads as a cap on every `Public` procedure in the phase. The correct phrase is **"six public `PCCM_` automation/API entry points"**. The six are unchanged. This does **not** prohibit numerical helper procedures from being `Public` where cross-module VBA calls require it; it prohibits **additional `PCCM_` endpoints**, none of which may be added in Phase 5 without review. | §11.13, §27, §29 |

One further **Gate-B requirement is locked here and deliberately not implemented
yet**: direct Windows/VBA vector coverage, specified in §24.1.

### Erratum C1 — raised later, by implementation

| # | Correction | Applied in |
|---|---|---|
| **C1** | **Reconciliation conditioning is on the UNDERLYING CONTRIBUTIONS**, not on the headline totals or the annual row aggregates. Gate-A Step 2 proved both of the former insufficient: each is an already-cancelled number, and each made a correct calculation on a valid model report a false internal-invariant failure. **No tolerance number changed** — only the operands. The allowance is also restated as `max(floor, coefficient × max(scale_floor, scale))`, a maximum rather than a sum. | §15, and the `conditioning_terms` names in `spec/calc_contract.yaml` |

C1 is **not** editorial. It is a numerical correction to an accepted definition
that did not satisfy its own stated objective, and it is recorded here rather than
folded silently into a new revision.

### Erratum C2 — raised later, by implementation

| # | Correction | Applied in |
|---|---|---|
| **C2** | **A representable final result is never refused because an intermediate left the range.** §19.2 states the objective — "a mathematically equivalent expression avoids an intermediate that can overflow while the final result is representable" — but the stable forms alone do not reach it for two families. **(a) Signed sums**: any sum of validated Doubles whose *final* value is representable must be produced, even where the canonical-order partial sums are not (`MAX + MAX − MAX` is `MAX`). **(b) Convex statistics**: a deterministic central value or distribution mean is a convex combination of its points, so it always lies between `Min` and `Max` and always has a representable answer; it must not be refused because an internal averaging step overflowed or underflowed, nor reported as zero when its exact value is non-zero. **(c) Products**: the same rule and the same exact criterion apply to the `safe_product` rescue. **(d) Materialization**: the boundary is a named, published Phase-5 value, not whatever subexpression the implementation assigns to a local variable. **No tolerance, weight, formula or expected value changed**, and the stable forms of the table below remain the required ordinary path. | §19.2, and `docs/phase5_gate_a_step2.md` §18–§20 for the algorithms |

**What C2 locks, precisely.**

1. **Canonical order is the calculation.** The supplied order — ascending
   permanent ID, project year ascending — is evaluated first, with the existing
   `SafeAccumulate` semantics. **If it produces a value, that value is the
   result, bit for bit.** A sum that already works is never reordered, so no
   model that calculates today produces a different number.
2. **Cancellation rescue is a second tier, and only for addition overflow.** It
   runs only where tier 1 produced no value at all. It is not triggered by
   underflow, by inaccuracy, or by anything else.
3. **The stable forms of §19.2 remain tier 1 for the convex statistics.** The
   rescue runs only where that form cannot produce a result.
4. **A distribution with zero uncertainty returns its point exactly.** `Min ==
   Max` (Uniform) gives a midpoint of exactly `Min`; `Min == ML == Max`
   (Triangular, Beta-PERT) gives a mean of exactly `Min`. This holds across the
   whole usable Double range, including `0`, negatives, `±MAX_DOUBLE` and the
   smallest subnormal. **No last-ulp drift is acceptable for a distribution with
   no uncertainty**, and the stable forms alone do not give this: `x/3 + x/3 +
   x/3 ≠ x` for many subnormal `x`, and `x/2` cannot even be formed for the
   smallest one.
5. **A genuine range failure is still a refusal.** Where the *final* result has
   no usable Double — a signed total that really exceeds the range, a statistic
   whose exact value rounds to zero — the controlled `NumericalRangeRefusal` of
   §19.1 stands unchanged. C2 removes refusals of answers that exist; it creates
   no fabricated values.
6. **No positivity rule is introduced anywhere.** A negative profile weight, a
   negative contribution and a negative total all remain legal. What changed is
   how a sum is computed, not what a model may contain.
7. **A rescue must be faithful, not merely better.** A rescue that re-associates
   Double operations is a heuristic and proves nothing. Re-associating a sum
   discards the rounding residual of each intermediate subtraction, and that
   residual can BE the answer once the large terms cancel; reordering a product by
   magnitude decides neither whether the exact product is in range nor whether a
   representable one exists. Each rescue path is therefore judged against the
   **exact mathematical value of the already-converted IEEE-754 Double inputs**:

   ```
   |exact| > MAX_DOUBLE                      -> NumericalRangeRefusal
   exact non-zero but rounding to zero       -> NumericalRangeRefusal
   otherwise                                 -> the correctly rounded Double
   ```

   The range test is on the **exact** value, not on the rounded one: a result can
   exceed `MAX_DOUBLE` by less than half an ulp and still round to it, and
   returning `MAX_DOUBLE` there fabricates a value.
8. **Cross-language reproducibility.** Both rescues are specified as deterministic
   `Double`-and-integer algorithms — exact power-of-two scaling, fixed-width limbs
   holding exact integers, comparisons, counting loops and truncation — with a
   definite order and a definite tie rule. No `math.fsum`, `frexp`, `ldexp`,
   `Decimal`, `Fraction` or arbitrary-precision integer appears in production;
   `Decimal` and `Fraction` remain legitimate in independent test oracles only.
9. **A zero is classified, not assumed.** For a NON-DEGENERATE convex statistic:

   * a **non-zero** successful tier-1 result is returned unchanged — tier 1's own
     rounding is the accepted ordinary path and moving it would change existing
     calculations;
   * a tier-1 result of **exactly zero** is **not** automatically accepted. It is
     classified as either an **exact mathematical zero**, which returns `0`, or a
     **non-zero statistic whose Double result collapses to zero**, which is a
     `NumericalRangeRefusal`.

   `midpoint(-20s, 19s)` with `s = 5e-324` evaluates to `0` in the stable form
   without raising, while the exact midpoint is `-0.5s` — non-zero, and with no
   usable non-zero Double. Reporting `0` there is the silent deletion §19.3 exists
   to prevent. The classification is a deterministic Double-portable numerical
   test, not `Decimal` or `Fraction`, and it applies to the Uniform midpoint, the
   Triangular mean and the Beta-PERT mean alike.

10. **A representability boundary sits at a NAMED, MATERIALIZED value.** A
    non-materialized implementation subexpression is **not** an independent
    representability boundary, and a rescue may carry it in the exact wide
    representation until the next named Phase-5 output is reached. Every named
    output must still be a usable Double.

    | Boundaries — each must be a Double | Not boundaries |
    |---|---|
    | resolved FX rates, annual inflation factors, discount factors | `w_y × infl_y` and `w_y × infl_y × disc_y` |
    | driver Central Value, Mean Value, `Knom`, `Kpv` | the pre-FX sum inside `Knom` / `Kpv` |
    | every published per-driver audit amount — Deterministic, Mean-Basis, Uncertainty Mean Shift, Expected Risk, nominal and PV | one Cost Line's or one Risk's contribution to one annual aggregate |
    | each of the six `tblCalcAnnual` aggregate columns | the temporary nominal per-driver/year value used before its PV discount |
    | the A/B/C/D/E headline totals | an unscaled annual contribution used only to build a C1 conditioning magnitude |
    | published reconciliation allowance and difference values | |

    **A materialized per-driver amount remains a real boundary.** If a driver's own
    Mean-Basis Nominal exceeds `MAX_DOUBLE`, the model is refused even where another
    driver would cancel it in the headline — `tblCalcDrivers` has to publish that
    row. The wider exact form exists ONLY across the non-materialized intermediates
    needed to produce one required named Double; it never turns the model into one
    arbitrary-precision final calculation.

    The same rule governs C1 conditioning. The quantity C1 needs is
    `coefficient × |contribution|`, and with the locked `1e-12` that is finite even
    where the contribution is `2 × MAX_DOUBLE`. On a rescue path the coefficient is
    folded into the same exact factor expression, so the unscaled contribution never
    has to become a Double just to record its metadata. The accepted
    conditioning-underflow rule is unchanged, and a final conditioning magnitude or
    allowance that genuinely has no Double is still a controlled refusal, never a
    silent cap.

C2 applies to the profile-weight sum validation, `Knom` and `Kpv`, the A/B/C/D/E
accumulations, the annual Base Cost / Expected Risk / Total series, the
annual-to-headline reconciliation sums, the C1 conditioning magnitudes, the
`safe_product` rescue and the three convex statistics. It does **not**
magnitude-sort or re-associate the normal calculation, and it changes nothing in
`spec/calc_contract.yaml`.

Like C1, C2 is **not** editorial: it is a numerical correction to an accepted
definition that did not reach its own stated objective.

---

## 1. Phase objective

Convert a structurally valid workbook into reproducible **SAR nominal and PV cost
measures**, analytically and deterministically, in a form the later Monte Carlo
phase reuses without touching a worksheet inside an iteration loop.

Phase 5 produces numbers. It produces no random numbers.

---

## 2. Mapping to the locked architecture roadmap

| Implementation phase | Layer | Status |
|---|---|---|
| 1–3 | Workbook skeleton, Setup/Config inputs, driver registers | closed |
| **4** | **Structural runtime** — permanent IDs, timeline application, profiling and inflation synchronisation, add/delete, rollback, Stage-B runtime harness | **closed** |
| **5** | **Calculation / `_Calc` factor engine** — FX resolution, inflation factors, discount factors, profiling weights, `Knom` / `Kpv`, deterministic base, mean-basis base, expected risk, analytical annual cash flow, calculation fingerprint | **this phase** |
| 6+ | RNG (MRG32k3a), sampling, simulation, percentiles, contingency, sensitivity, Results, Dashboard, Model Check UI | not started |

Phase 5 is **not** the RNG / Monte Carlo validation gate. "Mean" throughout means
the **analytical** expected value of a distribution, never a simulation output.

---

## 3. Scope

### In scope

Resolved FX to SAR · annual inflation factors · annual discount factors · cost and
risk profiling weights by permanent ID · `Knom` / `Kpv` per driver · Escalated
Deterministic Base Estimate (Nominal, PV) · Mean-Basis Base Cost (Nominal, PV) ·
Expected Risk / EMV (Nominal, PV) · analytical mean total reconciliation inputs ·
per-year analytical cash-flow components · the **Calculation Input Fingerprint**
and CURRENT/STALE detection · refusal behaviour on every invalid numerical input ·
a worksheet-independent numerical kernel the simulation phase reuses.

### Out of scope — explicitly

MRG32k3a implementation · seed derivation · RNG stream identity · random variate
generation · Bernoulli occurrence simulation · Triangular / Beta-PERT / Uniform
sampling · simulation iterations · percentile ladder · P10/P50/P90 · Selected Px ·
contingency · histogram · CDF · sensitivity / Spearman · `_SimData` iteration
storage · Dashboard finalisation · Results finalisation · Model Check UI and
warning aggregation · annual percentiles · selected-Px annual profiles ·
simulation reconciliation · **any user-facing Calculate button** (§17).

---

## 4. Locked decisions D1–D6

These were open in Revision A. They are now **locked**.

### D1 — Uniform with a populated Most Likely

**Calculation succeeds. Most Likely is ignored numerically for Uniform.**

No Phase-5 warning or result-note mechanism is created for this. The locked
driver contract already states the position:

> `note: "Not used by the Uniform distribution. Greyed by conditional formatting
> when Distribution = Uniform; that is presentation only, not input enforcement."`

For calculation:

```
Uniform central = (Min + Max) / 2
Uniform mean    = (Min + Max) / 2
Uniform ML      = EXCLUDED from the calculation fingerprint, because it is unused
```

A future Model Check may raise a WARNING that a Uniform driver carries a populated
ML. That aggregation is **out of scope for Phase 5**.

### D2 — Inflation rate lower bound

**`rate > −1`**, equivalently `1 + rate > 0`. No additional business floor.
`rate = −1` or lower is **refused**.

### D3 — Discount rate sign

**`r > −1`.** Negative discount rates are allowed. No business minimum or maximum
is invented. `r = −1` or lower is **refused**.

### D4 — Blank profiling cells

**Any blank project-year profiling cell on an identified driver makes that profile
numerically incomplete, and calculation is refused** — even when the numeric cells
present happen to sum to 100%.

```
numeric 0  =  an explicit allocation of zero   → valid
blank      =  nothing has been stated          → refusal
```

Blank is not zero. Profiling year cells are **born as `0`**
(`initial_value: 0` in the structure contract), so a blank is a deliberate
clearing, never an unfilled default.

### D5 — Discounting terminology

Arithmetic locked:

```
discount factor(t) = 1 / (1 + r)^(t − 1)      Project Year 1 = period 0
```

Methodology wording locked to: **"Discounted from the start of the project;
Project Year 1 = period 0."** The phrase *"end-of-year discounting"* is not used
anywhere in the model, the code or the documentation.

This matches the locked input contract, quoted for the record:

> `project_start_year` — `note: "Whole calendar year. Discounting will treat this
> as period 0."`

### D6 — Precision and Phase-5 display

**All calculation and all stored numerical values use full VBA `Double`
precision. No calculation-rounding occurs anywhere.**

| Phase-5 `_Calc` audit value | Number format |
|---|---|
| SAR amounts | `#,##0.00` |
| factors (`Knom`, `Kpv`, inflation, discount, weights) | `0.000000` — at least six decimals, contract-defined |
| rates | `0.00%` |
| calendar year, project index | `0` |

Rounding is a **display** property of `_Calc` audit cells only. The final
Dashboard / Results presentation format is **not** locked here; it belongs to the
later presentation phase.

---

## 5. Mathematical definitions

### 5.1 Deterministic central value — risks excluded

| Distribution | Deterministic central | Central Basis label (§10) |
|---|---|---|
| Triangular | `ML` | `ML` |
| Beta-PERT | `ML` | `ML` |
| Uniform | `(Min + Max) / 2` | `Midpoint` |

Then `central × Quantity × FX × profiled inflation`.

Headline label: **Escalated Deterministic Base Estimate (Nominal SAR)**, with a PV
counterpart. **Never called "mean".**

### 5.2 Distribution expected value — the mean basis

| Distribution | Expected value |
|---|---|
| Triangular | `(Min + ML + Max) / 3` |
| Beta-PERT, λ = 4 | `(Min + 4·ML + Max) / 6` |
| Uniform | `(Min + Max) / 2` |

Applied to Quantity, FX, profile, inflation and discounting → **Mean-Basis Base
Cost**, Nominal and PV. The two bases coincide for Uniform and differ for
Triangular and Beta-PERT.

### 5.3 Expected Risk / EMV

Per risk: `Probability × Expected Severity`, expected severity by the same three
formulas on `impact_min` / `impact_most_likely` / `impact_max`, then
`× FX × risk profile × inflation`, and discounted for PV.

**All entered risks are included analytically.** No selection, no filtering.
`Probability` is stored as a fraction in `[0, 1]` and used directly.

---

## 6. Inflation convention

```
inflation_factor(Y) = 1                                    if Y = BaseYear
                    = Π (1 + rate_k),  k = BaseYear+1 … Y   otherwise
```

Base 2026, spend 2029 → `(1+r₂₀₂₇)(1+r₂₀₂₈)(1+r₂₀₂₉)`. If `BaseYear = StartYear`
the first project-year factor is **1** (empty product).

Consistent with the Phase-4 structural span already implemented:

> `nmInflFirstYear = BaseYear_Applied + 1` — "Escalation applies from the year
> after the applied base year." `nmInflLastYear = nmLastYear_Applied`.

When `BaseYear < StartYear` the span deliberately includes calendar years before
the project starts; those rates are required and are generated by Apply.

Profiles are **calendar-year anchored**. A start-year shift uses the rates of the
new calendar years; values never move positionally — proven on target.

**A missing required rate never becomes zero.** Inflation year cells are seeded
`null` precisely so an unmade assumption cannot be fabricated as 0%. Phase 5
refuses.

---

## 7. Discounting

Per D5. `inpDiscountRate` (Setup C20, `required: true`, `0.00%`), `r > −1`.

---

## 8. FX resolution — referenced currencies only

Locked convention: **1 source-currency unit = X SAR**, constant, no uncertainty.

| Rule | Scope |
|---|---|
| `SAR` resolves to exactly `1` | **global invariant**, always enforced (build-time locked seed row `["SAR", 1]`) |
| exactly one valid positive rate per currency | **referenced currencies only** |
| missing / blank / duplicate / non-numeric / `≤ 0` rate | refusal — **only if referenced** |

A currency is **referenced** when it appears in the `currency` column of an
identified Cost Line or Risk.

**An invalid, duplicate or incomplete FX row for an entirely unreferenced currency
must not block an otherwise valid analytical calculation.** This aligns with the
locked stale-input principle that unreferenced configuration changes do not make
results stale — and it falls out naturally, because the fingerprint (§11) records
only *resolved* FX per driver.

The same rule applies to inflation:

- only inflation profiles **referenced** by identified drivers must resolve;
- for each referenced profile, **every** required calendar year
  (`BaseYear+1 … LastProjectYear`) must carry a valid rate;
- an unused Config profile with incomplete assumptions does **not** block.

Both rules live in the resolution layer (`modCalcResolve`), which builds its
reference set from the driver registers before touching `tblFXRates` or
`tblInflation`.

---

## 9. Profiling semantics

Weights are applied **by permanent ID**, never by row position — proven on target
by Gate-B `B2` and `K2`.

For every identified driver, `Σ (weights over applied project years) = 100%`
within tolerance (§15), **and** no cell in that row may be blank (D4).

The engine uses the project-year columns of the **APPLIED** timeline only.
Entered-but-not-applied values never drive calculation. If `nmStructuralState`
reads `STRUCTURE CHANGE PENDING` or `No timeline applied`, calculation is refused.
Both labels come from the manifest (`state_labels`), never restated.

---

## 10. Precomputed factors and the per-driver audit record

```
Knom_i = FX_i × Σ_y ( w_{i,y} × infl_y )
Kpv_i  = FX_i × Σ_y ( w_{i,y} × infl_y × disc_y )
```

```
Cost line, nominal = unit_cost × Quantity × Knom     Risk, nominal = severity × Knom
Cost line, PV      = unit_cost × Quantity × Kpv      Risk, PV      = severity × Kpv
```

Probability is handled **separately** — multiplied in for analytical EMV, replaced
by a Bernoulli draw in Monte Carlo. It is deliberately **not** folded into `Kpv`.

### Central Basis is an explicit audit field

The per-driver audit record carries a **`Central Basis`** column with the literal
value `ML` or `Midpoint` (§5.1). An auditor must not have to infer the basis from
the Distribution column.

---

## 11. Calculation Input Fingerprint

Revision A's applied-triple fingerprint was **insufficient**: a calculation goes
stale with no timeline change at all — a changed Quantity, Probability,
Distribution, Min/ML/Max, Currency, FX, Inflation Profile, inflation rate,
profiling percentage or Discount Rate all invalidate stored results while the
applied triple is untouched.

Phase 5 therefore introduces the **analytical subset of the locked computational
fingerprint design**, as one canonical mechanism that Phase 6+ extends rather than
replaces.

### 11.1 Covered inputs

Sections are emitted in the fixed order `HEADER`, `COST`, `RISK` (§11.4), and the
field order within a record is exactly as listed here — the order is part of the
locked format, not an implementation choice.

**Header scalars** — one record, four fields

```
Applied Base Year · Applied Start Year · Applied Duration · Discount Rate
```

**Per identified Cost Line**

```
Permanent ID · Distribution · Quantity · Min · Max
ML  — ONLY when the distribution uses ML (Triangular, Beta-PERT). Excluded for Uniform (D1)
resolved FX · resolved inflation-factor vector · profiling-weight vector
```

**Per identified Risk**

```
Permanent ID · Distribution · Probability · Min · Max
ML  — ONLY when the distribution uses ML
resolved FX · resolved inflation-factor vector · profiling-weight vector
```

### 11.2 Exclusions — inherited from the locked stale-results design

Driver **row order** · per-driver digests **sorted by Permanent ID** · descriptive
fields · Category · Description · Risk Owner · UOM · Selected Confidence Level ·
Iterations · Random Seed · **unreferenced Config data** (a consequence of
recording *resolved* FX and inflation vectors rather than raw tables).

### 11.3 Canonical serialisation — LOCKED

Revision B joined fields with `U+001F` "which cannot appear in a workbook string".
**That assumption is withdrawn.** A cell can hold any character, including control
characters, and a delimiter-only design is ambiguous the moment one appears.

The encoding is **length-prefixed and self-delimiting**. No character is reserved,
and no value is escaped.

#### Field encoding

```
<TAG><LEN>:<VALUE>
```

| Part | Definition |
|---|---|
| `TAG` | one ASCII character: `S` text · `N` number · `I` stream integer |
| `LEN` | the count of **UTF-16 code units** in `VALUE`, in ASCII decimal, no sign, no padding |
| `:` | one ASCII colon, a fixed position marker — **not** a delimiter, because `LEN` already fixes where `VALUE` ends |
| `VALUE` | the canonical value (below) |

Because `LEN` states the exact extent of `VALUE`, a colon, `U+001F`, a line break,
a `NUL` or any other character **inside** `VALUE` cannot be mistaken for structure.
Two different field sequences can never produce the same stream.

#### Record and section encoding

```
section  ::=  F_S(name)  F_I(record_count)  record*
record   ::=  F_I(field_count)  field*
stream   ::=  F_S("PCCM-FP")  F_I(version)  section*
```

The per-record `field_count` is what makes a variable-length record unambiguous —
a 2-year weight vector cannot be confused with a 1-year vector plus one extra
field, and an omitted ML (D1) changes the count rather than shifting a position.

#### Text encoding — LOCKED

Hashed content is the sequence of **UTF-16 code units in logical VBA string
order**. Not "Unicode", not code points, not a byte encoding.

- VBA obtains each unit with `AscW`. **`AscW` returns a signed `Integer`, so any
  negative result must be normalised back to unsigned by adding 65536** before it
  enters the recurrence. A value above `0x7FFF` arrives negative; failing to
  normalise it changes the digest.
- A non-BMP character is stored by VBA as a surrogate **pair**, and therefore
  enters as **two** code units. Python must iterate the same way: the reference
  implementation encodes to `utf-16-le` and reads 16-bit units, so a non-BMP
  character contributes its two surrogates, exactly as VBA does. Iterating Python
  `str` by character would contribute **one** unit for such a character and
  silently diverge.
- `LEN` counts UTF-16 code units on both sides, by the same rule.

#### Numeric encoding — LOCKED and locale-invariant

**`Format$` is not assumed to be locale-invariant.** On a machine whose decimal
separator is `,` it emits `1,0000000000000000E+00`, which would produce a
different digest on the same inputs. The canonical form is therefore defined
independently of locale, and the implementation normalises before hashing:

| Rule | Value |
|---|---|
| significant digits | **17** — one before the point, sixteen after |
| decimal separator | **`.` always**, whatever the locale |
| exponent marker | **`E`**, uppercase |
| exponent sign | **always present**, `+` or `-` |
| exponent digits | **minimum 2**, zero-padded; more when needed (`E+308`, `E-324`) |
| thousands separator | **never** |
| negative zero | **normalised to positive zero** — `+0` and `−0` both encode as `0.0000000000000000E+00` |
| non-usable values | **refused before fingerprinting** (§19); NaN and ±∞ never reach the encoder |

**Implementation rule — and a purity constraint.** VBA's `Format$` emits the
*system* decimal separator, so normalisation to `.` before hashing is required.
But Revision C proposed reading
`Application.International(xlDecimalSeparator)` **inside** the canonical encoder,
which would put an Excel Application dependency into the numerical fingerprint
module and break the very boundary §17 exists to enforce.

**LOCKED:** `modCalcFingerprint` receives everything it needs to canonicalise a
number **as arguments**, and calls no Excel object.

> The resolution layer (`modCalcResolve`, or the orchestration) obtains the decimal
> separator **once** and passes it into the canonical numeric encoder.

The encoder is therefore a pure function of its arguments: the same value and the
same separator always produce the same string, on any machine, testable with no
Excel present. A Gate-A test injects `,` as the separator and asserts the output
is byte-identical to the `.` case.

*(An acceptable alternative is for the encoder to identify the mantissa's decimal
marker from its own `Format$` output and normalise that one character, still
without touching Excel state. The passed-in separator is preferred because it is
explicit and trivially testable.)*

#### Locked numeric test vectors

Computed from the specification, not from the implementation:

| Input | Canonical encoding |
|---|---|
| `0` | `0.0000000000000000E+00` |
| `-0` | `0.0000000000000000E+00` |
| `1` | `1.0000000000000000E+00` |
| `-1` | `-1.0000000000000000E+00` |
| `0.1` | `1.0000000000000001E-01` |
| `1e-20` | `9.9999999999999995E-21` |
| `1e+20` | `1.0000000000000000E+20` |
| `0.1 + 0.2` (all 17 digits significant) | `3.0000000000000004E-01` |
| `1.7976931348623157e308` (max double) | `1.7976931348623157E+308` |
| `5e-324` (min subnormal) | `4.9406564584124654E-324` |

`1e-20 → 9.9999999999999995E-21` is deliberately included: it is the true
17-digit form of the nearest double, and an implementation that "tidies" it to
`1.0000000000000000E-20` has broken round-tripping and will diverge.

Every one of these round-trips back to the identical double.

### 11.4 Hash algorithm — LOCKED

A change detector, not a security primitive. Nothing is left for an implementation
to choose.

| Constant | Value |
|---|---|
| `FP_BASE` | `131` |
| `FP_MOD_1` | `2147483647` (2³¹ − 1, prime) |
| `FP_MOD_2` | `2147483629` (prime) |
| `FP_INIT_1` | `1` |
| `FP_INIT_2` | `1` |
| `FP_VERSION` | `1` |

**Recurrence — mathematically unchanged.** For each UTF-16 code unit `u` of the
stream, in order, with `u` normalised to `0 … 65535`:

```
h1 = (h1 * 131 + u) mod 2147483647
h2 = (h2 * 131 + u) mod 2147483629
```

Both accumulators start at `1`, so a stream beginning with `NUL` is not absorbed.

**Hashed stream.** Tags, lengths, the colon and values are **all** hashed — the
entire canonical stream, **UTF-16 code-unit for UTF-16 code-unit**, nothing
excluded. The hash is defined over code units, never over a byte encoding.

**Section ordering — fixed, not sorted:** `HEADER`, `COST`, `RISK`. Phase 6
appends its sections after these; the analytical sections keep their positions so
the analytical subset stays comparable across phases.

**Driver-record ordering:** ascending by **Permanent ID**, ordinal comparison on
UTF-16 code units (`Option Compare Binary` semantics,
`StrComp(..., vbBinaryCompare)`). Never by row, never by digest.

**Final representation:** `HEX8(h1) & HEX8(h2)` — 16 characters, **uppercase**,
zero-padded to 8 each.

### 11.5 VBA implementation of the reduction — LOCKED

**The mathematics above is exact in a `Double`. The obvious VBA expression is
not.**

Revision D observed correctly that the intermediate is bounded by

```
(2147483647 − 1) × 131 + 65535 = 281,320,423,161  <  2⁵³
```

and concluded the recurrence "is therefore exact in VBA arithmetic". That is true
of the *arithmetic* and false of the *operator*. The VBA language specification
defines `Mod` on floating-point operands using an **effective integral type of
`Long`**, and the intermediate is **approximately 131 times the signed-`Long`
maximum**:

```
131 × 2,147,483,647  =  281,320,357,757
281,320,423,161      =  131 × Long.MaxValue + 65,404
```

*(Erratum E2: `281,320,423,161` is **not** equal to `131 × Long.MaxValue`; the two
exact statements above replace that claim. The reduction vectors below and their
expected remainders are unchanged.)*

Either way the intermediate is far outside signed 32-bit range. Writing

```vb
h = (h * 131 + u) Mod m          ' WRONG — overflows Long
```

would fail at runtime, or worse, silently mis-reduce. The same objection applies
to VBA integer division `\`, which is also `Long`-typed.

**LOCKED reduction — `Double` arithmetic only, no `Mod`, no `\`:**

```
x = h * FP_BASE + u
q = Fix(x / modulus)
r = x - q * modulus
If r >= modulus Then r = r - modulus
If r < 0        Then r = r + modulus
h = r
```

**Preconditions that make this exact for this hash:**

| Precondition | Why it holds |
|---|---|
| `0 ≤ h < modulus` | invariant of the loop; `h` starts at `1` |
| `0 ≤ u ≤ 65535` | UTF-16 code unit after `AscW` sign normalisation |
| `x` is an integer-valued `Double`, `x < 2⁵³` | max `281,320,423,161` |
| `q` is a small non-negative integer-valued `Double` | max **`131`** |
| `q × modulus < 2⁵³` | max `131 × 2147483647 = 281,320,357,757` |
| `r` is integer-valued, `0 ≤ r < modulus` | guaranteed by the two corrections |

The two corrections are not cosmetic. `x / modulus ≤ 131` carries a relative error
of at most `2⁻⁵³`, so the absolute error is under `1.5 × 10⁻¹⁴` and `Fix` can be
off by **at most one** in either direction; the `>=` and `< 0` adjustments absorb
exactly that.

**No pre-reduction intermediate — `x`, `q × modulus`, or any partial — may be
converted to `Long`.** Only *after* reduction, where `0 ≤ h < 2³¹`, may the
accumulator be converted to `Long` if the final hex rendering requires it.

**Verified.** The `Double`-only reducer was checked against exact integer
arithmetic over both moduli: all boundary combinations of
`h ∈ {0, 1, 2, m−2, m−1, m/2, …}` with `u ∈ {0, 1, 32768, 65535, …}`, plus
**300,000** random `(h, u, modulus)` triples. **Zero mismatches.** Running the
full reference stream through the reducer reproduces the locked digest exactly.

#### Locked reduction vectors

At the maximum recurrence intermediate — the case the `Long`-typed `Mod` would
fail on:

| Modulus | `h` | `u` | `x = h·131 + u` | `q = Fix(x/m)` | `r = x − q·m` |
|---|---|---|---|---|---|
| `FP_MOD_1` `2147483647` | `2147483646` | `65535` | `281320423161` | `131` | **`65404`** |
| `FP_MOD_2` `2147483629` | `2147483628` | `65535` | `281320420803` | `131` | **`65404`** |

And a non-degenerate mid-range case, so the test set is not only extremes:

| Modulus | `h` | `u` | `x` | `q` | `r` |
|---|---|---|---|---|---|
| `FP_MOD_1` | `1234567890` | `41` | `161728393631` | `75` | **`667120106`** |
| `FP_MOD_2` | `1234567890` | `41` | `161728393631` | `75` | **`667121456`** |

Every `x` above exceeds `Long.MaxValue`, which is the point: these vectors fail
against a `Mod`-based implementation and pass against the locked one.

**No hash constant changes, and the end-to-end digest is unchanged:
`50B6EB0E26857EA7`.**

### 11.6 The locked reference vector

Golden case 1 (§23): Base 2026, Start 2026, Duration 1, `r = 0.10`; one Cost Line
`CL-001`, Triangular, Qty 10, Min 80, Max 150, ML 100, FX 1, inflation vector
`[1]`, weight vector `[1]`; no risks.

Field order within a cost record:
`ID · Distribution · Quantity · Min · Max · [ML if used] · FX · infl_y… · w_y…`
Within a risk record, `Probability` replaces `Quantity`.

Canonical stream — **366 UTF-16 code units**:

```
S7:PCCM-FPI1:1S6:HEADERI1:1I1:4N22:2.0260000000000000E+03N22:2.0260000000000000E+03
N22:1.0000000000000000E+00N22:1.0000000000000001E-01S4:COSTI1:1I1:9S6:CL-001
S10:TriangularN22:1.0000000000000000E+01N22:8.0000000000000000E+01
N22:1.5000000000000000E+02N22:1.0000000000000000E+02N22:1.0000000000000000E+00
N22:1.0000000000000000E+00N22:1.0000000000000000E+00S4:RISKI1:0
```

*(shown wrapped for readability; the stream contains no line breaks)*

```
EXPECTED FINGERPRINT = 50B6EB0E26857EA7
```

**Python and VBA must both produce this exact literal.** The two proofs live in
different gates and must not be conflated (erratum E3, §21.0):

- **Gate A / Linux** — the **Python** reference implementation produces
  `50B6EB0E26857EA7` and is asserted against this literal. The VBA is checked
  **statically only**: no VBA is executed on Linux, so no Gate-A test may claim
  that VBA produced this digest.
- **Gate B / Windows Excel** — real VBA execution produces the digest through
  `PCCM_CurrentInputFingerprint()`, and parity with the Python literal is asserted
  on target (§24, §25.6).

### 11.7 Collision probes — locked test set

Field sequences whose *content* contains the characters a delimiter design would
have collided on. All must produce distinct digests:

| Fields | Digest |
|---|---|
| `["A:B", "C"]` | `041ACBD05C7BF72C` |
| `["A", "B:C"]` | `52704E9A542869CA` |
| `["AB", "C"]` | `0C8A057A0BE7EB51` |
| `["A", "B", "C"]` | `7674F1C35E639F98` |
| `["A" U+001F "B", "C"]` | `7D26D4C95587DE0C` |
| `["A" U+0000 "B", "C"]` | `0821AFB0608291C8` |
| `["A" U+000A "B", "C"]` | `5B4CA2E133AD91A2` |
| `["A", U+001F, "C"]` | `101504AC7803B226` |

All eight are distinct under the length-prefixed encoding.

**Revision C's explanation of this table was wrong and is corrected here.** It
claimed rows 1–2, 3–5 and 4–8 would collide under a `U+001F` join. A direct
analysis of the flattened streams shows otherwise:

| Encoding | Colliding rows |
|---|---|
| `U+001F` join | **4 ↔ 5** only — `["A","B","C"]` and `["A" U+001F "B","C"]` both flatten to `A␟B␟C` |
| colon join | **1 ↔ 2** and **1 ↔ 4** |
| length-prefixed | **none** |

Rows 1–2 do *not* collide under `U+001F` (`A:B␟C` versus `A␟B:C`), and neither do
3–5 or 4–8. The corrected claim is the more useful one anyway:

> The probe set demonstrates that **arbitrary delimiter-based encodings are
> unsafe**. The `U+001F`-specific collision is included explicitly (rows 4–5), and
> the colon and control-character content proves that **changing the delimiter
> does not solve the general ambiguity** — it only moves which content breaks it.

The eight vectors and their expected digests are unchanged; only the prose was
wrong.

### 11.8 `calc_state` — two orthogonal axes

Revision B conflated four things. Revision C separated the snapshot from the
attempt but still let `REFUSED` appear on the status axis. **They are now fully
orthogonal, and each axis answers exactly one question:**

```
DERIVED STATUS   "What do the CURRENT inputs say about the stored successful snapshot?"
ATTEMPT RESULT   "What happened the last time Calculate was explicitly attempted?"
```

| Field | Axis | Meaning |
|---|---|---|
| `Last Successful Stamp` | snapshot | timestamp of the last calculation that **completed and committed** |
| `Last Successful Fingerprint` | snapshot | fingerprint of the inputs of that calculation |
| `Fingerprint Version` | snapshot | `FP_VERSION` at the time of that success |
| `Last Successful Applied Timeline` | snapshot | the applied triple it used |
| `Last Attempt Result` | attempt | `NONE` / `SUCCESS` / `REFUSED` / `FAILED` |
| `Last Attempt Detail` | attempt | blank after `NONE`/`SUCCESS`; the refusal reason after `REFUSED`; the controlled internal/write failure detail after `FAILED` |
| `Calculation Status (last evaluated)` | derived | `NOT CALCULATED` / `CURRENT` / `STALE` / `INVALID` |
| `Status Evaluated At` | derived | when that status was last computed |

`Last Refusal Reason` is **renamed** to `Last Attempt Detail`. A refusal-only
field cannot report a write failure, and there is no compatibility requirement to
preserve — no implementation exists yet — so the generic field is simply the
correct one.

### 11.9 Derived status — LOCKED, four values only

`PCCM_CalculationStatus()` returns **only**:

```
NOT CALCULATED   no last-successful fingerprint, and the current inputs resolve
INVALID          the current inputs cannot resolve
CURRENT          inputs resolve and current fingerprint == last successful fingerprint
STALE            inputs resolve and current fingerprint != last successful fingerprint
```

**`REFUSED` is not a derived status.** Revision C's first branch —
*"no last-successful fingerprint and the most recent attempt refused → REFUSED"* —
mixed the two axes and contradicted the rest of §11.9. It is removed.

If no successful calculation has **ever** occurred and the first attempt refuses:

```
Derived Status       INVALID          (the inputs genuinely do not resolve)
Last Attempt Result  REFUSED
Last Attempt Detail  the specific refusal message
```

The status is `INVALID` because that is what the *inputs* are; the fact that
someone pressed Calculate and was turned away is *attempt history*, and lives on
the other axis. No special case is made.

### 11.10 Attempt result — LOCKED, four values

| Value | Meaning |
|---|---|
| `NONE` | no `PCCM_Calculate` attempt has occurred |
| `SUCCESS` | calculation completed **and committed** |
| `REFUSED` | calculation did not begin write-back, because structural or numerical prerequisites, or the controlled numerical-range rules, rejected the current inputs |
| `FAILED` | an internal, runtime, write or verification failure occurred **after** calculation orchestration began |

`REFUSED` and `FAILED` are genuinely different events: one is the model correctly
declining invalid inputs, the other is the machinery failing on inputs it had
already accepted. Revision C offered only `SUCCESS` / `REFUSED` while also
requiring an injected mid-write failure scenario — there was no value to report it
with.

**The derived status after a `FAILED` attempt is computed independently**, from
the *restored* successful snapshot and the current inputs:

| Current inputs | Derived status after rollback |
|---|---|
| fingerprint == restored successful fingerprint | `CURRENT` |
| a different valid fingerprint | `STALE` |
| do not resolve | `INVALID` |

**Status is never forced to `FAILED`.** In particular, a mid-write failure while
the inputs describe a *new* calculation leaves `STALE` — a successful rollback
must not make the new inputs look `CURRENT` merely because the rollback worked.

### 11.11 The locked revert-to-CURRENT answer

Successful calculation → user makes an input invalid → `PCCM_Calculate` refuses →
user restores the input *exactly* to the previously successful state → status is
queried, with no recalculation:

```
Derived Status       CURRENT
Last Attempt Result  REFUSED     (unchanged — it is history)
Last Attempt Detail  the refusal message, still readable
```

The inputs resolve and their fingerprint equals the last successful fingerprint,
so the stored snapshot genuinely describes them. **A historical failed attempt
never overrides a currently matching successful snapshot**, and equally, a
successful snapshot never erases the record that an attempt was refused.

`INVALID` is distinct from `STALE`: an unresolvable current state cannot produce a
fingerprint at all, so claiming "stale" would assert a comparison that was never
made. `CURRENT` is never returned in that case.

### 11.12 Status is last-evaluated, not live

**There are no change events** — no `Worksheet_Change`, no
`Workbook_SheetChange`, consistent with the Phase-4 rule that structural state is
never maintained by hidden automation.

The `_Calc` status cell therefore holds a **last-evaluated** status, refreshed by
`PCCM_Calculate`, by `PCCM_CalculationStatus`, and by later point-of-consumption
guards (Run Check, Run Simulation, output refresh). **It does not update
spontaneously**, and the block is labelled `Calculation Status (last evaluated)`
with a `Status Evaluated At` timestamp beside it, so an auditor reading the sheet
can see how old the reading is.

### 11.13 Callable surface for Gate B


Public, invoked by `Application.Run` — **no button** (§17):

```
PCCM_Calculate                     orchestration; transactional (§12)
PCCM_CalculationStatus()           re-evaluates and returns the DERIVED status
                                   NOT CALCULATED | CURRENT | STALE | INVALID
PCCM_CalculationAttemptResult()    NONE | SUCCESS | REFUSED | FAILED
PCCM_CalculationAttemptDetail()    the refusal reason, the failure detail, or empty
PCCM_CalculationFingerprint()      the LAST SUCCESSFUL fingerprint
PCCM_CurrentInputFingerprint()     the fingerprint of the inputs as they are NOW
```

**Six public `PCCM_` automation/API entry points in total: `PCCM_Calculate` plus
five accessors.**

*(Erratum E4.)* This is a bound on the **`PCCM_` endpoint surface**, not on the
`Public` keyword. Numerical helper procedures inside `modCalcFactors`,
`modCalcAnalytical`, `modCalcFingerprint`, `modCalcResolve`, `modCalcCheck` and
`modCalcReport` may be declared `Public` wherever a cross-module VBA call requires
it — VBA has no narrower visibility between standard modules. What is prohibited is
**any additional `PCCM_`-prefixed entry point in Phase 5**, which may not be
introduced without review.

`PCCM_CalculationRefusal()` from Revision C is **replaced**, not kept. A
refusal-only accessor cannot report a `FAILED` write, so it could not express the
mid-write scenario the plan already requires; and with no implementation in
existence there is no compatibility argument for carrying it. The generic attempt
pair is the whole surface.

The two axes are separately readable, and the two fingerprint accessors stay
separate, so Gate B can assert all four values independently — the stored snapshot
**unchanged** while the current fingerprint has moved (§25).

#### Empty and invalid cases — LOCKED

| Accessor | Boundary behaviour |
|---|---|
| `PCCM_CalculationFingerprint()` | `""` when no successful calculation exists; otherwise the 16-character last-successful fingerprint |
| `PCCM_CurrentInputFingerprint()` | the current 16-character fingerprint when the inputs resolve; **`""` when they cannot** |
| `PCCM_CalculationAttemptResult()` | `NONE` before the first explicit `PCCM_Calculate` attempt |
| `PCCM_CalculationAttemptDetail()` | blank for `NONE` and for `SUCCESS` |
| `PCCM_CalculationStatus()` | never `REFUSED`; see §11.9 |

**No sentinel hash strings are invented.** An empty string means "there is no
digest", and must never be read as a valid one — in particular, two unresolvable
states both returning `""` must not be treated as "matching fingerprints", which
is why §11.9 derives `INVALID` from the *resolution failure*, not from a
fingerprint comparison.

When the current inputs cannot resolve, the resolution detail remains available
through the calculation and status failure surface (`modAppState`), so an empty
fingerprint is never the only thing the user is told.

---

## 12. Transactional write-back

Revision B required that a refused calculation not overwrite the previous
snapshot. **That is extended to write failures.** The Phase-5 audit blocks are one
snapshot, and a half-old / half-new `_Calc` is not an acceptable outcome.

### 12.1 Locked orchestration

```
1. resolve everything into memory
2. validate everything
3. calculate everything in memory
4. reconcile everything in memory        (identities I1–I5)
5. build the complete fingerprint in memory
6. ONLY THEN begin workbook write-back
```

**No `_Calc` analytical result is written during steps 1–5.**

### 12.2 What a pre-write refusal leaves behind — corrected

Revision C said a refusal *"leaves the workbook byte-identical"* while also
requiring the attempt metadata to be written on refusal. **Both cannot be true**,
and the correct statement is narrower:

> A pre-write refusal leaves every **last-successful analytical snapshot block**
> logically unchanged. Attempt and status metadata **do** change — that is the
> point of recording an attempt.

| Must be UNCHANGED on refusal | May / must CHANGE on refusal |
|---|---|
| `calc_totals` | `Last Attempt Result` → `REFUSED` |
| `tblCalcYears` | `Last Attempt Detail` → the specific reason |
| `tblCalcInflationFactors` | `Calculation Status (last evaluated)` → `INVALID` |
| `tblCalcFX` | `Status Evaluated At` → the current timestamp |
| `tblCalcDrivers` | |
| `tblCalcAnnual` | |
| `Last Successful Fingerprint` | |
| `Last Successful Stamp` | |
| `Fingerprint Version` (of that success) | |
| `Last Successful Applied Timeline` | |

**The phrase "workbook byte-identical" is not used of a refused attempt anywhere
in this plan.** Acceptance compares the two groups separately (§28).

The physical layout of §16.4 makes this trivially checkable: the four
last-successful fields are one contiguous block, and the four attempt/status
fields are another, so "unchanged" and "changed" are each a single range
comparison.

### 12.3 The success commit is ONE range assignment — and it is itself fallible

Revision C wrote three fields "last" and called that *"a single scalar write"*.
Three writes are not one write, and a failure between them is exactly the mixed
state the design forbids.

**LOCKED: the success commit is one contiguous `Range.Value2` array assignment.**

`calc_state` is ordered (§16.4) so that this is possible:

```
C13:C16   Last Successful Stamp / Fingerprint / Fingerprint Version / Applied Timeline
C17:C18   Last Attempt Result / Last Attempt Detail
C19:C20   Calculation Status (last evaluated) / Status Evaluated At
```

| Operation | Range written | Assignments |
|---|---|---|
| **success commit** | `C13:C20` | **one** 8×1 `Value2` array |
| refusal / failure record | `C17:C20` | one 4×1 array — `C13:C16` provably untouched |
| status-only refresh | `C19:C20` | one 2×1 array |

Option A of the review is taken: **status and `Status Evaluated At` are inside the
commit block**, so a successful calculation publishes its analytical snapshot, its
attempt result and its derived status in a single assignment. There is no window
in which the snapshot is committed but the status still says `STALE`.

#### The assignment is a COMMIT ATTEMPT, not an infallible act

**Revision D over-claimed.** A single `Range.Value2` assignment is indivisible
from the *model's* point of view — no half-written state is observable — but the
assignment itself can raise, and its verification can fail. It is therefore
**inside** the transaction, not outside it:

> If the `C13:C20` assignment raises, or verification of that assignment fails,
> **the calculation has NOT committed**, and the same rollback path applies.

The commit range remains one assignment, and remains the last success write. The
correction is only that failure *of* that assignment is explicitly transactional
rather than sitting outside the rollback envelope.

### 12.4 Rollback — the locked sequence

Write-back covers `tblCalcFX`, `tblCalcYears`, `tblCalcInflationFactors`,
`tblCalcDrivers`, `tblCalcAnnual`, `calc_totals` and `calc_state`.

The table strategy reuses **the Phase-4 logical-rollback mechanism already proven
on target** — `modWorkbook.SnapshotTable` / `RestoreTable`, including its
collision-safe header restoration.

#### Scalar rollback is explicit — those helpers do not cover it

`SnapshotTable` / `RestoreTable` operate on ListObjects. The two scalar blocks
need their own, stated scope.

**Captured before write-back:** the **value cells only** —

```
calc_totals   C23:C32     ten SAR values
calc_state    C13:C20     eight values, including blanks and timestamps
```

**Not captured, and never rewritten during calculation:** labels (`B`), notes
(`E`), number formats, and every structural property. Those are build-owned; a
calculation that rewrote them would be repairing structure, which is Phase-4
territory.

**On rollback:** the captured value cells are restored **exactly**, including
blanks (restored as blank, never as `0` or `""`) and timestamps. Labels and
formats are untouched because they were never captured.

#### The sequence — LOCKED

```
1. snapshot the five ListObjects
2. snapshot calc_totals      C23:C32
3. snapshot calc_state       C13:C20
4. write the analytical tables
5. write calc_totals
6. verify the analytical tables and calc_totals against the in-memory values
7. attempt ONE C13:C20 Value2 success-commit assignment, then verify it
8. if step 7 succeeds:
       the calculation is COMMITTED
       no further workbook mutation may make the operation fail
9. if ANY of steps 4-7 fails:
       restore the five ListObjects
       restore C23:C32
       restore the prior C13:C20
       ONLY AFTER successful rollback:
           record the FAILED attempt and the re-derived status in C17:C20
```

Steps 4–7 are all inside the envelope: a mid-table failure and a failed commit
assignment take the same path. Step 8 is the only point after which the operation
cannot be turned into a failure.

**If recording the `FAILED` attempt metadata itself fails**, the Phase-4
`modAppState` failure surface still reports the failure to the user, and **the
previous successful snapshot remains authoritative** — it was restored in step 9
before any attempt metadata was touched. The ordering is deliberate: restoration
never depends on the success of the bookkeeping that follows it.

### 12.5 Post-failure state — two distinct moments

Revision D's sequence was right and its acceptance wording was not. After a
failure there are **two observable moments**, and only the first has `C13:C20`
wholly restored:

**Moment 1 — immediately after rollback, before failure metadata**

```
C13:C20   restored EXACTLY to the previous values, all eight cells
```

**Moment 2 — the final observable state, after the failed attempt is recorded**

| Cells | State |
|---|---|
| `C13:C16` | **exactly** the restored last-successful snapshot |
| `C17` | `FAILED` |
| `C18` | the failure detail |
| `C19` | the derived status, from current inputs vs the restored snapshot |
| `C20` | a new status-evaluation timestamp |
| `C23:C32` and the five ListObjects | **exactly** restored to the previous successful calculation |

**Acceptance therefore compares `C13:C16` plus the analytical blocks — never all
of `C13:C20` after failure metadata has been written.** Claiming the latter would
be asserting that the failure was not recorded.

### 12.6 Acceptance

A Windows injected-failure scenario (§25) proves the path at **both** boundaries —
a mid-table failure and a failure of the final commit assignment — and asserts:
previous totals restored · previous driver audit rows restored · previous annual
rows restored · `C13:C16` and `C23:C32` restored exactly, blanks as blanks · no
mixed snapshot survives · `Last Attempt Result = FAILED` · the derived status
computed independently (`STALE` when the current inputs describe the attempted new
calculation) · Excel application state restored.

This is a **real acceptance requirement**, not a diagnostic.

---

## 13. The three headline measures

Let `Q_i` = Quantity, `c_i` = deterministic central, `m_i` = distribution mean,
`p_j` = probability, `s_j` = expected severity.

```
A_nom = Σ_i c_i·Q_i·Knom_i     A_pv = Σ_i c_i·Q_i·Kpv_i     Escalated Deterministic Base
C_nom = Σ_i m_i·Q_i·Knom_i     C_pv = Σ_i m_i·Q_i·Kpv_i     Mean-Basis Base Cost
D_nom = Σ_j p_j·s_j·Knom_j     D_pv = Σ_j p_j·s_j·Kpv_j     Expected Risk / EMV
```

---

## 14. Annual analytical cash flow

Per applied project year `y`, six values:

```
Base Cost — Nominal        Σ_i m_i·Q_i·FX_i·w_{i,y}·infl_{i,y}
Expected Risk — Nominal    Σ_j p_j·s_j·FX_j·w_{j,y}·infl_{j,y}
Total — Nominal            the two above
Base Cost — PV             same × disc_y
Expected Risk — PV         same × disc_y
Total — PV                 the two above
```

**Annual Base Cost uses the distribution expected value `m_i`, not the
deterministic ML/Midpoint basis.** The locked Results requirement is that annual
cash flow is mean-only, so the annual series is *Mean-Basis Base Cost + Expected
Risk*. The deterministic basis has no annual series.

No annual percentiles. No selected-Px annual profile.

---

## 15. Reconciliation identities and tolerances

`B` and `E` are **independently accumulated**, not derived, so the identities are
real checks rather than tautologies:

```
B_nom = Σ_i (m_i − c_i)·Q_i·Knom_i      B_pv = Σ_i (m_i − c_i)·Q_i·Kpv_i
E_nom = Σ_i m_i·Q_i·Knom_i + Σ_j p_j·s_j·Knom_j     (accumulated in its own pass)
```

| # | Identity |
|---|---|
| I1 | `A + B = C` — Nominal and PV |
| I2 | `C + D = E` — Nominal and PV |
| **I3a** | `Σ_y annual Mean-Basis Base Nominal = C_nom` |
| **I3b** | `Σ_y annual Expected Risk Nominal = D_nom` |
| **I3c** | `Σ_y annual Total Nominal = E_nom` |
| **I4a** | `Σ_y annual Mean-Basis Base PV = C_pv` |
| **I4b** | `Σ_y annual Expected Risk PV = D_pv` |
| **I4c** | `Σ_y annual Total PV = E_pv` |
| I5 | `Σ_y w_{i,y} = 1` per driver — profiling validation |

Splitting the annual reconciliation into base / risk / total is strictly stronger
than checking the total alone: a base amount misclassified as risk would cancel in
a total-only check.

### `A_pv ≤ A_nom` is NOT an identity — removed from acceptance

Revision A listed it. It is **not** universally true: no locked contract imposes a
non-negative Unit Cost rule, so a negative deterministic contribution reverses the
inequality, and `r < 0` is explicitly allowed (D3). Inventing a non-negativity
rule to rescue it would be inventing a business rule.

It is retained **only as a conditional diagnostic**, reported and never gating:
when every deterministic nominal contribution is `≥ 0` **and** `r ≥ 0`, a
violation of `A_pv ≤ A_nom` indicates a discount factor applied with the wrong
sign or index. **It is not a Phase-5 validity gate.**

### Tolerances and identity conditioning

| Purpose | Tolerance |
|---|---|
| profiling sum = 100% | `\|Σw − 1\| ≤ 1e-9` absolute |
| identities I1–I4 | `\|Δ\| ≤ max(1e-6, 1e-12 × conditioning_scale)` SAR |
| FX positivity | `rate > 0`, no epsilon |
| `1 + rate > 0`, `1 + r > 0` | strict |

The profiling tolerance is unchanged: percentages are entered to 2 dp and stored
as binary doubles, so across the 200-column structural maximum the accumulated
representation error is bounded by roughly `200 × 2⁻⁵² ≈ 4.4e-14`. `1e-9` gives
~4 orders of headroom and is still 3 orders tighter than a 1-in-a-million entry
slip. Exact binary equality is never used.

The absolute floor `1e-6 SAR` and the relative coefficient `1e-12` are unchanged.

#### The conditioning scale must survive cancellation

Revision B used `scale = max(|A|, |C|, |E|)`. **That is wrong when contributions
cancel.** Negative unit costs and impacts are not prohibited by any locked
contract, so a model can hold billions of positive and billions of negative
contributions whose *net* is near zero. `max(|A|,|C|,|E|)` would then collapse to
almost nothing, the tolerance would fall back to the `1e-6` floor, and ordinary
floating-point accumulation error across hundreds of large terms would be reported
as a bookkeeping mismatch.

The scale must reflect the **magnitude of the arithmetic performed**, not the
magnitude of its net result.

#### ERRATUM C1 — conditioning is on CONTRIBUTIONS, not on totals or aggregates

*Raised by Gate-A Step 2 and applied here. This is a narrow numerical correction
to the conditioning OPERANDS only. No tolerance number changes, no business rule
changes, and no change to A / B / C / D / E themselves.*

Revision E named the scales as follows, and **both halves were proven
insufficient by implementation**:

| Identity | Superseded scale | Why it fails |
|---|---|---|
| I1 `A + B = C` | `max(1, \|A\| + \|B\| + \|C\|)` | A, B and C are already-cancelled **totals** |
| I2 `C + D = E` | `max(1, \|C\| + \|D\| + \|E\|)` | same |
| I3 / I4 | `max(1, Σ_y \|annual aggregate\| + \|headline\|)` | the annual **row aggregate** has already cancelled *within* the year |

Both were demonstrated on valid models — ordered three-point sets, positive
Quantity, weights summing to 1 — that a correct calculation then reported as
failing:

```
HEADLINE.  Three cost lines, two of them exact mirrors:
             CL-001  Min 0      ML 1e17   Max 4e17
             CL-002  Min 10     ML 30     Max 110
             CL-003  Min -4e17  ML -1e17  Max 0
           A = 32, B = 16, C = 64, so |A|+|B|+|C| = 112 and the scale
           collapses to the 1e-6 floor. But the accumulation ran through
           partial sums of 1e17, where ONE ULP IS ALREADY 16 SAR.
           I1 difference = -16   ->  reported as failing.

ANNUAL.    Duration 2, one currency, zero inflation, zero discount:
             CL-001  1e16   profile 100% / 0%
             CL-002  1      profile   0% / 100%
             CL-003  -1e16  profile 100% / 0%
           Year 1 aggregate = 0 (the two 1e16 contributions annihilate),
           year 2 aggregate = 1, so Σ_y |annual| = 1 — yet the annual
           arithmetic processed about 2e16.
           I3a difference = 1   ->  reported as failing.
```

**The corrected scales sum the UNDERLYING CONTRIBUTIONS**, per driver and per
driver per year, before any aggregation. For each cost driver `i` let `A_i`,
`B_i`, `C_i` be its deterministic, uncertainty-mean-shift and mean-basis
contributions; for each risk `j` let `D_j` be its expected-risk contribution; and
let `E_k` be the contribution actually accumulated into `E` (`C_i` for a cost
line, `D_j` for a risk):

| Identity | Conditioning scale |
|---|---|
| I1 `A + B = C` | `max(1, Σ_i \|A_i\| + Σ_i \|B_i\| + Σ_i \|C_i\|)` |
| I2 `C + D = E` | `max(1, Σ_i \|C_i\| + Σ_j \|D_j\| + Σ_k \|E_k\|)` |
| I3a / I4a | `max(1, Σ_y Σ_i \|base_{i,y}\| + Σ_i \|C_i\|)` |
| I3b / I4b | `max(1, Σ_y Σ_j \|risk_{j,y}\| + Σ_j \|D_j\|)` |
| I3c / I4c | `max(1, Σ_y (Σ_i \|base_{i,y}\| + Σ_j \|risk_{j,y}\|) + Σ_k \|E_k\|)` |

Nominal and PV identities each use their own basis's contribution magnitudes.
**`Σ_y |annual aggregate_y|` must not be used as a substitute**: cancellation may
already have happened inside that aggregate, which is exactly the annual failure
above.

`max(1, …)` keeps the scale from going below unity for a genuinely tiny model, so
the `1e-6` absolute floor remains the binding constraint there.

**Unchanged by this erratum:** `profiling_sum_absolute = 1e-9`,
`identity_absolute_floor = 1e-6`, `identity_relative_coefficient = 1e-12`,
`conditioning_scale_floor = 1`. `spec/calc_contract.yaml` carries the corrected
term NAMES; the arithmetic remains the plan's.

#### The allowance is a maximum, not a sum

For the avoidance of the implementation ambiguity Step 2 found:

```
allowance = max( absolute_floor, coefficient * max(scale_floor, conditioning_scale) )
```

The inner operation is a **maximum**. Adding `coefficient × scale_floor` to the
scaled scale would widen every allowance slightly, and a tolerance may not be
loosened by accident. The scaled form — distributing the coefficient across the
terms rather than multiplying their sum — remains required, because the raw
conditioning sum can exceed `Double` while the tolerance itself is perfectly
representable.

The objective is stated deliberately: **tolerate normal floating-point
accumulation error without hiding a real bookkeeping mismatch.** A conditioning
scale driven by the terms rather than the total is what achieves both.

No Monte Carlo statistical tolerance is invented here.

---

## 16. Physical `_Calc` layout — locked

`_Calc` is `hidden` (not veryHidden) so an auditor can inspect it. `_SimData`
remains `veryHidden` and **untouched and unused** by Phase 5. No user-facing input
sheet is written by Phase 5.

### 16.1 Phase-4 reservation

The frozen sheet currently occupies **rows 1–11** in columns B, C, E:

```
B2  title            B3  subtitle          B6  note
B8  "Permanent ID Counters"
B10 "Cost Line ID Counter"  C10 counter (nmCounterCostLine)  E10 note
B11 "Risk ID Counter"       C11 counter (nmCounterRisk)      E11 note
```

**Rows 1–11 are Phase-4 territory and are not touched.** `_Calc!C10:C11` is
declared reserved in the new contract so no Phase-5 block can be placed there, and
a build-time assertion fails if any Phase-5 block overlaps rows 1–11 or the
counter cells.

### 16.2 Growth strategy — column bands, not vertical stacking

Every Phase-5 table has a **fixed column schema** and an **unbounded row count**.
Stacking them vertically would make a growing table collide with the block below,
so each dynamic ListObject gets its **own column band**, all anchored at the same
header row, growing downward with nothing beneath it.

**No block is capped at 200 Cost Lines or 100 Risks.** Those are design targets,
not business maxima — consistent with the Phase-4 refusal to encode 25 years.

### 16.3 Exact anchors — LOCKED, no longer illustrative

All ListObjects anchor their header at **row 15**. Two-column gutters separate the
bands, so a widened schema is caught by the build-time overlap assertion rather
than by silently overwriting a neighbour.

| Block | Kind | Columns | Header row | Rows | Growth | Visible when unhidden | Owner |
|---|---|---|---|---|---|---|---|
| *(Phase-4 counters)* | scalars | `B`, `C`, `E` — rows 8–11 | — | fixed | none | yes | **Phase 4 — reserved** |
| `calc_state` | scalars | `B` label, `C` value, `E` note — rows 13–20 | — | 8, fixed | none | yes | Phase 5 |
| `calc_totals` | scalars | `B` label, `C` value, `E` note — rows 23–32 | — | 10, fixed | none | yes | Phase 5 |
| `tblCalcYears` | ListObject | **`H:J`** (3) | 15 | applied year count | vertical | yes | Phase 5 |
| `tblCalcInflationFactors` | ListObject | **`M:P`** (4) | 15 | referenced profiles × factor years | vertical | yes | Phase 5 |
| `tblCalcFX` | ListObject | **`S:U`** (3) | 15 | referenced currencies | vertical | yes | Phase 5 |
| `tblCalcDrivers` | ListObject | **`X:AR`** (21) | 15 | identified cost lines + risks | vertical | yes | Phase 5 |
| `tblCalcAnnual` | ListObject | **`AU:BB`** (8) | 15 | applied year count | vertical | yes | Phase 5 |

`calc_contract.yaml` **encodes these anchors; it does not choose them.** The
loader asserts the contract's anchors equal these values, that no two bands
overlap given their declared schemas, and that no Phase-5 block intersects rows
1–11 or the counter cells `_Calc!C10:C11`.

The B/C scalar column is shared by `calc_state` and `calc_totals` at disjoint row
ranges; both are fixed-height, so neither can grow into the other, and rows 21–22
and 33 onward are left free.

### 16.4 Schemas — LOCKED

#### `calc_state` — `B13:C20`

**Row order is load-bearing** — it is what makes the success commit one contiguous
assignment (§12.3). The three groups are kept adjacent and in this order.

| Row | Group | Label | Value | Format |
|---|---|---|---|---|
| 13 | **snapshot** | Last Successful Stamp | timestamp | `yyyy-mm-dd hh:mm:ss` |
| 14 | **snapshot** | Last Successful Fingerprint | 16 uppercase hex characters | `@` |
| 15 | **snapshot** | Fingerprint Version | integer | `0` |
| 16 | **snapshot** | Last Successful Applied Timeline | `base/start/duration` | `@` |
| 17 | **attempt** | Last Attempt Result | `NONE` / `SUCCESS` / `REFUSED` / `FAILED` | `@` |
| 18 | **attempt** | Last Attempt Detail | blank after `NONE`/`SUCCESS`; refusal reason after `REFUSED`; failure detail after `FAILED` | `@` |
| 19 | **derived** | Calculation Status (last evaluated) | `NOT CALCULATED` / `CURRENT` / `STALE` / `INVALID` | `@` |
| 20 | **derived** | Status Evaluated At | timestamp of the last status evaluation | `yyyy-mm-dd hh:mm:ss` |

- `C13:C16` is the last-successful snapshot — **provably untouched** by a refusal.
- `C17:C18` is attempt history.
- `C19:C20` is the derived reading.
- `C13:C20` is the **success commit**, written as one 8×1 array (§12.3).

`REFUSED` no longer appears among the status values (§11.9); it is an attempt
result. Row 19's label says **"(last evaluated)"** because it is not live
(§11.12), and row 20 makes that concrete: an auditor can see how old the reading
is.

##### Initial values, seeded by Stage A — LOCKED

| Cell | Field | Initial value |
|---|---|---|
| `C13` | Last Successful Stamp | **blank** |
| `C14` | Last Successful Fingerprint | **blank** |
| `C15` | Fingerprint Version | **blank** |
| `C16` | Last Successful Applied Timeline | **blank** |
| `C17` | Last Attempt Result | **`NONE`** |
| `C18` | Last Attempt Detail | **blank** |
| `C19` | Calculation Status (last evaluated) | **`NOT CALCULATED`** |
| `C20` | Status Evaluated At | **blank** |

**`Fingerprint Version` is blank before any successful snapshot exists.**
`FP_VERSION = 1` is the current algorithm version held in the contract and the
code; it is written into `C15` **only as part of a successful commit**. Seeding it
at build time would make a never-calculated workbook look as though it held a
partial successful snapshot — the same class of mistake as fabricating a zero
where a blank belongs.

#### `calc_totals` — `B23:C32`, all `#,##0.00` SAR

```
A_nom  Escalated Deterministic Base — Nominal      A_pv  ... PV
B_nom  Uncertainty Mean Shift — Nominal            B_pv  ... PV
C_nom  Mean-Basis Base Cost — Nominal              C_pv  ... PV
D_nom  Expected Risk / EMV — Nominal               D_pv  ... PV
E_nom  Analytical Mean Total — Nominal             E_pv  ... PV
```

#### `tblCalcYears` — `H:J`

| # | Column | Type | Format | Units |
|---|---|---|---|---|
| 1 | Project Index | integer | `0` | index, from 1 |
| 2 | Calendar Year | integer | `0` | year |
| 3 | Discount Factor | double | `0.000000` | dimensionless |

#### `tblCalcInflationFactors` — `M:P`

**Row rule LOCKED: one row per referenced profile per FACTOR year, spanning
`BaseYear … LastProjectYear` inclusive.**

Revision B's "required years" span began at `BaseYear + 1` and therefore omitted
the Base Year itself — so the factor of `1` that every calculation depends on
appeared nowhere in the audit, and when `BaseYear = StartYear` the first project
year's factor could not be explained from this table at all.

| # | Column | Type | Format | Units |
|---|---|---|---|---|
| 1 | Inflation Profile | text | `@` | key |
| 2 | Calendar Year | integer | `0` | year |
| 3 | Annual Rate | double | `0.00%` | rate |
| 4 | Cumulative Inflation Factor | double | `0.000000` | dimensionless |

| Year | Annual Rate | Cumulative Inflation Factor |
|---|---|---|
| `BaseYear` | **blank** — model-controlled audit output; no rate is required or read for the base year | **`1`** |
| `BaseYear + 1` … `LastProjectYear` | the resolved rate | the compounded factor |

The blank Base-Year rate is a *model-controlled output*, not user data, and is
distinct from the D4 blank-is-refusal rule, which governs **user-entered profiling
cells**. The Base-Year row carries no rate because none exists to carry.

When `BaseYear < StartYear` this table also exposes the **pre-project compounding
years**, so every project-year factor can be explained from the audit alone
without re-deriving anything.

#### `tblCalcFX` — `S:U`

| # | Column | Type | Format | Units |
|---|---|---|---|---|
| 1 | Currency | text | `@` | key |
| 2 | FX to SAR | double | `0.000000` | SAR per unit |
| 3 | Referenced By | integer | `0` | driver count |

#### `tblCalcDrivers` — `X:AR`, one row per identified Cost Line and Risk

Revision B's 16 columns could not reconstruct A, B, C and D from the audit rows,
and `Mean-Basis Nominal` was ambiguous on a Risk row. **No column now carries two
meanings depending on Driver Kind.** A field that does not apply to a kind is
**blank** — never zero, never reused.

| # | Column | Type | Format | Units | Cost Line | Risk |
|---|---|---|---|---|---|---|
| 1 | Permanent ID | text | `@` | key | yes | yes |
| 2 | Driver Kind | text | `@` | `Cost Line` / `Risk` | yes | yes |
| 3 | Distribution | text | `@` | — | yes | yes |
| 4 | Central Basis | text | `@` | `ML` / `Midpoint` | yes | yes |
| 5 | Currency | text | `@` | — | yes | yes |
| 6 | FX to SAR | double | `0.000000` | SAR per unit | yes | yes |
| 7 | Inflation Profile | text | `@` | — | yes | yes |
| 8 | Quantity | double | `#,##0.00` | units | yes | **blank** |
| 9 | Probability | double | `0.0%` | fraction | **blank** | yes |
| 10 | Central Value | double | `#,##0.00` | source currency | yes | **blank** |
| 11 | Mean Value | double | `#,##0.00` | source currency | yes | yes — expected severity |
| 12 | Knom | double | `0.000000` | SAR per source unit | yes | yes |
| 13 | Kpv | double | `0.000000` | SAR per source unit | yes | yes |
| 14 | Deterministic Nominal | double | `#,##0.00` | SAR | yes | **blank** |
| 15 | Deterministic PV | double | `#,##0.00` | SAR | yes | **blank** |
| 16 | Mean-Basis Nominal | double | `#,##0.00` | SAR | yes | **blank** |
| 17 | Mean-Basis PV | double | `#,##0.00` | SAR | yes | **blank** |
| 18 | Uncertainty Mean Shift Nominal | double | `#,##0.00` | SAR | yes | **blank** |
| 19 | Uncertainty Mean Shift PV | double | `#,##0.00` | SAR | yes | **blank** |
| 20 | Expected Risk Nominal | double | `#,##0.00` | SAR | **blank** | yes |
| 21 | Expected Risk PV | double | `#,##0.00` | SAR | **blank** | yes |

`Quantity = 1` and `Probability = 1` remain the multiplicative identities used by
the **in-memory** `DriverFactors` (§26). The **audit table** shows blank, so an
auditor is never shown a fabricated `1` the user did not enter.

**Reconstruction from the audit rows alone** — the property Revision B lacked:

```
A_nom = SUM(col 14)   A_pv = SUM(col 15)      over Cost Line rows
B_nom = SUM(col 18)   B_pv = SUM(col 19)      over Cost Line rows
C_nom = SUM(col 16)   C_pv = SUM(col 17)      over Cost Line rows
D_nom = SUM(col 20)   D_pv = SUM(col 21)      over Risk rows
E     = C + D
```

Every headline component is a plain column sum over rows of one kind. No
inference, no guessing.

#### `tblCalcAnnual` — `AU:BB`

**Calendar Year is added** so an annual audit row stands on its own, without a
join to `tblCalcYears` merely to learn which year it describes.

| # | Column | Type | Format | Units |
|---|---|---|---|---|
| 1 | Project Index | integer | `0` | index |
| 2 | **Calendar Year** | integer | `0` | year |
| 3 | Base Cost Nominal | double | `#,##0.00` | SAR |
| 4 | Expected Risk Nominal | double | `#,##0.00` | SAR |
| 5 | Total Nominal | double | `#,##0.00` | SAR |
| 6 | Base Cost PV | double | `#,##0.00` | SAR |
| 7 | Expected Risk PV | double | `#,##0.00` | SAR |
| 8 | Total PV | double | `#,##0.00` | SAR |

### 16.5 Update trigger and validation

Every block: **trigger** = `PCCM_Calculate`; **validation** = the numerical
prerequisites of §18 plus the identities of §15; **ownership** = Phase 5;
**units** as tabulated.

`_Calc` is a **written record of what the in-memory kernel computed**. Nothing
reads it back to compute anything else — that would recreate the worksheet
dependency the whole design exists to avoid. **No per-iteration data is ever
written here.**

---

## 17. VBA / numerical module boundaries

The hard rule: **mathematical functions must not read worksheet cells.**

| Module | Layer | Responsibility | Worksheets |
|---|---|---|---|
| `modCalcContract` | generated | Phase-5 constants projected from `calc_contract.yaml` | n/a |
| `modCalcResolve` | resolution | builds the referenced-currency and referenced-profile sets from the registers, then reads Setup, profiling grids, `tblFXRates`, `tblInflation` into plain numeric structures | **yes** |
| `modCalcFactors` | numerical | the safe arithmetic primitives (§19.1), `InflationFactors`, `DiscountFactors`, `BuildKnom`, `BuildKpv` | **no** |
| `modCalcAnalytical` | numerical | `TriangularMean`, `PertMean`, `UniformMean`, `DeterministicCentral`, `ExpectedRisk`, A–E accumulations, annual series | **no** |
| `modCalcFingerprint` | numerical | canonical encoding, per-driver digest, sort by ID, fold (§11.3) | **no** |
| `modCalcCheck` | validation | Phase-5 numerical prerequisites; reports, never repairs | **yes**, read only |
| `modCalcReport` | presentation | writes the `_Calc` blocks; reports refusal through the Phase-4 `modAppState` surface | **yes** |

`modCalcFingerprint` is numerical deliberately: the fingerprint must be computable
from resolved data alone, so Phase 6 can extend it without a worksheet.

**A static sweep enforces the boundary.** In `modCalcFactors`,
`modCalcAnalytical` and `modCalcFingerprint`, **none** of the following may
appear:

```
Application.        ThisWorkbook        ActiveWorkbook
Worksheets          Worksheet           Range
Cells               ListObject          ListObjects
modWorkbook.
```

`Application.` is on the list because of Revision C: it proposed reading
`Application.International(xlDecimalSeparator)` inside the numeric encoder, which
would have made the fingerprint kernel depend on a live Excel Application while
the plan claimed it was pure. The separator is now passed in (§11.3), and the
sweep makes the claim enforceable rather than aspirational.

**The public numerical routines must behave as pure functions of their
arguments** — same inputs, same outputs, on any machine, with no Excel present.

The resolver, check and report layers may use Excel objects freely; the boundary
is what makes the kernel reusable by Phase 6 inside an iteration loop.

A second sweep extends the Phase-4 `On Error Resume Next` whitelist across the
Phase-5 modules, so the only error handlers are the documented safe primitives of
§19.1. Both are mechanically checkable on Linux, permanent, in the established
style.

### No user-facing Calculate button

**LOCKED for Phase 5: `PCCM_Calculate` — yes. Calculate button — no.**

A standalone Calculate button was not part of the locked Dashboard command set,
and adding user-facing workflow before Results, Model Check and Run Simulation
exist would clutter the UI. The Windows harness invokes `PCCM_Calculate` directly
through `Application.Run`. Later phases call the same orchestration from Run
Check / Run Simulation / output-refresh pathways.

**The workbook keeps exactly the five Phase-4 buttons.** Gate B proves it (§24).

---

## 18. Validity and failure behaviour

Phase 5 refuses; it never manufactures a value and never repairs.

### Structural prerequisites — owned by Phase 4, invoked not duplicated

`STRUCTURE CHANGE PENDING` · no timeline applied · non-empty
`modStructuralCheck.ValidateStructure()` · duplicate permanent ID · orphan
profiling row · profiling ID missing from register · register ID missing from
profiling. Phase 5 calls the existing gate and quotes its report.

### Phase-5 numerical prerequisites — owned here

| Check | Refusal |
|---|---|
| Base Year > Start Year | yes |
| Discount Rate blank / non-numeric | yes |
| `1 + r ≤ 0` (D3) | yes |
| **referenced** currency missing from `tblFXRates` | yes |
| **referenced** currency duplicated in `tblFXRates` | yes |
| **referenced** FX rate `≤ 0`, blank or non-numeric | yes |
| SAR ≠ 1 | yes (global invariant) |
| unreferenced currency invalid / duplicated | **no — must not block** |
| **referenced** inflation profile missing | yes |
| required inflation year blank for a **referenced** profile | yes |
| inflation rate non-numeric | yes |
| `1 + rate ≤ 0` (D2) | yes |
| unreferenced profile incomplete | **no — must not block** |
| profiling cell non-numeric | yes |
| **any blank profiling cell** on an identified driver (D4) | yes |
| profiling sum ≠ 100% within tolerance | yes |
| distribution missing or not one of the three | yes |
| `Min ≤ ML ≤ Max` violated (Triangular, Beta-PERT) | yes |
| `Min ≤ Max` violated (Uniform) | yes |
| Uniform with a populated ML (D1) | **no — succeeds, ML ignored** |
| Quantity missing or non-numeric | yes |
| **Quantity ≤ 0** | **yes** — see below |
| Probability missing, non-numeric, or outside `[0,1]` | yes |
| any numerical result non-finite (§19) | yes |

**Quantity positivity.** The locked driver contract says:

> `note: "Deterministic. Fixed during simulation. No positivity rule here; that is
> a Model Check rule."`

That governs the **cell rule**, and no cell rule is added. But Phase 5 must be able
to refuse an invalid numerical calculation independently of a Model Check UI that
does not exist, so **`Quantity` must be numeric and `> 0`** is a Phase-5 numerical
prerequisite. The same predicate is later aggregated by Model Check — one
predicate, two consumers, no duplicated UI logic.

**No positivity rule is invented for Min / ML / Max**, because no locked contract
requires one. Only the three-point ordering rules apply.

Refusals are reported through the Phase-4 `modAppState` result surface, so there is
one refusal mechanism, not two.

---

## 19. Controlled arithmetic, overflow and underflow

Revision B applied a finiteness predicate **after** each stage. **That is not
sufficient.** In VBA

```vb
x = a * b
```

raises `Overflow` (error 6) **before** any value exists, so `IsUsableDouble(x)`
never executes. A post-hoc predicate cannot catch a failure that prevents its own
operand from being produced.

### 19.1 Safe arithmetic primitives

Phase 5 defines pure numerical primitives that convert representational failure
into the controlled Phase-5 numerical-range refusal:

```
SafeAdd(a, b, ByRef result)         -> Boolean
SafeSubtract(a, b, ByRef result)    -> Boolean
SafeMultiply(a, b, ByRef result)    -> Boolean
SafeDivide(a, b, ByRef result)      -> Boolean
SafeAccumulate(ByRef acc, term)     -> Boolean
```

Each returns `False` on representational failure and leaves `result` untouched;
callers propagate the failure with the stage, driver, profile or year that caused
it. They live in `modCalcFactors`, are worksheet-free, and are covered by their own
Gate-A tests.

**Error handling is tightly scoped.** A primitive may install a handler around its
single arithmetic expression to trap VBA `Overflow` and return a structured
failure. **No broad `On Error Resume Next` is permitted anywhere** — the Phase-4
whitelist test is extended to cover the Phase-5 modules, so a handler outside the
documented primitives fails Gate A. No uncontrolled runtime error escapes, and
**no overflow becomes a fabricated zero**.

`SafeDivide` additionally refuses a zero divisor rather than relying on error 11,
and reports underflow of a quotient to exactly zero where the operands were both
non-zero (§19.3).

### 19.2 Numerically stable formulas — mandatory, not optional

Where a mathematically equivalent expression avoids an intermediate that can
overflow while the final result is representable, **the stable form is required**.

| Quantity | Forbidden naive form | Required stable form |
|---|---|---|
| Triangular mean | `(Min + ML + Max) / 3` | `Min/3 + ML/3 + Max/3` |
| Beta-PERT mean, λ = 4 | `(Min + 4·ML + Max) / 6` | `Min/6 + ML·(2/3) + Max/6` |
| Uniform central / mean (midpoint) | `(Min + Max) / 2` | `Min/2 + Max/2` |

Each division is applied **before** accumulation, and the accumulation itself uses
`SafeAccumulate`.

This is not hypothetical. With `Min = ML = Max = 1e308`:

| Form | Result |
|---|---|
| naive `(Min + ML + Max) / 3` | numerator `3e308` → **overflow**, no result |
| stable `Min/3 + ML/3 + Max/3` | **`1e308`** — exactly the correct mean |
| naive `(Min + 4·ML + Max) / 6` | numerator `6e308` → **overflow** |
| stable `Min/6 + ML·(2/3) + Max/6` | **`1e308`** |
| naive `(Min + Max)/2` with both `1.5e308` | numerator `3e308` → **overflow** |
| stable `Min/2 + Max/2` | **`1.5e308`** |

The same policy governs `Knom` / `Kpv` products, per-driver contributions, annual
contributions and the A–E accumulators: every multiplication and every addition
goes through the safe primitives, and no naive intermediate is formed where a
stable equivalent exists.

> **Erratum C2 applies to this section.** The stable forms above are the required
> ordinary path and are unchanged. They are, however, not sufficient on their own
> to reach the objective stated at the top of §19.2: a signed sum can still lose a
> representable total to its partial sums, and a convex statistic can still be
> refused for an internal averaging step, even in the stable form. Both are
> repaired by the two-tier rules recorded in §0 Erratum C2 — canonical order
> first, rescue only where canonical order produced nothing, and a
> zero-uncertainty distribution returned exactly, with each rescue judged against
> the exact mathematical value of its Double inputs rather than against a
> re-associated evaluation of them. The algorithms are specified for VBA in
> `docs/phase5_gate_a_step2.md` §18–§20.

### 19.3 Inflation and discount — iterative, with overflow AND underflow detection

Both are built **iteratively**, one year at a time, never by forming a power:

```
infl(BaseYear)   = 1
infl(Y)          = SafeMultiply( infl(Y-1), 1 + rate_Y )

disc(1)          = 1
disc(t)          = SafeDivide( disc(t-1), 1 + r )
```

Iterating is itself a protection: `(1 + r)^(t-1)` can overflow as an intermediate
even when the reciprocal is perfectly representable.

**Overflow** is detected at the exact step that causes it. With a rate of `1e150`:

```
year 1 cumulative factor 1.000000e+150
year 2 cumulative factor 1.000000e+300
year 3 OVERFLOW   -> refusal names the profile and calendar year 3
```

**Underflow to exactly zero is detected too, and is the more dangerous case
because it is silent.** A discount factor that collapses to `0` would quietly
delete a year's entire PV contribution with no error anywhere. With `r = 1e10`:

```
t = 2   disc = 1.000000e-10
t = 30  disc = 1.000000e-290
t = 33  disc = 9.999889e-321      (subnormal, still non-zero)
t = 34  disc = 0                  -> REFUSED: discount factor underflowed to zero
```

The rule: while `1 + r > 0`, a discount factor is **required to remain strictly
positive**. Collapse to zero is refused with a numerical-range message naming the
project year, never accepted as "a very small number".

The same underflow rule applies to cumulative inflation factors when rates are
close to `−1`.

### 19.4 Stage coverage

| Stage | Guard |
|---|---|
| inflation compounding | `SafeMultiply` per year; overflow and underflow-to-zero both refuse, naming profile and calendar year |
| discount factor | `SafeDivide` per year; overflow and underflow-to-zero both refuse, naming the project year |
| central / expected value | stable formulas plus `SafeAccumulate` |
| `Knom` / `Kpv` | `SafeMultiply` + `SafeAccumulate`, checked **during** accumulation |
| driver contribution | `SafeMultiply`, per driver |
| annual accumulators | `SafeAccumulate`, per year and per series |
| A / B / C / D / E | `SafeAccumulate`, checked after **each driver**, not only at the end |

Checking **during** accumulation rather than at the end is deliberate: it names the
driver, profile or year responsible, instead of reporting that a total is
infinite.

### 19.5 The predicate still exists

`IsUsableDouble(v)` — not NaN, not ±∞, `|v| ≤ 1.7976931348623157e308` — remains,
as a **final** assertion on every value before it is written to `_Calc` or enters
the fingerprint. It is now a backstop, not the mechanism.

A genuinely unrepresentable result is refused with a specific numerical-range
message naming the stage and the input. **No arbitrary business cap is invented to
avoid implementing safe arithmetic.**

---

## 20. `calc_contract.yaml` — scoped authority

Accepted, with its authority narrowly bounded so no duplicate source of truth is
created.

**It owns:** `_Calc` physical layout and the **exact anchors of §16.3** · block
and table names · column schemas · labels · display formats · units · numerical
tolerances and the conditioning-scale coefficients · calculation-state labels ·
**`FP_VERSION`** · reserved-cell declarations.

**It must NOT restate:** driver schemas (`driver_contract.yaml`) · the
distribution list (`input_contract.yaml`, `tblDistributions`) · timeline
structural limits (`structure_contract.yaml`) · the FX convention
(`input_contract.yaml`) · permanent-ID rules (`structure_contract.yaml`) ·
Phase-4 structure rules. These are **referenced or projected** from the existing
authorities, and the loader asserts the projection matches.

**It must NOT own the hash mathematics.** `FP_BASE`, `FP_MOD_1`, `FP_MOD_2`,
`FP_INIT_1`, `FP_INIT_2`, the recurrence, the canonical encoding and the section
order are **not** contract data. Hand-maintaining primes and a recurrence in a YAML
file *and* in two implementations is three copies of one algorithm, and they will
drift. Their single source is:

```
this document (§11.3–§11.7)  +  the tested Python/VBA implementations
                             +  the fixed test vectors 50B6EB0E26857EA7 and §11.7
```

The contract carries only `FP_VERSION`, because that *is* workbook-representable
state: it is written into `calc_state` and must be comparable across saved files.

**Mathematical semantics are not defined in YAML** for the same reason. A formula
written once in YAML and again in VBA/Python is two sources of truth that will
diverge. The division is:

```
this document + the tested numerical oracle   define the numerical semantics
calc_contract.yaml                            defines their workbook representation
```

If a constant ever must appear in both places, **one generates the other** — the
emitter projects it into `modCalcContract.bas`, exactly as Phases 1–4 already do.

---

## 21. Gate A — Linux / static

### 21.0 The proof split — LOCKED (erratum E3)

**No VBA is executed on Linux.** There is no VBA interpreter, no Excel and no
`AscW` on the Gate-A host, so no Gate-A test may claim — in an assertion, a test
name, a document or a report — that VBA executed, that VBA produced a fingerprint,
or that VBA arithmetic was observed.

| Gate | Host | What is actually proven |
|---|---|---|
| **Gate A** | Linux / Python | the **Python numerical and fingerprint oracle**; fixed literal vectors asserted independently of the implementation under test; **generated / source / static conformance** of the VBA — text, structure, declarations, forbidden constructs, projected constants. **No execution of VBA.** |
| **Gate B** | Windows / real Excel | **actual VBA execution**; actual canonical numeric encoding under a real locale; actual UTF-16 / `AscW` behaviour including sign normalisation; the actual `Double`-only reducer; actual end-to-end fingerprint parity |

Gate-A **static** validation of VBA source is **not weakened** by this split. Every
mechanical sweep, source rule and generated-artifact assertion stands, and new ones
may be added; what is forbidden is describing any of them as runtime proof.

Where a Gate-A test mirrors VBA semantics in Python — the `Double`-only reducer of
§11.5 is the leading example — the test proves that **the reference semantics are
self-consistent**, not that VBA implements them. The VBA side of that claim is
Gate B's, and §24.1 makes it a direct, vector-level requirement rather than an
inference from one golden digest.

### 21.1 Gate-A deliverables

1. **Pure-Python numerical oracle** — `builder/pccm_builder/calc_oracle.py`,
   implementing §5–§15 and the §11 fingerprint independently, in the
   `structure_oracle.py` pattern: it defines the semantics the VBA must match
   **and** emits the expected values the Windows harness asserts, so the two
   cannot drift.
2. **Golden-case tests** — §23, every value hand-derived and written as a literal.
3. **Oracle-independence test** — below.
4. **Contract validation** for `calc_contract.yaml`, including the authority
   boundary (§20) and the exact-anchor assertions of §16.3.
5. **Source sweeps** — the existing mechanical ones extended to the new modules.
6. **Post-build verification** extended with the `_Calc` block layout.

### Revision-C specific Gate-A proofs

| Proof | How |
|---|---|
| length-prefixed serialisation is collision-free | the eight §11.7 probes — strings containing `:`, `U+001F`, `U+0000`, `U+000A` — must yield eight distinct digests, asserted against the literals |
| UTF-16 code-unit parity | Python encodes `utf-16-le` and reads 16-bit units; test vectors include a non-BMP character, proving it contributes **two** units, and a character above `U+7FFF`, proving the `AscW` sign normalisation |
| numeric encoding is locale-invariant | the ten §11.3 literals asserted exactly; plus a test that the **Python reference** normalisation step maps a `,` decimal separator to `.` before hashing. This proves the **reference normalisation semantics**; it does **not** prove VBA `Format`/`Str` runtime behaviour under a comma locale, which is reserved for Gate B (§21.0, §24.1) |
| exact hash constants and digest | `FP_BASE`, both moduli, both initial states asserted as literals; the §11.6 stream asserted at **366 code units** and its digest at **`50B6EB0E26857EA7`** — produced by the **Python** reference implementation (§21.0) |
| **no native `Mod` or `\` in the reduction** | a **static source rule** over the **executable** hash-recurrence code of `modCalcFingerprint`, not a whole-file word ban — the word "modulus" must remain usable in prose and identifiers. Static conformance only; the VBA is not executed (§21.0) |
| **the reference `Double`-only reduction is exact** | the four §11.5 reduction vectors, every `x` of which exceeds `Long.MaxValue`; plus a randomised sweep asserting the **Python mirror of the locked VBA reducer** equals exact integer `%` for both moduli, then the full stream re-digested through that mirror to `50B6EB0E26857EA7`. The mirror is a reference oracle, **not** a VBA execution: parity with real VBA arithmetic is proven at Gate B (§21.0, §24.1) |
| row order excluded | the same drivers in reversed order produce an identical fingerprint |
| Uniform ML excluded | two Uniform drivers differing **only** in ML produce an identical fingerprint |
| stable formulas avoid naive overflow | the `1e308` and `1.5e308` cases of §19.2: naive overflows, stable returns the exact mean |
| true overflow refuses | a genuinely unrepresentable result yields a structured refusal, never an uncontrolled error and never zero |
| discount underflow refuses | the `r = 1e10` sequence: refusal at the project year where the factor reaches exactly zero |
| exact anchors do not overlap | every §16.3 band checked pairwise against its declared schema width, and against rows 1–11 and `C10:C11` |
| driver audit reconstructs A/B/C/D | the column sums of §16.4 equal `calc_totals` for a multi-driver fixture |
| annual schema includes Calendar Year | asserted in the emitted layout and in the oracle fixture |
| safe primitives are the only error handlers | the `On Error Resume Next` whitelist extended; a handler outside the documented primitives fails |

### Golden oracle independence

```
hand-derived literals  ->(verify)->  Python oracle  ->(emit)->  phase5_cases.json  ->(assert)->  Windows/VBA
```

A static test asserts that **every expected value emitted into
`phase5_cases.json` equals its separately hard-coded hand-derived literal**. The
JSON must not become self-validating merely because the oracle produced it. The
same rule covers the fingerprint literal.

For refusal cases the harness verifies: the specific refusal **class/message** ·
**no partial analytical totals** written · the previous successful snapshot **not
overwritten** · `calc_state` reflecting the refused state.

Gate A ends with a source review, exactly as Phase 4 did.

---

## 22. Simulation-artefact acceptance wording — corrected

Revision A's *"no Monte Carlo artefact exists anywhere in the phase"* was too
broad: the frozen workbook already legitimately contains
`inpMonteCarloIterations`, `inpRandomSeed`, `inpSelectedConfidenceLevel` and the
`_SimData` sheet, all from the locked architecture.

**Replacement criterion:**

> Phase 5 introduces no RNG implementation, no sampling implementation and no
> simulation output, and makes no use of Iterations, Random Seed or Selected
> Confidence Level in any analytical calculation or in the Phase-5 calculation
> fingerprint. `_SimData` remains unchanged and unused.

**Sweeps test dependencies, not vocabulary.** Fragile word-matching on `Sample` or
`Iteration` in a comment is replaced by:

- no Phase-5 module reads `inpMonteCarloIterations`, `inpRandomSeed` or
  `inpSelectedConfidenceLevel` — checked against the **generated constant names**,
  in code with comments and strings stripped;
- no Phase-5 module references `_SimData` or `shSimData`;
- no Phase-5 module calls `Rnd`, `Randomize` or `Timer`;
- the fingerprint section list (§11.1) contains none of the three excluded inputs
  — asserted against the oracle's field list, not against prose.

---

## 23. Golden and refusal matrix

Cases 1–15 are unchanged from Revision A and remain hand-derived. Cases 16–25 lock
the behaviour required by D1–D4 and §18–§19.

Shared unless overridden: Triangular `Min 80 / ML 100 / Max 150`, `Quantity 10`,
SAR, one profile at 100%.

| # | Case | Setup | Hand-derived expected |
|---|---|---|---|
| 1 | SAR, no inflation, one project year | Base 2026, Start 2026, Dur 1; inflation span empty (`2027 > 2026`); `r = 10%` | `infl = 1`, `disc₁ = 1`, `Knom = Kpv = 1`; `A = 1000`, `C = 1100`, `B = 100` |
| 2 | foreign currency | as 1, USD `FX = 3.75`, unit cost 100, Qty 4 | `Knom = 3.75`; `A = 100×4×3.75 = 1500 SAR` |
| 3 | multi-year profiling, compounded inflation | Base 2026, Start 2027, Dur 3; rates 5%; profile `20/50/30` | `f = 1.05, 1.1025, 1.157625`; `Knom = 0.21+0.55125+0.3472875 = ` **`1.1085375`**; `A_nom = ` **`1108.5375`** |
| 4 | PV across multiple years | 3 with `r = 10%` | `Kpv = 0.21+0.501136363636+0.287014462810 = ` **`0.998150826446`**; `A_pv = ` **`998.150826446`**; `C_pv = ` **`1097.965909091`** |
| 5 | Triangular deterministic vs mean | `80/100/150` | central `100`, mean **`110`**; on case 3 `C_nom = ` **`1219.39125`**, `B_nom = ` **`110.85375`** |
| 6 | Beta-PERT deterministic vs mean | `80/100/150`, λ=4 | central `100`, mean `(80+400+150)/6 = ` **`105`** |
| 7 | Uniform midpoint = mean | `Min 80 / Max 150` | central `115`, mean `115`; **`B = 0`** for this driver |
| 8 | risk EMV, `P < 1` | `P = 30%`, severity `100/200/450`, case-1 factors | mean severity `250`; `D = ` **`75`** |
| 9 | multi-year risk profile | risk of 8 on case-3 factors | `D_nom = ` **`83.1403125`**; `D_pv = ` **`74.8613119835`** |
| 10 | Base Year = Start Year | Base 2027, Start 2027, Dur 2; rate 2028 = 5% | `infl(2027) = 1`, `infl(2028) = 1.05` |
| 11 | Base Year earlier than Start Year | case 3 | project year 1 (2027) already carries `1.05` |
| 12 | zero inflation | rates `0%` | every `infl = 1`; `Knom = FX` |
| 13 | negative but valid inflation | rate `−2%`, three years | `0.98, 0.9604, 0.941192` (D2) |
| 14 | blank required inflation rate | case 3, 2028 blank | **refusal**, naming profile and year 2028 |
| 15 | profile sum ≠ 100% | case 3, profile `20/50/20` | **refusal**, naming the permanent ID and the sum |
| **16** | **Quantity = 0** | case 1 with `Qty 0` | **refusal** (§18) |
| **17** | **Quantity < 0** | case 1 with `Qty −5` | **refusal** |
| **18** | **Discount Rate = −100%** | `r = −1` | **refusal** — `1 + r = 0` (D3) |
| **19** | **Discount Rate negative but > −100%** | case 3 with `r = −5%` | **accepted**; `disc = 1, 1/0.95, 1/0.9025`, so `A_pv > A_nom` — correct, and why I5 is not a gate (§15) |
| **20** | **Inflation Rate = −100%** | one required year at `−100%` | **refusal** — `1 + rate = 0` (D2) |
| **21** | **Inflation Rate negative but > −100%** | case 13 | **accepted** |
| **22** | **Uniform with populated ML** | `Min 80 / ML 999 / Max 150` | **accepted**; central = mean = `115`; ML ignored **and excluded from the fingerprint** (D1) |
| **23** | **100%-summing profile containing a blank** | Dur 3, weights `50% / blank / 50%` | **refusal** (D4) — sums to 100% and is still refused |
| **24** | **Double overflow** | inflation `1e300%` compounded over several years, or an extreme unit cost | **controlled refusal** with a numerical-range message naming the stage; no uncontrolled VBA error, no fabricated zero (§19) |
| **25** | **unreferenced incomplete FX / Config row** | valid SAR-only model + a duplicate, blank-rate `EUR` row referenced by nothing | **does NOT block**; calculation succeeds and the fingerprint is unaffected (§8) |

Cases 14–18, 20, 23 and 24 assert a **refusal**, not a number, and must produce no
partial result. Cases 19, 21, 22 and 25 assert **acceptance** where a naive
implementation would over-block.

Cases 16–25 need not each be a separate full fixture; the Gate-B matrix may
exercise them compactly. Their **behaviour is locked** regardless.

### Revision-C additions

| # | Case | Setup | Expected |
|---|---|---|---|
| **26** | **fingerprint reference vector** | golden case 1 | canonical stream of exactly **366** UTF-16 code units; fingerprint **`50B6EB0E26857EA7`** — identical in Python, in VBA, and on real Excel |
| **27** | **delimiter-hostile field content** | the eight §11.7 probes | eight **distinct** digests, asserted against the §11.7 literals |
| **28** | **naive-overflow, representable result** | `Min = ML = Max = 1e308`, Triangular and Beta-PERT; midpoint with both `1.5e308` | **accepted**; stable forms return `1e308`, `1e308`, `1.5e308`. A naive implementation overflows and is thereby detected |
| **29** | **discount factor underflow** | `r = 1e10`, duration ≥ 34 | **controlled refusal** at project year 34, where the factor reaches exactly zero — never silently accepted |
| **30** | **cancellation-heavy reconciliation** | large positive and negative unit costs whose net is near zero, all representable | identities I1–I4 **hold**; the conditioning scale of §15 keeps the tolerance proportional to the arithmetic performed, not to the near-zero net |
| **31** | **Base-Year factor row** | `Base 2026, Start 2028, Dur 3` | `tblCalcInflationFactors` contains a `2026` row with **blank rate and cumulative factor `1`**, plus the pre-project rows `2027`, `2028` |
| **32** | **status reverts to CURRENT** | calculate → break an input → refuse → restore the input exactly → query | derived **`CURRENT`** with no recalculation, while the attempt axis still reads `REFUSED` (§11.11, §25.5) |
| **33** | **mid-write failure** | injected failure after `tblCalcDrivers` is mutated | full logical rollback in two observable moments (erratum E1). **After rollback, before failure metadata:** `C13:C20` restored **exactly**, all eight cells. **Final observable state:** `C13:C16` = the previous successful snapshot restored exactly · `C17:C20` = **new** failed-attempt and derived-status metadata · `C23:C32` = the previous totals restored exactly · all five analytical ListObjects restored exactly. **The final acceptance comparison is `C13:C16` + `C23:C32` + the five analytical ListObjects — never all of `C13:C20`,** which would assert the failure was not recorded. No mixed state; attempt result `FAILED`; derived status **`STALE`**, not forced from the attempt (§12.4, §12.5, §25.7) |
| **34** | **invalid input, no Calculate attempted** | valid success, then break an input and query without calculating | derived **`INVALID`**, attempt result still **`SUCCESS`** — the two axes moving independently (§25.1 row 3) |
| **35** | **locale-separator injection** | canonical numeric encoder given `,` as the decimal separator | output **byte-identical** to the `.` case; the encoder is a pure function of its arguments (§11.3) |
| **36** | **reduction beyond `Long`** | `h = 2147483646, u = 65535` for `FP_MOD_1`, and the three other §11.5 vectors | `x = 281,320,423,161` — **approximately 131 times the signed-`Long` maximum**; exactly, `131 × 2,147,483,647 = 281,320,357,757` and `281,320,423,161 = 131 × Long.MaxValue + 65,404` (erratum E2) — reduces to **`65404`**; the `Double`-only reducer equals exact integer `%` for both moduli. A native `Mod` implementation fails here. Reduction vectors and expected remainders unchanged |
| **37** | **failure at the commit boundary** | injected failure at the final `C13:C20` assignment | the same rollback path as a mid-table failure; `C13:C16` and all analytical blocks restored; attempt `FAILED`; status derived independently (§12.3–§12.5) |

Cases 26–27, 31 and 35 assert **format and audit content**; 28, 32 and 34 assert
**acceptance and correct status** where a naive implementation would fail,
over-block or conflate the two axes; 29 and 33 assert **controlled refusal and
rollback**.

---

## 24. Gate B — real Windows / Excel

Extends the accepted `phase4_functional_test.ps1` matrix; it does not replace it.

### 24.1 Direct Windows vector coverage — LOCKED REQUIREMENT, IMPLEMENTED LATER

**Recorded now so it cannot be forgotten; deliberately not implemented in Gate-A
Step 1.**

The golden-case fingerprint parity of §25.6 is necessary but **not sufficient**. A
single end-to-end digest can hide a compensating pair of encoder and reducer
defects, and it exercises none of the extreme vectors the design was built around.
Real Windows / VBA must therefore exercise the **canonical encoder** and the
**reducer** *directly*, against the **complete locked vector set**:

| Vector group | What real VBA must reproduce |
|---|---|
| **the ten numeric canonical encodings** (§11.3) | `0`, `-0`, `1`, `-1`, `0.1`, `1e-20`, `1e+20`, `0.1+0.2`, the maximum usable `Double`, and the minimum subnormal **where it is safely constructible in VBA** — each asserted against its locked literal |
| **decimal-separator injection** | the canonical encoder given `.` and given `,`, producing identical output on a real Windows locale — the runtime half of the proof Gate A cannot make (§21.0) |
| **all four locked reduction vectors** (§11.5) | `(2147483647, 2147483646, 65535) → 65404` · `(2147483629, 2147483628, 65535) → 65404` · `(2147483647, 1234567890, 41) → 667120106` · `(2147483629, 1234567890, 41) → 667121456`, each computed by the real `Double`-only reducer |
| **UTF-16 code-unit handling** | a code unit above `U+7FFF`, proving `AscW` sign normalisation on target; and a **non-BMP** character, proving it contributes **two** surrogate code units and that length prefixes count **UTF-16 units** |
| **the complete reference stream** | the §11.6 stream at **366 code units** digesting to **`50B6EB0E26857EA7`** on real Excel |

**Preferred Gate-B design — a transient, test-only VBA diagnostic module.** It is
imported **only** into the disposable Windows harness working copy, and it may call
the numerical fingerprint helpers directly so the encoder and reducer are exercised
without an analytical fixture. Its constraints are absolute:

- it **must not** enter the Stage-B production manifest;
- it **must not** persist in the accepted workbook;
- it **must not** create a button or any user-facing surface;
- it **must not** add a `PCCM_` production API entry point (erratum E4 — the
  endpoint surface stays at six).

**This module is NOT part of Gate-A Step 1 and must not be written there.** It is
recorded here as a Gate-B acceptance requirement only.

### Additive expectations

| Claim | How |
|---|---|
| all 8 original Phase-4 modules still persist | the reopen-and-verify step reads the module list from the manifest, which now declares Phase-4 **and** Phase-5 modules; the Phase-4 eight are asserted present **by name** |
| all Phase-5 modules persist | same list, new names asserted present by name |
| all 5 Phase-4 command buttons still persist | asserted by shape name from the manifest, exactly as today |
| **no Phase-5 button was added** | the manifest's button count is asserted `= 5`, and the sheet shape inventory is asserted to contain **no** shape whose `OnAction` is `PCCM_Calculate` (§17) |
| the VBA project still compiles | `A1` unchanged — still the first `Application.Run` of the run |

**The full 35/35 Phase-4 functional matrix remains mandatory** and must pass before
any Phase-5 scenario is considered accepted.

### Phase-5 functional coverage

FX resolution (foreign currency and the SAR identity) · inflation compounding
(`Base = Start` and `Base < Start`) · discount factors at indices 1, 2, 3 ·
profiling factor application by permanent ID · deterministic base (Nominal, PV) ·
mean-basis base (Nominal, PV) · expected risk (Nominal, PV) · all six annual
series · identities I1–I5 asserted in the workbook · refusal on every §18
numerical prerequisite · refusal on `STRUCTURE CHANGE PENDING` · clean shutdown
and clean transient COM release.

### Revision-C specific Gate-B proofs

| Proof | Detail |
|---|---|
| **fingerprint parity on real VBA** | `PCCM_CurrentInputFingerprint()` on the golden-case-1 fixture must equal the literal **`50B6EB0E26857EA7`** — the same value Python produces. This is the one assertion that proves the two implementations agree on real Excel, including `AscW` sign normalisation and locale-invariant numeric formatting |
| **the full status matrix** | all six rows of §25.1, each asserting derived status, attempt result, attempt detail and snapshot state — including row 3 (invalid input, no Calculate) and row 6 (rollback must not report `CURRENT`) |
| **status reverts to CURRENT** | the §23 case 32 sequence, driven end to end (§25.5) |
| **refusal preserves the snapshot** | break an input, `PCCM_Calculate`; assert derived status `INVALID`, attempt result `REFUSED` with a reason, and — comparing the two groups of §12.2 **separately** — that `calc_totals`, all five tables and `C13:C16` are unchanged while `C17:C20` has changed (§25.4) |
| **mid-write failure and full rollback** | inject a failure after one or more `_Calc` blocks have been mutated; assert previous totals, driver rows, annual rows, `C13:C16` and `C23:C32` all restored, no mixed snapshot, Excel application state restored. Uses the Phase-4 `FailPointCheck` mechanism already proven on target (§25.7) |
| **commit-boundary failure** | inject a failure at the final `C13:C20` commit assignment; the **same** rollback path must apply, proving the commit attempt is inside the transaction and not outside it (§12.3, §25.7) |
| **reduction parity on real VBA** | the fingerprint-parity case exercises the `Double`-only reducer end to end on Excel; a `Mod`-based implementation cannot produce `50B6EB0E26857EA7` because it would overflow `Long` first (§11.5) |
| **Base-Year factor visible** | `tblCalcInflationFactors` contains the `BaseYear` row with blank rate and cumulative factor `1` (§23 case 31) |
| **cancellation-heavy reconciliation** | a fixture with large offsetting contributions and a near-zero net; identities must hold (§23 case 30) |

Every expected value comes from `build/phase5_cases.json`. **The harness asserts
every calculated value; the user inspects no cells manually.**

Phase-4 harness disciplines carry over unchanged and are non-negotiable:
caller-side `@(...)`, one pipeline object per row, container factories emitted
non-enumerated, `catch` attached to its `try`, keyed-only fixtures, failure-safe
cleanup, per-scenario clean-structure prerequisites, `$excelIdentity`.

---

## 25. Gate-B status and fingerprint matrix

An explicit Windows functional oracle for the fingerprint and for the two status
axes, using the **same canonical semantics intended for later simulation reuse**.

### 25.1 The locked status matrix

Every row asserts **all four** accessors, because the whole point of §11.8 is that
the axes move independently. `snapshot` means the four `C13:C16` cells plus the
five analytical tables and `calc_totals`.

| # | Situation | Derived Status | Last Attempt Result | Attempt Detail | Snapshot |
|---|---|---|---|---|---|
| 1 | successful calculation, nothing touched | `CURRENT` | `SUCCESS` | blank | new |
| 2 | a **valid** input changed, no Calculate | `STALE` | `SUCCESS` *(unchanged)* | blank | unchanged |
| 3 | an **invalid** input, no Calculate yet | `INVALID` | `SUCCESS` *(unchanged)* | blank | unchanged |
| 4 | invalid input, `PCCM_Calculate` attempted | `INVALID` | `REFUSED` | specific refusal | **unchanged** |
| 5 | prior successful input restored **exactly**, no Calculate | `CURRENT` | `REFUSED` *(unchanged — history)* | still readable | unchanged |
| 6 | injected **write** failure while current inputs differ from the prior success, after rollback | `STALE` | `FAILED` | write/verification detail | **restored** |

**Row 3 is the one Revision C could not express.** An input goes bad *without*
anyone pressing Calculate: the status axis moves to `INVALID` while the attempt
axis still correctly reports the last real attempt as `SUCCESS`. Nothing is
overwritten, and nothing is invented.

**Row 6 is the most important.** A successful rollback must **not** make the new
inputs look `CURRENT`. The restored snapshot describes the *old* inputs; the
current inputs describe the calculation that failed. `STALE` is the only honest
answer, and it is derived independently of the fact that the attempt was `FAILED`.

### 25.2 Primary staleness sequence

1. establish a valid analytical fixture;
2. `PCCM_Calculate`;
3. assert `PCCM_CalculationStatus() = CURRENT` and
   `PCCM_CalculationAttemptResult() = SUCCESS`;
4. capture headline values **and** `PCCM_CalculationFingerprint()`;
5. change Quantity **or** one profiling weight, **without** touching the timeline
   and **without** recalculating;
6. assert status `STALE`, attempt result still `SUCCESS` *(matrix row 2)*;
7. assert `PCCM_CalculationFingerprint()` is **still the old value** while
   `PCCM_CurrentInputFingerprint()` has changed;
8. `PCCM_Calculate`;
9. assert status `CURRENT`, attempt result `SUCCESS`;
10. assert the stored fingerprint **changed**;
11. assert the affected analytical value changed **to the oracle value**.

### 25.3 Non-staleness proofs

Each must leave status `CURRENT`, attempt result `SUCCESS`, and the stored
fingerprint unchanged:

- changing a **Description**;
- changing **row order** (a real `ListObject.Sort`, as `B2`/`K2` already do);
- changing **Selected Confidence Level**;
- changing an **unreferenced** FX row or Config value.

### 25.4 Refusal proof — matrix row 4

Make an input invalid and run `PCCM_Calculate`. Assert:

```
PCCM_CalculationStatus()         = INVALID          (NOT "REFUSED")
PCCM_CalculationAttemptResult()  = REFUSED
PCCM_CalculationAttemptDetail()  = the specific refusal message
```

and then, comparing the two groups of §12.2 **separately**:

- `calc_totals`, all five tables and `C13:C16` are **unchanged**;
- `C17:C20` **has** changed, as it must.

No partial analytical totals were written.

### 25.5 Revert-to-CURRENT proof — matrix row 5

16. from the refused state above, restore the changed input **exactly** to the
    value it held at the last successful calculation;
17. query status **without** recalculating;
18. assert **`CURRENT`** — the inputs resolve and their fingerprint equals the last
    successful fingerprint, so a historical refusal must not keep the workbook
    marked stale;
19. assert `PCCM_CalculationAttemptResult()` is **still `REFUSED`** and the detail
    is still readable, proving the attempt axis was not silently reset;
20. assert the two axes disagreed — a status of `CURRENT` alongside an attempt of
    `REFUSED` is the orthogonality made visible.

### 25.6 Fingerprint parity proof

On the golden-case-1 fixture, `PCCM_CurrentInputFingerprint()` must return the
literal **`50B6EB0E26857EA7`** — the same value the Python oracle produces. This
is the assertion that proves the two implementations agree on real Excel,
including `AscW` sign normalisation and locale-invariant numeric formatting.

### 25.7 Mid-write rollback proof — matrix row 6

Inject a failure after one or more `_Calc` blocks have been mutated (§12.4), using
the Phase-4 `FailPointCheck` mechanism already proven on target. Assert:

- previous totals, driver rows and annual rows restored exactly;
- `calc_totals` `C23:C32` restored exactly, including blanks;
- **`C13:C16` restored exactly** — the last-successful snapshot;
- `C17:C20` correctly **published as new**: `FAILED`, the failure detail, the
  re-derived status and a fresh evaluation timestamp. **The comparison is
  `C13:C16`, never all of `C13:C20`** (§12.5) — asserting the latter would be
  asserting that the failure was never recorded;
- no mixed snapshot survives;
- `PCCM_CalculationAttemptResult() = FAILED` with the failure detail;
- `PCCM_CalculationStatus() = STALE`, derived independently — **not** forced from
  the attempt result;
- Excel application state restored.

Run this at **both** injection boundaries: after a table has been mutated, and at
the final `C13:C20` commit assignment (§12.3). Both must take the same path.

If the harness can observe it, also assert **moment 1** of §12.5 — immediately
after rollback and before failure metadata, all eight `C13:C20` cells equal their
previous values.

### 25.8 No change events

No `Worksheet_Change` or `Workbook_SheetChange` handler exists; status is computed
on demand (§11.12), and a sweep asserts neither handler was introduced. The
`_Calc` status cell is **last-evaluated**, and the harness reads it only after an
explicit `PCCM_Calculate` or `PCCM_CalculationStatus` call — never treating it as
live.

---

## 26. Performance

Design target: **200 Cost Lines, 100 Risks, 25 project years, 100,000
iterations** — targets, not caps (§16.2).

| Quantity | Complexity |
|---|---|
| resolution (worksheet reads) | `O(D × N)` = 7,500 cells, **once** |
| `Knom` / `Kpv` build | `O(D × N)` = 7,500 multiply-adds, **once** |
| fingerprint | `O(D × N)` encode + `O(D log D)` sort, **once per status query** |
| analytical totals A–E | `O(D)` = 300 |
| annual series | `O(D × N)` = 7,500 |
| later, per simulation iteration | `O(D)` = 300 — **no worksheet access** |

Memory: `DriverFactors` ≈ 26 KB; per-driver weight arrays `300 × 25 × 8 B = 60 KB`;
`YearFactors` negligible. **Under 100 KB resident**, which is why the simulation
phase can hold it all for an entire run.

The decisive property: at 100,000 iterations the kernel performs ~30 million
multiply-adds and **zero** Excel round trips. A worksheet-reading design would
perform ~750 million COM calls.

### Structures carried into Phase 6

```vb
Type DriverFactors            ' one per Cost Line and per Risk
    PermanentId   As String
    IsRisk        As Boolean
    Knom          As Double
    Kpv           As Double
    Quantity      As Double   ' 1 for risks
    Probability   As Double   ' 1 for cost lines
    DistKind      As Long     ' Triangular | BetaPert | Uniform
    CentralBasis  As String   ' "ML" | "Midpoint"
    MinValue      As Double
    MostLikely    As Double
    MaxValue      As Double
    Central       As Double
    MeanValue     As Double
End Type

Type YearFactors
    ProjectIndex  As Long
    CalendarYear  As Long
    DiscountF     As Double
End Type
```

plus one `Double` array per driver of `w_{i,y} × infl_{i,y}`, length `N`.
Simulation needs **only** these: no worksheet, no ListObject, no Range.

---

## 27. Implementation sequence

1. **`spec/calc_contract.yaml`** — the fifth authority, scoped per §20: the exact
   `_Calc` anchors of §16.3, schemas, tolerances and conditioning coefficients,
   labels, formats, state labels, `FP_VERSION`, reserved-cell declarations.
   Loader + validator, fail-loud, including the band non-overlap and Phase-4
   reservation assertions. **No hash constants.**
2. **Fingerprint reference implementation and vectors** — the canonical encoder,
   the double-modulus digest, and the locked literals: the ten numeric encodings,
   the eight collision probes, the 366-unit stream and `50B6EB0E26857EA7`. Written
   **before** any consumer, because everything downstream is compared against it.
3. **`builder/pccm_builder/calc_oracle.py`** — pure-Python implementation of
   §5–§15. **Golden-case tests written first**, from the hand derivations in §23,
   plus the oracle-independence test (§21).
4. **Safe arithmetic primitives** — `SafeAdd` … `SafeAccumulate`, the stable
   formulas of §19.2 and the iterative factor builders of §19.3, with their own
   Gate-A tests including the naive-overflow and underflow cases.
5. **Stage-A emission** — `_Calc` blocks per the contract; `modCalcContract.bas`
   generated; `build/phase5_cases.json` emitted. Post-build verification extended.
6. **`modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint`** — the pure
   numerical kernel, with the no-worksheet sweep, the `On Error Resume Next`
   whitelist and the §19 guards active from the first commit, and the fingerprint
   vectors passing against the Python side.
7. **`modCalcResolve`** — reference-set construction, then worksheet → numeric
   structures, honouring the referenced-only rule of §8.
8. **`modCalcCheck`** — Phase-5 numerical prerequisites (§18); reports, never
   repairs.
9. **`modCalcReport` + the transactional orchestration of §12** —
   snapshot, write, verify, commit-last; `calc_state` maintenance; refusal through
   `modAppState`. `PCCM_Calculate` **and the five accessors** of §11.13 — six
   public `PCCM_` automation/API entry points in total (erratum E4: a bound on the
   `PCCM_` endpoint surface, not on the `Public` keyword). **No button.**
10. **Gate-A source review.**
11. **Gate-B harness extension** — additive module/button assertions, the new
    functional coverage, the fingerprint-parity assertion, the stale/revert
    sequence (§25) and the mid-write rollback scenario. On approval, the Windows
    run.

Steps 1–10 are Linux-only. No Windows execution before Gate A is approved.

The ordering is deliberate: **the fingerprint vectors and the safe primitives come
before the code that uses them**, because both are things a later implementation
would otherwise be tempted to define by whatever it happens to produce.

---

## 28. Acceptance criteria

1. every Phase 1–4 Linux/static test still passes, none weakened;
2. post-build verification passes, extended to the `_Calc` blocks;
3. every golden case in §23 passes against the oracle **and** on Windows, with
   refusal cases producing no partial result and leaving the previous successful
   snapshot intact;
4. every emitted expected value equals its hand-derived literal (§21);
5. the canonical stream is exactly 366 code units for the reference vector and its
   digest is **`50B6EB0E26857EA7`** — proven at **Gate A** by the Python reference
   implementation and at **Gate B** by real VBA on real Excel. *(Erratum E3: there
   is no VBA execution on Linux, so no "VBA on Linux-side tests" proof exists or
   may be claimed — §21.0.)* The eight collision probes are distinct; the ten
   numeric encodings match, and the canonical encoder is separator-invariant —
   proven for the Python reference at Gate A and for real VBA at Gate B (§24.1);
6. **the modular reduction uses `Double` arithmetic only** — no native `Mod` and
   no `\` on any pre-reduction intermediate, proven at Gate A by a **static**
   source rule over the executable recurrence code; the four §11.5 reduction
   vectors match exact integer arithmetic for both moduli — in the Python mirror at
   Gate A, and in the **real VBA reducer** at Gate B (§24.1) — and the same reducer
   reproduces the end-to-end digest;
7. identities I1, I2, I3a–c, I4a–c, I5 hold within the §15 tolerances **using the
   per-identity conditioning scales**, with `B` and `E` independently accumulated.
   `A_pv ≤ A_nom` is a conditional diagnostic, **not** a gate;
8. calculation is refused, with a specific message, for every §18 numerical
   prerequisite and for `STRUCTURE CHANGE PENDING`; **no uncontrolled VBA overflow
   escapes, no overflow becomes a fabricated zero, and no discount factor
   underflows silently to zero**;
9. the stable formulas of §19.2 return the correct mean where the naive form
   overflows;
10. `modCalcFactors`, `modCalcAnalytical` and `modCalcFingerprint` contain **none**
   of `Application.`, `ThisWorkbook`, `ActiveWorkbook`, `Worksheets`, `Worksheet`,
   `Range`, `Cells`, `ListObject`, `ListObjects` or `modWorkbook.`, and no
   `On Error Resume Next` exists outside the documented safe primitives — both
   proven by sweep. The canonical numeric encoder is a **pure function of its
   arguments**, including the decimal separator, proven by the separator-injection
   case (§23 case 35);
11. **the two status axes are orthogonal.** `PCCM_CalculationStatus()` returns only
    `NOT CALCULATED` / `CURRENT` / `STALE` / `INVALID` — never `REFUSED` — and
    `PCCM_CalculationAttemptResult()` returns `NONE` / `SUCCESS` / `REFUSED` /
    `FAILED`. All six rows of the §25.1 matrix hold on real Excel, including an
    invalid input with no attempt (`INVALID` + `SUCCESS`) and a rolled-back write
    failure (`STALE` + `FAILED`);
12. the fingerprint detects staleness for every covered input and **not** for
    Description, row order, Selected Confidence Level or unreferenced Config;
    status returns to `CURRENT` when inputs are restored exactly, with no
    recalculation and with the attempt axis unchanged; there is no change-event
    handler anywhere;
13. **write-back is transactional**: the success commit is **one contiguous
    `C13:C20` `Range.Value2` assignment** and the last write of the operation, but
    it is a commit **attempt** — a failure of that assignment, or of its
    verification, takes the same rollback path as a mid-table failure. Both
    injection boundaries restore every table **and both scalar value ranges**
    completely — blanks as blanks — with no mixed state surviving, and the attempt
    is recorded as `FAILED` only after rollback has completed. The post-failure
    comparison of the successful snapshot is **`C13:C16` plus the analytical
    blocks**, never all of `C13:C20` (§12.5);
14. **a pre-write refusal leaves the last-successful analytical snapshot and
    `C13:C16` unchanged, while `C17:C20` changes as it must.** The two groups are
    compared separately, and the phrase "workbook byte-identical" is not used of a
    refused attempt;
15. the `tblCalcDrivers` column sums reconstruct A, B, C and D exactly, and no
    column carries two meanings by Driver Kind;
16. the full **35/35** Phase-4 functional matrix still passes; all 8 Phase-4
    modules and all 5 Phase-4 buttons persist; Phase-5 modules persist; **no
    Calculate button exists**;
17. the harness asserts every calculated value with no manual inspection;
18. Excel shuts down naturally with clean transient COM release;
19. **Phase 5 introduces no RNG implementation, no sampling implementation and no
    simulation output, and makes no use of Iterations, Random Seed or Selected
    Confidence Level in any analytical calculation or in the fingerprint;
    `_SimData` remains unchanged and unused** (§22);
20. **real Windows / VBA exercises the canonical encoder and the reducer directly
    against the complete locked vector set** of §24.1 — the ten numeric encodings,
    both decimal separators, all four reduction vectors, the `> U+7FFF` and
    non-BMP UTF-16 vectors, and the 366-code-unit reference stream — not merely the
    golden-case digest. Where a transient test-only diagnostic module provides that
    access, it does not enter the Stage-B production manifest, does not persist in
    the accepted workbook, creates no button, and adds no `PCCM_` endpoint.

---

## 29. Decisions

**D1–D6 are locked** (§4), and Revision E leaves them untouched — as it leaves
the deterministic / mean / EMV mathematics, the inflation and discount
conventions, the referenced-only FX and inflation scope, `Quantity > 0`, the
profile blank semantics, the length-prefixed canonical encoding, the UTF-16 rules,
the ten numeric vectors, the eight collision probes, **every hash constant, the
mathematical recurrence and `50B6EB0E26857EA7`**, the safe-arithmetic design, the
cancellation-aware tolerances, the `_Calc` anchors and schemas, the two orthogonal
status axes, the snapshot/rollback philosophy, the no-button and no-event rules,
the Phase-4 35/35 requirement and golden cases 1–35.

Revision E changed one implementation blocker and four consistency points:

| Was | Now |
|---|---|
| the recurrence intermediate declared "exact in VBA arithmetic" — true of the arithmetic, **false of the operator**, since VBA `Mod` and `\` are `Long`-typed and `x` reaches `2.8 × 10¹¹` | a locked `Double`-only reduction — `Fix`, subtract, two one-step corrections — verified against exact integer arithmetic over 300,000 random cases with zero mismatches, reproducing the same digest |
| the final `C13:C20` assignment treated as infallible | it is a commit **attempt**; its failure takes the same rollback path, and steps 4–7 are all inside the envelope |
| acceptance implying `C13:C20` equals its old values after a failure | the comparison is **`C13:C16` plus the analytical blocks**; `C17:C20` is correctly published as new |
| "`PCCM_Calculate` and the four accessors" | **six public `PCCM_` automation/API entry points**: `PCCM_Calculate` plus **five** accessors (erratum E4 — a bound on the `PCCM_` endpoint surface, not on the `Public` keyword) |
| initial `calc_state` values unstated | locked — blanks, `NONE`, `NOT CALCULATED`, and `Fingerprint Version` **blank** until a successful commit |
| accessor empty cases unstated | locked — `""` for no digest, `NONE` before the first attempt, no sentinel hash strings |
| "the entire canonical stream, byte for byte" | "UTF-16 code-unit for UTF-16 code-unit" — the hash is defined over code units, never over a byte encoding |

**No open decisions remain.**

### Revision E errata — applied in place, no design change

Four errata (E1–E4) were raised at acceptance and applied editorially; §0 records
them in full, and each amended passage cites its erratum number. They correct the
post-failure `calc_state` comparison (E1), the `131 × Long.MaxValue` arithmetic
(E2), acceptance wording that implied VBA executes on Linux (E3), and the
"callable" terminology (E4). **No constant, vector, anchor, schema, expected value
or locked decision changed.** In addition, §24.1 records a locked **Gate-B**
requirement — direct Windows/VBA vector coverage — which is deliberately **not
implemented** at Gate-A Step 1.

Model source, workbook artifacts, contracts, bootstrap, harness and Phase-4 tests
are unchanged by this document. No code has been written **by this document**;
implementation begins at Gate-A Step 1, recorded separately in
`docs/phase5_gate_a_step1.md`.

**PHASE 5 PLAN REVISION E READY FOR REVIEW**
