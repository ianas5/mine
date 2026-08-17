# PCCM — Phase 5 plan: deterministic and analytical calculation engine

**Revision B.** Revision A's mathematical core was accepted in principle. This
revision locks D1–D6, replaces the insufficient fingerprint design, corrects the
FX/inflation blocking scope, strengthens the reconciliation identities, removes an
identity that was not universally true, locks the physical `_Calc` layout, adds
numerical finiteness guarantees, scopes the new contract's authority, and removes
the user-facing Calculate button.

**DESIGN GATE. No code, no VBA, no workbook change, no build artifacts.** Phase 4
is accepted and closed; nothing in `src/`, `spec/`, `builder/`, `bootstrap/` or
`tests/` is touched by this document.

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
simulation reconciliation · **any user-facing Calculate button** (§8 / §16).

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
within tolerance (§14), **and** no cell in that row may be blank (D4).

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

**Header scalars**

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

### 11.3 Canonical mechanism

One reusable mechanism, designed for extension:

1. **Canonical field encoding.** Text fields verbatim. Numeric fields as a
   round-trip-exact 17-significant-digit form, format string
   `"0.0000000000000000E+00"`, so VBA and the Python oracle produce byte-identical
   input. A shared test vector pins this on both sides (§20).
2. **Per-driver digest.** Fields joined with a reserved separator (`U+001F`, which
   cannot appear in a workbook string), then hashed.
3. **Ordering.** Per-driver digests are **sorted by Permanent ID**, not by digest
   and not by row — so a reorder cannot change the result.
4. **Fold.** Header scalars, then the sorted per-driver digests, folded into one
   global digest, rendered as fixed-width hex.
5. **Extension point.** The fold takes a *named section list*. Phase 6 appends
   simulation-only sections (Iterations, Seed, RNG stream identity) **without
   altering the analytical sections**, so the analytical subset stays comparable
   across phases. There is exactly one fingerprint format, versioned by a
   `fingerprint_version` field stored alongside it.

**Hash choice.** A double-modulus polynomial hash (base 131, two distinct 31-bit
primes) computed entirely in `Double`. Maximum intermediate `2³¹ × 131 ≈ 2.8×10¹¹`,
well inside the `2⁵³` exact-integer range of a `Double`, so it is exact in VBA
without unsigned 64-bit arithmetic and reproduces identically in Python. Two
independent moduli give ~62 bits of separation — ample for change detection, and
not a security primitive.

### 11.4 `calc_state`

Stores at least:

| Field | Purpose |
|---|---|
| `fingerprint` | the input fingerprint of the last **successful** calculation |
| `fingerprint_version` | so a format change is detectable, never silently mis-compared |
| `stamp` | calculation timestamp |
| `applied_timeline` | the applied triple used |
| `status` | `NOT CALCULATED` / `CURRENT` / `STALE` / `REFUSED` |
| `refusal_reason` | populated only when `status = REFUSED` |

### 11.5 Status is computed on demand — no events

**No `Worksheet_Change` and no `Workbook_SheetChange` handler is permitted**, in
keeping with the Phase-4 rule that structural state is never maintained by hidden
automation.

Status is derived when asked, by recomputing the current input fingerprint and
comparing it with the stored one:

```
no stored fingerprint          → NOT CALCULATED
stored, and recomputed matches → CURRENT
stored, and recomputed differs → STALE
last attempt refused           → REFUSED  (with refusal_reason)
```

**Matching applied timelines alone never yields CURRENT.** The whole fingerprint
must match.

If the current inputs are themselves invalid, recomputation cannot produce a
fingerprint; status reports `STALE` with the resolution failure as the reason,
never `CURRENT`.

### 11.6 Callable surface for Gate B

Public, invoked by `Application.Run` — **no button** (§16):

```
PCCM_Calculate                  orchestration; refuses cleanly
PCCM_CalculationStatus()        NOT CALCULATED | CURRENT | STALE | REFUSED
PCCM_CalculationFingerprint()   the STORED fingerprint
PCCM_CurrentInputFingerprint()  the fingerprint of the inputs as they are NOW
PCCM_CalculationRefusal()       the refusal reason, or empty
```

The two fingerprint accessors are separate deliberately: Gate B must be able to
show the stored one **unchanged** while the current one has moved (§24 step 7).

---

## 12. The three headline measures

Let `Q_i` = Quantity, `c_i` = deterministic central, `m_i` = distribution mean,
`p_j` = probability, `s_j` = expected severity.

```
A_nom = Σ_i c_i·Q_i·Knom_i     A_pv = Σ_i c_i·Q_i·Kpv_i     Escalated Deterministic Base
C_nom = Σ_i m_i·Q_i·Knom_i     C_pv = Σ_i m_i·Q_i·Kpv_i     Mean-Basis Base Cost
D_nom = Σ_j p_j·s_j·Knom_j     D_pv = Σ_j p_j·s_j·Kpv_j     Expected Risk / EMV
```

---

## 13. Annual analytical cash flow

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

## 14. Reconciliation identities and tolerances

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

### Tolerances

| Purpose | Tolerance | Rationale |
|---|---|---|
| profiling sum = 100% | `\|Σw − 1\| ≤ 1e-9` absolute | percentages are entered to 2 dp and stored as binary doubles; across the 200-column structural maximum the accumulated representation error is bounded by roughly `200 × 2⁻⁵² ≈ 4.4e-14`. `1e-9` gives ~4 orders of headroom and is still 3 orders tighter than a 1-in-a-million entry slip. Exact binary equality is never used. |
| identities I1–I4 | `\|Δ\| ≤ max(1e-6 SAR, 1e-12 × scale)`, `scale = max(\|A\|,\|C\|,\|E\|)` | absolute floor for small models, relative term for large ones; a pure relative test degenerates at zero, a pure absolute test fails at billions |
| FX positivity | `rate > 0`, no epsilon | present and positive, or refused |
| `1 + rate > 0`, `1 + r > 0` | strict | D2, D3 |

No Monte Carlo statistical tolerance is invented here.

---

## 15. Physical `_Calc` layout — locked

`_Calc` is `hidden` (not veryHidden) so an auditor can inspect it. `_SimData`
remains `veryHidden` and **untouched and unused** by Phase 5. No user-facing input
sheet is written by Phase 5.

### 15.1 Phase-4 reservation

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

### 15.2 Growth strategy — column bands, not vertical stacking

Every Phase-5 table has a **fixed column schema** and an **unbounded row count**.
Stacking them vertically would make a growing table collide with the block below,
so each dynamic ListObject is given its **own column band**, all anchored at the
same header row, growing downward with nothing beneath it.

**No block is capped at 200 Cost Lines or 100 Risks.** Those are design targets,
not business maxima — consistent with the Phase-4 refusal to encode 25 years.

### 15.3 Locked layout

All blocks anchor at **header row 15**; scalar blocks occupy fixed rows in the
B:C band below the Phase-4 area.

| Block | Kind | Anchor | Columns | Rows | Growth | Visible when unhidden | Owner |
|---|---|---|---|---|---|---|---|
| *(Phase-4 counters)* | scalars | `B8:E11` | B, C, E | fixed | none | yes | **Phase 4 — reserved** |
| `calc_state` | scalars | `B13` label / `C13` value, 6 rows | B, C, E | 6, fixed | none | yes | Phase 5 |
| `calc_totals` | scalars | `B21` label / `C21` value, 10 rows | B, C, E | 10, fixed | none | yes | Phase 5 |
| `tblCalcYears` | ListObject | header `H15` | 3 | applied year count | vertical | yes | Phase 5 |
| `tblCalcInflationFactors` | ListObject | header `M15` | 4 | referenced profiles × required years | vertical | yes | Phase 5 |
| `tblCalcFX` | ListObject | header `S15` | 3 | referenced currencies | vertical | yes | Phase 5 |
| `tblCalcDrivers` | ListObject | header `X15` | 16 | identified cost lines + risks | vertical | yes | Phase 5 |
| `tblCalcAnnual` | ListObject | header `AQ15` | 7 | applied year count | vertical | yes | Phase 5 |

Column letters are illustrative of the **band pattern**; the contract owns the
exact anchors, and a build-time assertion proves no two bands overlap given their
schemas. What is locked is the **shape**: fixed-width, vertically growing, one
band each, none above another.

### 15.4 Schemas

**`calc_state`** — `B13:C18`, labels in B, values in C, notes in E

| Row | Label | Value | Format |
|---|---|---|---|
| 13 | Calculation Status | `NOT CALCULATED` / `CURRENT` / `STALE` / `REFUSED` | `@` |
| 14 | Calculation Stamp | timestamp of last successful calculation | `yyyy-mm-dd hh:mm:ss` |
| 15 | Input Fingerprint | stored fingerprint, hex | `@` |
| 16 | Fingerprint Version | integer | `0` |
| 17 | Applied Timeline Used | `base/start/duration` | `@` |
| 18 | Refusal Reason | empty unless `REFUSED` | `@` |

**`calc_totals`** — `B21:C30`, labels in B, values in C, all `#,##0.00` SAR

```
A_nom  Escalated Deterministic Base — Nominal      A_pv  … PV
B_nom  Uncertainty Mean Shift — Nominal            B_pv  … PV
C_nom  Mean-Basis Base Cost — Nominal              C_pv  … PV
D_nom  Expected Risk / EMV — Nominal               D_pv  … PV
E_nom  Analytical Mean Total — Nominal             E_pv  … PV
```

**`tblCalcYears`**

| Column | Type | Format | Units |
|---|---|---|---|
| Project Index | integer | `0` | index, from 1 |
| Calendar Year | integer | `0` | year |
| Discount Factor | double | `0.000000` | dimensionless |

**`tblCalcInflationFactors`** — long form, as required

| Column | Type | Format | Units |
|---|---|---|---|
| Inflation Profile | text | `@` | key |
| Calendar Year | integer | `0` | year |
| Annual Rate | double | `0.00%` | rate |
| Cumulative Inflation Factor | double | `0.000000` | dimensionless |

Long form is what makes this scale with a dynamic profile count without a
variable-width profile-by-year block colliding with its neighbours.

**`tblCalcFX`**

| Column | Type | Format | Units |
|---|---|---|---|
| Currency | text | `@` | key |
| FX to SAR | double | `0.000000` | SAR per unit |
| Referenced By | integer | `0` | driver count |

**`tblCalcDrivers`** — one row per identified Cost Line and Risk

| Column | Type | Format | Units |
|---|---|---|---|
| Permanent ID | text | `@` | key |
| Driver Kind | text | `@` | `Cost Line` / `Risk` |
| Distribution | text | `@` | — |
| **Central Basis** | text | `@` | `ML` / `Midpoint` (§10) |
| Currency | text | `@` | — |
| FX to SAR | double | `0.000000` | SAR per unit |
| Inflation Profile | text | `@` | — |
| Quantity | double | `#,##0.00` | units (`1` for risks) |
| Probability | double | `0.0%` | fraction (`1` for cost lines) |
| Central Value | double | `#,##0.00` | source currency |
| Mean Value | double | `#,##0.00` | source currency |
| Knom | double | `0.000000` | SAR per source unit |
| Kpv | double | `0.000000` | SAR per source unit |
| Deterministic Nominal | double | `#,##0.00` | SAR |
| Mean-Basis Nominal | double | `#,##0.00` | SAR |
| Mean-Basis PV | double | `#,##0.00` | SAR |

**`tblCalcAnnual`**

| Column | Type | Format | Units |
|---|---|---|---|
| Project Index | integer | `0` | index |
| Base Cost Nominal · Expected Risk Nominal · Total Nominal | double | `#,##0.00` | SAR |
| Base Cost PV · Expected Risk PV · Total PV | double | `#,##0.00` | SAR |

### 15.5 Update trigger and validation

Every block: **trigger** = `PCCM_Calculate`; **validation** = the numerical
prerequisites of §17 plus the identities of §14; **ownership** = Phase 5;
**units** as tabulated.

`_Calc` is a **written record of what the in-memory kernel computed**. Nothing
reads it back to compute anything else — that would recreate the worksheet
dependency the whole design exists to avoid. **No per-iteration data is ever
written here.**

---

## 16. VBA / numerical module boundaries

The hard rule: **mathematical functions must not read worksheet cells.**

| Module | Layer | Responsibility | Worksheets |
|---|---|---|---|
| `modCalcContract` | generated | Phase-5 constants projected from `calc_contract.yaml` | n/a |
| `modCalcResolve` | resolution | builds the referenced-currency and referenced-profile sets from the registers, then reads Setup, profiling grids, `tblFXRates`, `tblInflation` into plain numeric structures | **yes** |
| `modCalcFactors` | numerical | `InflationFactors`, `DiscountFactors`, `BuildKnom`, `BuildKpv` | **no** |
| `modCalcAnalytical` | numerical | `TriangularMean`, `PertMean`, `UniformMean`, `DeterministicCentral`, `ExpectedRisk`, A–E accumulations, annual series | **no** |
| `modCalcFingerprint` | numerical | canonical encoding, per-driver digest, sort by ID, fold (§11.3) | **no** |
| `modCalcCheck` | validation | Phase-5 numerical prerequisites; reports, never repairs | **yes**, read only |
| `modCalcReport` | presentation | writes the `_Calc` blocks; reports refusal through the Phase-4 `modAppState` surface | **yes** |

`modCalcFingerprint` is numerical deliberately: the fingerprint must be computable
from resolved data alone, so Phase 6 can extend it without a worksheet.

**A static sweep enforces the boundary:** no worksheet identifier (`ThisWorkbook`,
`Worksheets`, `Range`, `ListObjects`, `Cells`, `modWorkbook.*`) may appear in
`modCalcFactors`, `modCalcAnalytical` or `modCalcFingerprint`. Mechanically
checkable on Linux, permanent, in the established style.

### No user-facing Calculate button

**LOCKED for Phase 5: `PCCM_Calculate` — yes. Calculate button — no.**

A standalone Calculate button was not part of the locked Dashboard command set,
and adding user-facing workflow before Results, Model Check and Run Simulation
exist would clutter the UI. The Windows harness invokes `PCCM_Calculate` directly
through `Application.Run`. Later phases call the same orchestration from Run
Check / Run Simulation / output-refresh pathways.

**The workbook keeps exactly the five Phase-4 buttons.** Gate B proves it (§23).

---

## 17. Validity and failure behaviour

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
| any numerical result non-finite (§18) | yes |

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

## 18. Numerical finiteness and overflow protection

`rate > −1` alone does not guarantee representable arithmetic, and there is
deliberately no business upper bound on rates or costs. **No VBA runtime overflow
may escape as an uncontrolled error, and no overflow may silently become a
fabricated zero.**

A single predicate, `IsUsableDouble(v)` — not NaN, not ±∞, `|v| ≤ 1.7e308` — is
applied after **every** stage:

| Stage | Guard |
|---|---|
| inflation compounding | each cumulative factor checked as it is built; a product that overflows refuses at the profile and calendar year that caused it |
| discount factor | checked for overflow **and for underflow to exactly 0** — underflow is silent, and a zero discount factor would quietly delete a year's PV. Refused when `r > −1` yet the factor collapses |
| `Knom` / `Kpv` | checked after each accumulation step, not only at the end |
| central / expected value | checked after evaluation |
| driver contribution | checked per driver |
| annual accumulators | checked per year, per series |
| A / B / C / D / E accumulators | checked after each driver is added |

Checking **during** accumulation rather than only at the end is deliberate: it
names the driver, profile or year responsible, instead of reporting that a total
is infinite.

Refusal carries a specific **numerical-range message** naming the stage and the
input. **No arbitrary business cap is invented to avoid implementing safe
arithmetic.**

---

## 19. `calc_contract.yaml` — scoped authority

Accepted, with its authority narrowly bounded so no duplicate source of truth is
created.

**It owns:** `_Calc` physical layout · block and table names · column schemas ·
labels · display formats · units · numerical tolerances · calculation-state
labels · the fingerprint version and separator · reserved-cell declarations.

**It must NOT restate:** driver schemas (`driver_contract.yaml`) · the
distribution list (`input_contract.yaml`, `tblDistributions`) · timeline
structural limits (`structure_contract.yaml`) · the FX convention
(`input_contract.yaml`) · permanent-ID rules (`structure_contract.yaml`) ·
Phase-4 structure rules. These are **referenced or projected** from the existing
authorities, and the loader asserts the projection matches.

**Mathematical semantics are not defined in YAML.** A formula written once in YAML
and again in VBA/Python is two sources of truth that will diverge. The division is:

```
this document + the tested numerical oracle   define the numerical semantics
calc_contract.yaml                            defines their workbook representation
```

---

## 20. Gate A — Linux / static

1. **Pure-Python numerical oracle** — `builder/pccm_builder/calc_oracle.py`,
   implementing §5–§14 independently, in the `structure_oracle.py` pattern: it
   defines the semantics the VBA must match **and** emits the expected values the
   Windows harness asserts, so the two cannot drift.
2. **Golden-case tests** — §22, every value hand-derived and written as a literal
   in the test.
3. **Oracle-independence test** — see below.
4. **Fingerprint test vectors** — a fixed input set with a hand-checked canonical
   encoding and expected digest, so the VBA and Python implementations are pinned
   to the same bytes, including the numeric format string and the separator.
5. **Contract validation** for `calc_contract.yaml`, including the authority
   boundary (§19) and non-overlap of `_Calc` bands and the Phase-4 reservation.
6. **Source sweeps**, extending the existing mechanical ones:
   - no worksheet identifier in `modCalcFactors` / `modCalcAnalytical` /
     `modCalcFingerprint`;
   - **no simulation dependency** — see §21 for how this is tested;
   - VBA block balance, line length, declaration-section placement, `Optional`
     defaults, applied to the new modules;
   - every emitted constant referenced; no structural literal restated.
7. **Post-build verification** extended with the `_Calc` block layout.

### Golden oracle independence

The architecture is deliberately:

```
hand-derived literals  →(verify)→  Python oracle  →(emit)→  phase5_cases.json  →(assert)→  Windows/VBA
```

A static test asserts that **every expected value emitted into
`phase5_cases.json` equals its separately hard-coded hand-derived literal**. The
JSON must not become self-validating merely because the oracle produced it.

For refusal cases the harness verifies:

- the specific refusal **class/message**, not merely that something failed;
- **no partial analytical totals** were written;
- the previous successful calculation snapshot was **not overwritten** as though
  the failed calculation had succeeded;
- `calc_state.status` reflects the invalid/refused state.

Gate A ends with a source review, exactly as Phase 4 did.

---

## 21. Simulation-artefact acceptance wording — corrected

Revision A's *"no Monte Carlo artefact exists anywhere in the phase"* was too
broad: the frozen workbook already legitimately contains `inpMonteCarloIterations`,
`inpRandomSeed`, `inpSelectedConfidenceLevel` and the `_SimData` sheet, all from
the locked architecture.

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
- the fingerprint field list (§11.1) contains none of the three excluded inputs —
  asserted against the oracle's field list, not against prose.

---

## 22. Golden and refusal matrix

Cases 1–15 are unchanged from Revision A and remain hand-derived. Cases 16–25 lock
the behaviour required by D1–D4 and §17–§18.

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
| **16** | **Quantity = 0** | case 1 with `Qty 0` | **refusal** (§17) |
| **17** | **Quantity < 0** | case 1 with `Qty −5` | **refusal** |
| **18** | **Discount Rate = −100%** | `r = −1` | **refusal** — `1 + r = 0` (D3) |
| **19** | **Discount Rate negative but > −100%** | case 3 with `r = −5%` | **accepted**; `disc = 1, 1/0.95, 1/0.9025`, so `A_pv > A_nom` — correct, and why I5 is not a gate (§14) |
| **20** | **Inflation Rate = −100%** | one required year at `−100%` | **refusal** — `1 + rate = 0` (D2) |
| **21** | **Inflation Rate negative but > −100%** | case 13 | **accepted** |
| **22** | **Uniform with populated ML** | `Min 80 / ML 999 / Max 150` | **accepted**; central = mean = `115`; ML ignored **and excluded from the fingerprint** (D1) |
| **23** | **100%-summing profile containing a blank** | Dur 3, weights `50% / blank / 50%` | **refusal** (D4) — sums to 100% and is still refused |
| **24** | **Double overflow** | inflation `1e300%` compounded over several years, or an extreme unit cost | **controlled refusal** with a numerical-range message naming the stage; no uncontrolled VBA error, no fabricated zero (§18) |
| **25** | **unreferenced incomplete FX / Config row** | valid SAR-only model + a duplicate, blank-rate `EUR` row referenced by nothing | **does NOT block**; calculation succeeds and the fingerprint is unaffected (§8) |

Cases 14–18, 20, 23 and 24 assert a **refusal**, not a number, and must produce no
partial result. Cases 19, 21, 22 and 25 assert **acceptance** where a naive
implementation would over-block.

Cases 16–25 need not each be a separate full fixture; the Gate-B matrix may
exercise them compactly. Their **behaviour is locked** regardless.

---

## 23. Gate B — real Windows / Excel

Extends the accepted `phase4_functional_test.ps1` matrix; it does not replace it.

### Additive expectations

Phase 5 adds VBA modules, so the bootstrap and harness expectations become
**additive without weakening Phase 4**. Gate B proves:

| Claim | How |
|---|---|
| all 8 original Phase-4 modules still persist | the reopen-and-verify step reads the module list from the manifest, which now declares Phase-4 **and** Phase-5 modules; the Phase-4 eight are asserted present **by name** |
| all Phase-5 modules persist | same list, new names asserted present by name |
| all 5 Phase-4 command buttons still persist | asserted by shape name from the manifest, exactly as today |
| **no Phase-5 button was added** | the manifest's button count is asserted `= 5`, and the sheet's shape inventory is asserted to contain **no** shape whose `OnAction` is `PCCM_Calculate` (§16) |
| the VBA project still compiles | `A1` unchanged — still the first `Application.Run` of the run |

**The full 35/35 Phase-4 functional matrix remains mandatory** and must pass before
any Phase-5 scenario is considered accepted.

### New Phase-5 functional coverage

FX resolution (foreign currency and the SAR identity) · inflation compounding
(`Base = Start` and `Base < Start`) · discount factors at indices 1, 2, 3 ·
profiling factor application by permanent ID · deterministic base (Nominal, PV) ·
mean-basis base (Nominal, PV) · expected risk (Nominal, PV) · all six annual
series · identities I1–I5 asserted in the workbook · refusal on every §17
numerical prerequisite · refusal on `STRUCTURE CHANGE PENDING` · the §24 stale
fingerprint sequence · clean shutdown and clean transient COM release.

Every expected value comes from `build/phase5_cases.json`. **The harness asserts
every calculated value; the user inspects no cells manually.**

Phase-4 harness disciplines carry over unchanged and are non-negotiable:
caller-side `@(...)`, one pipeline object per row, container factories emitted
non-enumerated, `catch` attached to its `try`, keyed-only fixtures, failure-safe
cleanup, per-scenario clean-structure prerequisites, `$excelIdentity`.

---

## 24. Gate-B stale-fingerprint scenario

An explicit Windows functional oracle for the fingerprint, using the **same
canonical semantics intended for later simulation reuse**.

**Primary sequence**

1. establish a valid analytical fixture;
2. `PCCM_Calculate`;
3. assert `PCCM_CalculationStatus() = CURRENT`;
4. capture headline values **and** `PCCM_CalculationFingerprint()`;
5. change Quantity **or** one profiling weight, **without** touching the timeline
   and **without** recalculating;
6. assert `PCCM_CalculationStatus() = STALE`;
7. assert `PCCM_CalculationFingerprint()` is **still the old value** — the stored
   snapshot is identifiable as the old one — while
   `PCCM_CurrentInputFingerprint()` has changed;
8. `PCCM_Calculate`;
9. assert status `CURRENT`;
10. assert the stored fingerprint **changed**;
11. assert the affected analytical value changed **to the oracle value**.

**Non-staleness proofs** — each must leave status `CURRENT` with the stored
fingerprint unchanged:

- changing a **Description**;
- changing **row order** (a real `ListObject.Sort`, as `B2`/`K2` already do);
- changing **Selected Confidence Level**;
- changing an **unreferenced** FX row or Config value.

**Refusal-state proof** — make an input invalid, run `PCCM_Calculate`, assert
status `REFUSED` with a specific reason, assert no partial totals were written and
that the previous successful snapshot was not overwritten (§20).

No `Worksheet_Change` or `Workbook_SheetChange` handler exists; status is computed
on demand (§11.5), and a sweep asserts neither handler was introduced.

---

## 25. Performance

Design target: **200 Cost Lines, 100 Risks, 25 project years, 100,000
iterations** — targets, not caps (§15.2).

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

## 26. Implementation sequence

1. **`spec/calc_contract.yaml`** — the fifth authority, scoped per §19: `_Calc`
   physical layout (§15), tolerances, labels, formats, fingerprint version and
   separator, reserved-cell declarations. Loader + validator, fail-loud, including
   the band non-overlap and Phase-4 reservation assertions.
2. **`builder/pccm_builder/calc_oracle.py`** — pure-Python implementation of
   §5–§14 and the §11.3 fingerprint. **Golden-case tests written first**, from the
   hand derivations in §22, plus the oracle-independence test (§20).
3. **Stage-A emission** — `_Calc` blocks per the contract; `modCalcContract.bas`
   generated; `build/phase5_cases.json` emitted. Post-build verification extended
   to the new blocks.
4. **`modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint`** — the pure
   numerical kernel, with the no-worksheet sweep and the §18 finiteness guards
   active from the first commit, and the fingerprint test vectors passing against
   the Python side.
5. **`modCalcResolve`** — reference-set construction, then worksheet → numeric
   structures, honouring the referenced-only rule of §8.
6. **`modCalcCheck`** — Phase-5 numerical prerequisites (§17); reports, never
   repairs.
7. **`modCalcReport` + `PCCM_Calculate` and the four status accessors (§11.6)** —
   orchestration, `_Calc` write-back, `calc_state` maintenance, refusal through
   `modAppState`. **No button.**
8. **Gate-A source review.**
9. **Gate-B harness extension** — additive module/button assertions (§23), the new
   functional coverage, and the stale-fingerprint scenario (§24). On approval, the
   Windows run.

Steps 1–8 are Linux-only. No Windows execution before Gate A is approved.

---

## 27. Acceptance criteria

1. every Phase 1–4 Linux/static test still passes, none weakened;
2. post-build verification passes, extended to the `_Calc` blocks;
3. every golden case in §22 passes against the oracle **and** on Windows, with
   refusal cases producing no partial result and leaving the previous successful
   snapshot intact;
4. every emitted expected value equals its hand-derived literal (§20);
5. identities I1, I2, I3a–c, I4a–c, I5 hold within §14 tolerances, with `B` and
   `E` independently accumulated. **`A_pv ≤ A_nom` is a conditional diagnostic,
   not a gate**;
6. calculation is refused, with a specific message, for every §17 numerical
   prerequisite and for `STRUCTURE CHANGE PENDING`; no uncontrolled VBA overflow
   escapes and no overflow becomes a fabricated zero (§18);
7. `modCalcFactors`, `modCalcAnalytical` and `modCalcFingerprint` contain no
   worksheet access, proven by sweep;
8. the fingerprint detects staleness for every covered input and **not** for
   Description, row order, Selected Confidence Level or unreferenced Config
   (§24), with no change-event handler anywhere;
9. the full **35/35** Phase-4 functional matrix still passes; all 8 Phase-4
   modules and all 5 Phase-4 buttons persist; Phase-5 modules persist; **no
   Calculate button exists**;
10. the harness asserts every calculated value with no manual inspection;
11. Excel shuts down naturally with clean transient COM release;
12. **Phase 5 introduces no RNG implementation, no sampling implementation and no
    simulation output, and makes no use of Iterations, Random Seed or Selected
    Confidence Level in any analytical calculation or in the fingerprint;
    `_SimData` remains unchanged and unused** (§21).

---

## 28. Decisions

**D1–D6 are locked** (§4). No open decisions remain.

Model source, workbook artifacts, contracts, bootstrap, harness and Phase-4 tests
are unchanged by this document. No code has been written.

**PHASE 5 PLAN REVISION B READY FOR REVIEW**
