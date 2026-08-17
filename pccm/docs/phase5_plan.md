# PCCM — Phase 5 plan: deterministic and analytical calculation engine

**DESIGN GATE. No code, no VBA, no workbook change.** Phase 4 is accepted and
closed; nothing in `src/`, `spec/`, `builder/`, `bootstrap/` or `tests/` is
touched by this document.

---

## 1. Phase objective

Convert a structurally valid workbook into reproducible **SAR nominal and PV cost
measures**, analytically and deterministically, and do it in a form the later
Monte Carlo phase can reuse without touching a worksheet inside an iteration loop.

Phase 5 produces numbers. It produces no random numbers.

---

## 2. Mapping to the locked architecture roadmap

| Implementation phase | Layer | Status |
|---|---|---|
| 1–3 | Workbook skeleton, Setup/Config inputs, driver registers | closed |
| **4** | **Structural runtime** — permanent IDs, timeline application, profiling and inflation synchronisation, add/delete, rollback, Stage-B runtime harness | **closed** |
| **5** | **Calculation / `_Calc` factor engine** — FX resolution, inflation factors, discount factors, profiling weights, `Knom` / `Kpv`, deterministic base, mean-basis base, expected risk, analytical annual cash flow | **this phase** |
| 6+ | RNG (MRG32k3a), sampling, simulation, percentiles, contingency, sensitivity, Results, Dashboard, Model Check UI | not started |

Phase 5 is **not** the RNG / Monte Carlo validation gate. The word "mean" appears
throughout this document in its **analytical** sense — the mathematical expected
value of a distribution — never as a simulation output.

---

## 3. Scope

### In scope

- resolved FX to SAR, per driver
- annual inflation factors, calendar-year anchored
- annual discount factors, project-year indexed
- cost-line profiling weights by permanent ID
- risk profiling weights by permanent ID
- precomputed nominal factor `Knom` per driver
- precomputed PV factor `Kpv` per driver
- Escalated Deterministic Base Estimate — Nominal SAR
- Escalated Deterministic Base Estimate — PV SAR
- Mean-Basis Base Cost — Nominal SAR
- Mean-Basis Base Cost — PV SAR
- Expected Risk / EMV — Nominal SAR
- Expected Risk / EMV — PV SAR
- analytical mean total reconciliation inputs
- per-year analytical cash-flow components (mean basis + expected risk)
- refusal behaviour when any required numerical input is invalid
- a worksheet-independent numerical kernel the simulation phase will reuse

### Out of scope — explicitly

MRG32k3a implementation · seed derivation · RNG stream identity · random variate
generation · Bernoulli occurrence simulation · Triangular sampling · Beta-PERT
sampling · Uniform sampling · simulation iterations · percentile ladder ·
P10/P50/P90 · Selected Px calculation · contingency · histogram · CDF ·
sensitivity / Spearman · `_SimData` iteration storage · Dashboard finalisation ·
Results finalisation · Model Check final UI · annual percentiles · selected-Px
annual profiles · simulation reconciliation.

No Monte Carlo output is permitted in Phase 5.

---

## 4. Mathematical definitions

Locked, and preserved exactly.

### 4.1 Deterministic central value — risks excluded

Per Cost Line, from its distribution:

| Distribution | Deterministic central unit cost |
|---|---|
| Triangular | `ML` |
| Beta-PERT | `ML` |
| Uniform | `(Min + Max) / 2` |

Then:

```
central unit cost × Quantity × FX to SAR × profiled inflation
```

Headline label: **Escalated Deterministic Base Estimate (Nominal SAR)**, with a
PV counterpart. The row-level basis must visibly read **ML / Midpoint**.

**This is never called "mean".**

### 4.2 Distribution expected value — the mean basis

| Distribution | Expected value |
|---|---|
| Triangular | `(Min + ML + Max) / 3` |
| Beta-PERT, λ = 4 | `(Min + 4·ML + Max) / 6` |
| Uniform | `(Min + Max) / 2` |

Applied to Quantity, FX, profile, inflation and discounting → **Mean-Basis Base
Cost**, Nominal and PV. It exists primarily to reconcile against the later
simulation mean.

Note the two bases coincide for Uniform and differ for Triangular and Beta-PERT.
That is expected and is itself an oracle case (§22, case 7).

### 4.3 Expected Risk / EMV

Per risk: `Probability × Expected Severity`, where expected severity uses the same
three formulas above on `impact_min` / `impact_most_likely` / `impact_max`. Then
FX × risk profile × inflation, and discounting for PV.

**All entered risks are included analytically.** No selection, no filtering.
`Probability` is stored as a fraction in `[0, 1]` (`0.0%` display format,
validated `between 0 and 1`), so it is used directly with no ÷100.

Occurrence probability lives on the Risk Register and is **entirely separate**
from the risk profiling allocation percentages, which distribute severity across
project years.

---

## 5. Inflation convention

Base Year is the price base. `Base Year ≤ Start Year`.

For spend in calendar year `Y`:

```
inflation_factor(Y) = 1                                  if Y = BaseYear
                    = Π  (1 + rate_k)   for k = BaseYear+1 … Y   otherwise
```

Worked example — Base 2026, spend 2029:

```
(1 + rate_2027) × (1 + rate_2028) × (1 + rate_2029)
```

If `BaseYear = StartYear`, the first project-year factor is **1** (empty product).

This is exactly consistent with the locked structural span already implemented in
Phase 4:

> `nmInflFirstYear` — `=IF(nmBaseYear_Applied="","",nmBaseYear_Applied+1)`
> "Escalation applies from the year after the applied base year."
> `nmInflLastYear` — `=IF(nmLastYear_Applied="","",nmLastYear_Applied)`

So the required rate span is `BaseYear+1 … LastProjectYear`, and when
`BaseYear < StartYear` that span deliberately includes calendar years **before**
the project starts. Those rates are required and are generated by Apply.

Inflation profiles are **calendar-year anchored**. A start-year shift uses the
rates of the new calendar years; values never move positionally. Phase 4 proved
this on target.

**A missing required rate never becomes zero.** The structure contract seeds
inflation year cells as `initial_value: null` precisely so that an unmade
assumption cannot be fabricated as 0%:

> "BLANK, never zero. A new annual escalation assumption the user has not made
> must not be fabricated as 0%."

Phase 5 **refuses** rather than manufacturing a value.

---

## 6. Discounting convention

Uses `inpDiscountRate` (Setup C20, `required: true`, `0.00%`).

```
PV factor for project-year index t  =  1 / (1 + DiscountRate)^(t - 1),  t starting at 1
```

so project year 1 → `1`, year 2 → `1/(1+r)`, year 3 → `1/(1+r)²`.

### Conflict check against the locked architecture

The locked `input_contract.yaml` says of Project Start Year:

> `note: "Whole calendar year. Discounting will treat this as period 0."`

Project year 1 **is** the Project Start Year, and exponent `t-1 = 0` there. The
locked contract and the requested convention **agree**. No conflict, and no new
timing convention has been chosen silently.

### One terminological flag — not a numerical one

The request labels this *"end-of-project-year discounting from the project
start"*. Arithmetically, discounting year 1's cash flow by zero periods is a
**start-of-period** (beginning-of-year) convention; a true end-of-year convention
would give year 1 a factor of `1/(1+r)¹`. The **arithmetic requested is
unambiguous and is what this plan adopts** — exponent `t-1`, matching the locked
"period 0" note. Only the label is inconsistent, and the Methodology sheet should
say *"discounted from the start of the project, Project Year 1 = period 0"*
rather than "end-of-year". Raised in §26 as a documentation decision, not a
calculation decision.

---

## 7. FX resolution

Locked convention: **1 source-currency unit = X SAR**. Constant across the
project. No FX uncertainty.

- `SAR` must resolve to exactly `1`. This is a model invariant already enforced at
  build time: `tblFXRates` row 1 is the locked seed `["SAR", 1]`.
- Every currency referenced by a Cost Line or Risk must resolve to **exactly one**
  valid positive rate in `tblFXRates`.
- A missing non-SAR rate is **never** defaulted to 1.
- Duplicate currency rows → refusal. Zero, negative or non-numeric rate → refusal.
- Blank rate cell for a referenced currency → refusal.

Resolution is a lookup by currency **name**, matching how the Setup FX table is
keyed. Rate cells are `#,##0.000000` with a `> 0` cell rule, but the cell rule
allows blanks and can be bypassed by paste — so Phase 5 revalidates.

---

## 8. Profiling semantics

Weights are applied **by permanent ID**, never by row position — proven on target
by Gate-B scenarios `B2` and `K2`.

For every identified driver:

```
Σ (profile weights over the applied project years) = 100%
```

is required **before** any numerical result is produced.

### Blank versus numeric zero

These are different, and the difference is load-bearing:

| Cell state | Meaning | Phase 5 |
|---|---|---|
| numeric `0` | the user has stated that this year carries **no** spend | valid; contributes 0 to the sum |
| blank | the user has stated **nothing** about this year | see §26 decision D4 |

The structure contract seeds profiling year cells with `initial_value: 0`, so a
freshly generated grid is complete and sums to 0% until the user fills it. A
blank in a profiling grid therefore means a cell that was **deliberately cleared**
— it is not the birth state. Inflation is the opposite: seeded `null`, because an
unmade escalation assumption must not be fabricated.

A row of all-blanks sums to 0% and is refused by the 100% rule regardless. The
open question is only a row that reaches 100% **with** blanks present (§26, D4).

### Applied timeline only

The engine uses the project-year columns of the **APPLIED** timeline
(`nmStartYear_Applied`, `nmDuration_Applied`, `nmYearCount_Applied`). Entered-but-
not-applied values must never drive calculation.

If `nmStructuralState` reads `STRUCTURE CHANGE PENDING`, calculation is
**invalid / stale / not runnable** and is refused. Same for
`No timeline applied`. Both labels come from the manifest
(`state_labels.pending`, `state_labels.not_applied`), not from restated strings.

---

## 9. Precomputed factors `Knom` and `Kpv`

Per driver `i`, over applied project years `y = 1 … N`:

```
Knom_i = FX_i × Σ_y ( w_{i,y} × infl_y )
Kpv_i  = FX_i × Σ_y ( w_{i,y} × infl_y × disc_y )
```

where `infl_y` is the inflation factor of the **calendar** year of project year
`y` under the driver's inflation profile, and `disc_y = 1/(1+r)^(y-1)`.

Then:

```
Cost line, nominal = unit_cost × Quantity × Knom
Cost line, PV      = unit_cost × Quantity × Kpv
Risk,  nominal     = severity  × Knom
Risk,  PV          = severity  × Kpv
```

`unit_cost` is the deterministic central value, the distribution mean, or (later)
a sampled draw — the factor is identical in all three cases. Probability is
handled **separately**: multiplied in for analytical EMV, and replaced by a
Bernoulli draw in Monte Carlo. It is deliberately **not** folded into `Kpv`.

This is the whole point of the optimisation: the kernel resolves worksheets once
and never traverses them inside a simulation loop.

---

## 10–12. The three headline measures

Let `Q_i` = Quantity, `c_i` = deterministic central, `m_i` = distribution mean,
`p_j` = probability, `s_j` = expected severity.

```
A_nom = Σ_i  c_i · Q_i · Knom_i          Escalated Deterministic Base — Nominal
A_pv  = Σ_i  c_i · Q_i · Kpv_i           Escalated Deterministic Base — PV

C_nom = Σ_i  m_i · Q_i · Knom_i          Mean-Basis Base Cost — Nominal
C_pv  = Σ_i  m_i · Q_i · Kpv_i           Mean-Basis Base Cost — PV

D_nom = Σ_j  p_j · s_j · Knom_j          Expected Risk / EMV — Nominal
D_pv  = Σ_j  p_j · s_j · Kpv_j           Expected Risk / EMV — PV
```

---

## 13. Annual analytical cash flow

Per applied project year `y`, six values:

```
Base Cost   — Nominal      Σ_i  m_i · Q_i · FX_i · w_{i,y} · infl_{i,y}
Expected Risk — Nominal    Σ_j  p_j · s_j · FX_j · w_{j,y} · infl_{j,y}
Total       — Nominal      the two above
Base Cost   — PV           same, × disc_y
Expected Risk — PV         same, × disc_y
Total       — PV           the two above
```

**Base Cost in the annual cash flow uses the distribution expected value `m_i`,
not the deterministic ML/Midpoint basis `c_i`.** The locked Results requirement is
that annual cash flow is mean-only, so the annual series is *Mean-Basis Base Cost
+ Expected Risk*. The deterministic basis has no annual series.

By construction `Σ_y (annual nominal total) = C_nom + D_nom`, and likewise for PV
— an internal identity worth auditing (§14).

No annual percentiles. No selected-Px annual profile.

---

## 14. Reconciliation identities

Locked relationship:

```
A  Escalated Deterministic Base
+ B  Uncertainty Mean Shift
= C  Mean-Basis Base Cost

C  + D  Expected Risk  =  E  Analytical Mean Total
```

### A design point worth stating plainly

If `B` is *defined* as `C − A`, then `A + B = C` is a tautology and auditing it
proves nothing. To make it a real check, Phase 5 computes `B` **independently**:

```
B_nom = Σ_i (m_i − c_i) · Q_i · Knom_i          B_pv = Σ_i (m_i − c_i) · Q_i · Kpv_i
```

accumulated in its own pass, and then audits `|A + B − C| ≤ tol`. The same applies
to `E`: it is accumulated independently per driver rather than as `C + D`, so
`|C + D − E| ≤ tol` is a genuine identity.

Audited identities, Nominal and PV each:

| # | Identity | Why it can fail |
|---|---|---|
| I1 | `A + B = C` | a distribution mean/central mismatch, or a driver counted in one pass and not the other |
| I2 | `C + D = E` | a risk or cost line missing from one accumulation |
| I3 | `Σ_y annual_total = E` | a profiling weight applied in the total but not in the annual split |
| I4 | `Σ_y w_{i,y} = 1` per driver | profiling validation |
| I5 | `A_pv ≤ A_nom` when `r ≥ 0` and all `infl ≥ 0` | a discount factor applied with the wrong sign or index |

`Simulation Mean ≈ E` and `Selected Px − E` are **later** phases. Not implemented,
not asserted, no statistical tolerance invented here.

---

## 15. Workbook / `_Calc` design

`_Calc` is `hidden` (not veryHidden) so an auditor can inspect it. `_SimData`
remains `veryHidden` and is untouched by Phase 5. User-facing input sheets stay
clean — Phase 5 writes to none of them.

`_Calc` must not become a dumping ground. Each block is declared in a contract
(proposed: `spec/calc_contract.yaml`, a fifth authority in the established
pattern) with **purpose, ownership, source inputs, output units, update trigger,
validation rule** — the six fields required, per block.

Proposed blocks:

| Block | Purpose | Owner | Source inputs | Output units | Update trigger | Validation |
|---|---|---|---|---|---|---|
| existing counters (C10, C11) | permanent-ID counters | Phase 4 | — | integer | Add | Phase 4's `counter_integrity` |
| `calc_state` | last calculation stamp, applied-triple fingerprint, refusal reason | Phase 5 | applied triple, structural state | text/integer | Calculate | fingerprint matches applied triple |
| `calc_years` | per applied project year: calendar year, inflation factor per profile, discount factor | Phase 5 | applied triple, `tblInflation`, `inpDiscountRate` | factor, dimensionless | Calculate | every required factor resolved and finite |
| `calc_fx` | resolved rate per referenced currency | Phase 5 | `tblFXRates`, driver currencies | SAR per unit | Calculate | exactly one positive rate per referenced currency; SAR = 1 |
| `calc_drivers` | per permanent ID: `Knom`, `Kpv`, central, mean, quantity/probability | Phase 5 | registers, profiling grids, `calc_years`, `calc_fx` | SAR-factor / SAR | Calculate | profile sums to 100%; distribution valid |
| `calc_totals` | A, B, C, D, E — Nominal and PV | Phase 5 | `calc_drivers` | SAR | Calculate | identities I1–I3 within tolerance |
| `calc_annual` | six series over applied project years | Phase 5 | `calc_drivers`, `calc_years` | SAR | Calculate | I3 |

**Auditability, not computation.** These blocks are a *written record* of what the
in-memory kernel computed. Nothing reads them back to compute anything else; that
would recreate the worksheet dependency the design exists to avoid.

Everything on `_Calc` is model-controlled and carries the locked visual treatment,
consistent with Phases 1–4.

---

## 16. VBA / numerical module boundaries

The hard rule: **mathematical functions must not read worksheet cells.**
Structural and presentation modules may; numerical functions receive resolved
numeric inputs and return resolved numeric outputs. This is what lets Phase 6 call
the same functions 100,000 times without an Excel round trip.

Existing Phase-4 modules are unchanged. Proposed additions:

| Module | Layer | Responsibility | May touch worksheets |
|---|---|---|---|
| `modCalcContract` | generated | Phase-5 constants projected from `calc_contract.yaml` — block addresses, tolerances, labels | n/a (constants) |
| `modCalcResolve` | resolution | reads Setup, registers, profiling grids, `tblFXRates`, `tblInflation`; produces plain numeric arrays and Types | **yes** |
| `modCalcFactors` | numerical | `InflationFactors`, `DiscountFactors`, `BuildKnom`, `BuildKpv` | **no** |
| `modCalcAnalytical` | numerical | `TriangularMean`, `PertMean`, `UniformMean`, `DeterministicCentral`, `ExpectedRisk`, the A/B/C/D/E accumulations, the annual series | **no** |
| `modCalcReport` | presentation | writes `_Calc` blocks, reports refusals through the Phase-4 `modAppState` result surface | **yes** |
| `modCalcCheck` | validation | Phase-5 numerical prerequisites; returns a report, never repairs | **yes** (reads only) |

The names `modCalcFactors` / `modCalcAnalytical` from the request are adopted
because they fit the existing `modXxx` convention and the existing
report-never-repair split (`modStructuralCheck`). `modCalcResolve` and
`modCalcReport` are added because the resolution/calculation boundary the request
demands needs a *named* place to live on each side of it — folding resolution into
`modCalcFactors` would breach the rule the request sets.

A static test will enforce the boundary: **no worksheet-touching identifier**
(`ThisWorkbook`, `Worksheets`, `Range`, `ListObjects`, `Cells`, `modWorkbook.*`)
may appear in `modCalcFactors` or `modCalcAnalytical`. That is mechanically
checkable on Linux and becomes a permanent sweep, in the style already
established.

---

## 17. Data structures for later simulation reuse

Resolved once, held in memory, reused per iteration:

```vb
Type DriverFactors            ' one per Cost Line and per Risk
    PermanentId   As String
    IsRisk        As Boolean
    Knom          As Double
    Kpv           As Double
    Quantity      As Double   ' 1 for risks
    Probability   As Double   ' 1 for cost lines
    DistKind      As Long     ' Triangular | BetaPert | Uniform
    MinValue      As Double
    MostLikely    As Double
    MaxValue      As Double
    Central       As Double   ' ML or midpoint
    MeanValue     As Double   ' distribution expected value
End Type

Type YearFactors              ' one per applied project year
    ProjectIndex  As Long
    CalendarYear  As Long
    DiscountF     As Double
End Type
```

plus, for the annual split, a per-driver `Double` array of
`w_{i,y} × infl_{i,y}` of length `N` — the only per-year storage the kernel needs.

Simulation later needs **only** `DriverFactors` and that weight array: no
worksheet, no ListObject, no Range. `MinValue`/`MostLikely`/`MaxValue` are carried
so the sampler has its parameters without re-reading anything.

---

## 18. Validity and failure behaviour

Phase 5 refuses; it never manufactures a value and never repairs.

### Structural prerequisites — already owned by Phase 4, re-checked only by asking

Phase 5 calls the existing gate rather than duplicating it:

- structural state is `STRUCTURE CHANGE PENDING` → refuse
- no timeline applied → refuse
- `modStructuralCheck.ValidateStructure()` non-empty → refuse and quote it
- duplicate permanent ID · orphan profiling row · profiling ID missing from
  register · register ID missing from profiling — **all Phase 4**, surfaced by the
  same call

### Phase-5 numerical prerequisites — new, and owned here

| Check | Refusal |
|---|---|
| Base Year > Start Year | yes |
| Discount Rate blank or non-numeric | yes |
| Discount Rate makes `1 + r ≤ 0` | yes |
| referenced currency missing from `tblFXRates` | yes |
| duplicate currency in `tblFXRates` | yes |
| FX rate ≤ 0, blank or non-numeric | yes |
| SAR ≠ 1 | yes |
| referenced inflation profile missing from `tblInflation` | yes |
| required inflation year missing (blank) | yes |
| inflation rate non-numeric | yes |
| inflation rate makes `1 + rate ≤ 0` | yes (§26 D2) |
| cost profile non-numeric | yes |
| cost profile ≠ 100% within tolerance | yes |
| risk profile non-numeric or ≠ 100% | yes |
| distribution missing or not one of the three | yes |
| `Min ≤ ML ≤ Max` violated (Triangular, Beta-PERT) | yes |
| `Min ≤ Max` violated (Uniform) | yes |
| Uniform with a populated ML | §26 D1 |
| Quantity missing or non-numeric | yes |
| Probability missing, non-numeric, or outside `[0,1]` | yes |

Later, Model Check aggregates all of this. **Phase 5 must be able to refuse an
invalid calculation independently**, and does — through the Phase-4 result surface
(`modAppState.Failed` / `Announce`), so refusal reporting is one mechanism, not
two.

---

## 19. Numerical tolerances

| Purpose | Tolerance | Rationale |
|---|---|---|
| profiling sum = 100% | `|Σw − 1| ≤ 1e-9` absolute | percentages are entered to 2 dp and stored as binary doubles; over the 200-column structural maximum the accumulated representation error is bounded by roughly `200 × 2⁻⁵² ≈ 4.4e-14`. `1e-9` is ~4 orders of magnitude of headroom and still four orders tighter than a 1-in-a-million data-entry slip (`1e-6`). Exact binary equality is never used. |
| reconciliation identities I1–I3 | `|Δ| ≤ max(1e-6 SAR, 1e-12 × scale)` | absolute floor for small models, relative term for large ones, where `scale = max(|A|,|C|,|E|)`. A pure relative test degenerates at zero; a pure absolute test fails at billions. |
| FX positivity | `rate > 0`, no epsilon | a rate is either present and positive or it is refused |
| `1 + rate > 0` | strict | a factor of zero or below is not a price relationship |

No Monte Carlo statistical tolerance is invented here. That belongs with the phase
that produces a sampling distribution.

---

## 20. Gate A — Linux / static

Phase 4 established that static tests cannot prove Excel behaviour, and equally
that they catch a great deal before Windows time is spent. Gate A covers:

1. **A pure-Python numerical oracle** (`builder/pccm_builder/calc_oracle.py`) that
   implements every formula in §4–§13 independently, in the established
   `structure_oracle.py` pattern: it both defines the semantics the VBA must match
   **and** generates the expected values the Windows harness asserts, so the two
   cannot drift.
2. **Golden case tests** — §22, every value hand-derived and written in the test,
   never produced by the code under test.
3. **Contract validation** for `calc_contract.yaml`, in the pattern of the four
   existing contract validators.
4. **Source sweeps**, extending the existing mechanical ones:
   - no worksheet identifier in `modCalcFactors` / `modCalcAnalytical`
   - no Phase-6 construct (`Rnd`, `MRG`, `Percentile`, `Iteration`, `Sample`)
     anywhere in Phase-5 VBA
   - VBA block balance, line length, declaration-section placement, `Optional`
     defaults — the existing sweeps applied to the new modules
   - every emitted constant referenced, no structural literal restated
5. **Post-build verification** extended with the `_Calc` block layout.

Gate A ends with a source review, exactly as Phase 4 did.

---

## 21. Gate B — real Windows / Excel

Extends the accepted `phase4_functional_test.ps1` matrix rather than replacing it:
the whole Phase-4 matrix must continue to pass, and new scenarios are appended.

Functional validation of:

- FX resolution, including a foreign currency and the SAR identity
- inflation compounding, including `BaseYear = StartYear` and `BaseYear < StartYear`
- discount factors at project-year indices 1, 2, 3
- profiling factor application by permanent ID
- deterministic base, Nominal and PV
- mean-basis base, Nominal and PV
- expected risk, Nominal and PV
- analytical annual cash flow, all six series
- reconciliation identities I1–I3 asserted in the workbook
- calculation refusal on every §18 numerical prerequisite
- calculation refusal when `STRUCTURE CHANGE PENDING`
- clean Excel shutdown and clean transient COM release

Every expected value comes from `build/phase5_cases.json`, emitted by the oracle —
the harness hardcodes no number, exactly as `phase4_scenarios.json` works today.
**The harness asserts every calculated value; the user inspects no cells manually.**

The Phase-4 harness disciplines carry over unchanged and are non-negotiable:
caller-side `@(...)`, one pipeline object per row, container factories emitted
non-enumerated, `catch` attached to its `try`, keyed-only fixtures, failure-safe
cleanup, per-scenario clean-structure prerequisites, and `$excelIdentity`.

---

## 22. Golden hand-calculated cases

Every expected value below is derived by hand and stated here. The oracle must
reproduce these; they are not generated from it.

Shared unless overridden: Triangular `Min 80 / ML 100 / Max 150`, `Quantity 10`,
currency SAR, one profile at 100%.

| # | Case | Setup | Hand-derived expected |
|---|---|---|---|
| 1 | SAR, no inflation, one project year | Base 2026, Start 2026, Dur 1; inflation span empty (`2027 > 2026`); profile `100%`; `r = 10%` | `infl = 1`, `disc₁ = 1`, `Knom = Kpv = 1`. `A = 100×10 = 1000`. `C = 110×10 = 1100`. `B = 100`. `A_pv = 1000`, `C_pv = 1100` |
| 2 | foreign currency | as case 1, currency USD, `FX = 3.75`, unit cost 100 USD, Qty 4 | `Knom = 3.75`; `A = 100 × 4 × 3.75 = 1500 SAR` |
| 3 | multi-year profiling, compounded inflation | Base 2026, Start 2027, Dur 3; rates 2027/28/29 = 5%; profile `20/50/30` | `f = 1.05, 1.1025, 1.157625`. `Knom = 0.2(1.05)+0.5(1.1025)+0.3(1.157625) = 0.21+0.55125+0.3472875 = ` **`1.1085375`**. `A_nom = 100×10×Knom = ` **`1108.5375`** |
| 4 | PV across multiple years | case 3 with `r = 10%`; `disc = 1, 1/1.1, 1/1.21` | `Kpv = 0.21 + 0.501136363636 + 0.287014462810 = ` **`0.998150826446`**. `A_pv = ` **`998.150826446`**. `C_pv = 110×10×Kpv = ` **`1097.965909091`** |
| 5 | Triangular deterministic vs mean | `80/100/150` | central `100`; mean `(80+100+150)/3 = ` **`110`**. On case 3: `C_nom = 110×10×1.1085375 = ` **`1219.39125`**; `B_nom = 10×10×1.1085375 = ` **`110.85375`** |
| 6 | Beta-PERT deterministic vs mean | `80/100/150`, λ = 4 | central `100`; mean `(80+400+150)/6 = 630/6 = ` **`105`** |
| 7 | Uniform midpoint = Uniform mean | `Min 80 / Max 150` | central `(80+150)/2 = 115`; mean `115`. **`B = 0` for this driver** — the two bases coincide |
| 8 | risk EMV, probability < 1 | `P = 30%`, Triangular severity `100/200/450`, case-1 factors | mean severity `(100+200+450)/3 = 250`; `D = 0.30 × 250 × 1 = ` **`75`** |
| 9 | multi-year risk profile | risk of case 8 on case-3 factors | `D_nom = 75 × 1.1085375 = ` **`83.1403125`**; `D_pv = 75 × 0.998150826446 = ` **`74.8613119835`** |
| 10 | Base Year = Start Year | Base 2027, Start 2027, Dur 2; rate 2028 = 5% | `infl(2027) = 1`, `infl(2028) = 1.05` — **first project year factor is 1** |
| 11 | Base Year earlier than Start Year | case 3 | first project year (2027) already carries `1.05`, because 2027 = Base+1 |
| 12 | zero inflation | rates all `0%` | every `infl = 1`; `Knom = FX`; `Kpv = FX × Σ(w·disc)` |
| 13 | negative but valid inflation | rate `−2%` for three years | `0.98, 0.9604, 0.941192` — arithmetically sound while `1+rate > 0` (§26 D2) |
| 14 | blank required inflation rate | case 3 with 2028 blank | **refusal**, naming the profile and calendar year 2028. No value produced |
| 15 | profile sum ≠ 100% | case 3 with profile `20/50/20` (= 90%) | **refusal**, naming the permanent ID and the sum. No value produced |

Cases 14 and 15 assert a **refusal**, not a number — the failure surface is part of
the contract, and cases that refuse must produce no partial result.

---

## 23. Performance

Design target: **200 Cost Lines, 100 Risks, 25 project years, 100,000 iterations.**

Phase 5 runs no iterations, but its output shape determines Phase 6's cost.

| Quantity | Complexity |
|---|---|
| resolution (worksheet reads) | `O(D × N)` = 300 × 25 = 7,500 cells, **once** |
| `Knom` / `Kpv` build | `O(D × N)` = 7,500 multiply-adds, **once** |
| analytical totals A–E | `O(D)` = 300 |
| annual series | `O(D × N)` = 7,500 |
| later, per simulation iteration | `O(D)` = 300 — **no worksheet access** |

Memory: `DriverFactors` is ~11 fields × 300 ≈ 26 KB; the per-driver weight arrays
are `300 × 25 × 8 B = 60 KB`; `YearFactors` is negligible. **Well under 100 KB
resident**, which is why the simulation phase can keep it all in memory for the
entire run.

The decisive property: at 100,000 iterations the kernel performs ~30 million
multiply-adds and **zero** Excel round trips. A worksheet-reading design would
perform ~750 million COM calls, which is the difference between seconds and hours.

---

## 24. Implementation sequence inside Phase 5

1. `spec/calc_contract.yaml` — the fifth authority; block layout, tolerances,
   labels, refusal messages. Loader + validator, fail-loud, in the established
   pattern.
2. `builder/pccm_builder/calc_oracle.py` — pure-Python implementation of §4–§13.
   Golden-case tests written **first**, from the hand derivations in §22.
3. Stage-A emission — `_Calc` blocks per the contract; `modCalcContract.bas`
   generated; `build/phase5_cases.json` emitted from the oracle. Post-build
   verification extended.
4. `modCalcFactors` + `modCalcAnalytical` — the pure numerical kernel, with the
   no-worksheet sweep active from the first commit.
5. `modCalcResolve` — worksheet → numeric structures.
6. `modCalcCheck` — Phase-5 numerical prerequisites; reports, never repairs.
7. `modCalcReport` + a `PCCM_Calculate` command button — orchestration, `_Calc`
   write-back, refusal reporting through `modAppState`.
8. Gate-A source review.
9. Gate-B harness extension and, on approval, the Windows run.

Steps 1–8 are Linux-only. No Windows execution is requested before Gate A is
approved — the discipline that Phase 4 established and that this plan keeps.

---

## 25. Acceptance criteria

Phase 5 is complete when **all** hold:

1. every Phase 1–4 Linux/static test still passes, none weakened;
2. post-build verification passes, extended to the `_Calc` blocks;
3. every golden case in §22 passes against the oracle **and** on Windows, with the
   refusal cases producing no partial result;
4. reconciliation identities I1–I3 hold within §19 tolerances, Nominal and PV,
   with `B` and `E` computed independently;
5. calculation is refused, with a specific message, for every §18 numerical
   prerequisite and for `STRUCTURE CHANGE PENDING`;
6. `modCalcFactors` and `modCalcAnalytical` contain no worksheet access, proven by
   sweep;
7. the Gate-B harness asserts every calculated value with no manual inspection,
   and the full Phase-4 matrix still passes;
8. Excel shuts down naturally with clean transient COM release;
9. no Monte Carlo artefact exists anywhere in the phase.

---

## 26. Unresolved decisions — required before implementation

These are stated rather than silently chosen.

**D1 — Uniform with a populated ML.** The Uniform central value and mean are both
`(Min+Max)/2`; `most_likely` is unused. Is a populated ML for a Uniform driver
(a) refused, (b) an advisory that still calculates, or (c) silently ignored?
*Recommendation: (b) — advisory, not refusal.* The value is harmless
arithmetically, refusing it would block a user who switched distribution mid-entry,
and silence would hide a likely mistake. Model Check is the natural home for the
advisory; Phase 5 would surface it as a note alongside a successful result.

**D2 — Inflation rate lower bound.** Deflation is legitimate. Is any rate with
`1 + rate > 0` accepted (so `−99.9%` is valid), or is there a business floor?
*Recommendation: accept any rate with `1 + rate > 0`, refuse at or below.* No
business floor is invented. Case 13 depends on this answer.

**D3 — Discount rate sign.** `inpDiscountRate` has no validation and no business
bound (locked: "No business minimum or maximum is imposed yet"). Is a negative
discount rate accepted? *Recommendation: accept while `1 + r > 0`, consistent with
D2.* Then `A_pv > A_nom`, which is arithmetically correct and would break identity
I5 as stated — so I5 must be conditioned on `r ≥ 0`, as written in §14.

**D4 — Blank profiling cell in a row that otherwise sums to 100%.** Options:
(a) refuse — blank is not a stated zero; (b) treat as zero for the sum, since the
row already sums correctly. *Recommendation: (a) refuse.* Phase 4 deliberately
preserves a blank as invalid-but-surviving data specifically so a later phase can
report it, and profiling cells are **born as `0`**, so a blank is a deliberate
clearing rather than an unfilled default. Treating it as zero would erase the
distinction Phase 4 was built to keep. This is the one recommendation most worth
challenging, because it makes a row that "looks right" refuse.

**D5 — "End-of-year" label versus `t−1` arithmetic** (§6). The arithmetic is
settled and matches the locked "period 0" note; only the Methodology wording is
open. *Recommendation: describe it as "discounted from the start of the project;
Project Year 1 = period 0", and avoid the phrase "end-of-year".*

**D6 — Rounding and display precision of headline SAR measures.** Not specified
anywhere in the locked contracts. *Recommendation: compute and store at full
double precision, round only for display, and define the display format in
`calc_contract.yaml` rather than in VBA.*

None of D1–D6 blocks the design. Each changes a specific, named behaviour, and
each is small enough to settle at review.

---

## Status

Model source, workbook artifacts, contracts, bootstrap, harness and Phase-4 tests
are unchanged by this document.

**PHASE 5 PLAN READY FOR REVIEW**
