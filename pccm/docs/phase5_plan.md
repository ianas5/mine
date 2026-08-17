# PCCM — Phase 5 plan: deterministic and analytical calculation engine

**Revision C.** Revision B was substantially accepted; its mathematical
definitions, D1–D6, referenced-only FX/inflation scope, fingerprint *coverage*,
reconciliation structure, no-Calculate-button rule, Gate-A/Gate-B split and golden
matrix are all preserved. This revision makes the design implementation-safe:
the fingerprint serialisation becomes length-prefixed and locale-invariant with
every constant locked and a computed test literal; overflow protection moves from
a post-hoc predicate to controlled arithmetic primitives with numerically stable
formulas; calculation status separates the last successful snapshot from the last
attempt; write-back becomes transactional with logical rollback; the `_Calc`
anchors and schemas stop being illustrative; and reconciliation tolerances gain a
conditioning scale that survives cancellation.

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

**Implementation rule.** If `Format$` is used internally, the result is normalised
by replacing the locale decimal separator
(`Application.International(xlDecimalSeparator)`) with `.` **before** hashing, and
a Gate-A test asserts the normalised output against the literals below.

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

**Recurrence.** For each UTF-16 code unit `u` of the stream, in order, with `u`
normalised to `0 … 65535`:

```
h1 = (h1 * 131 + u) mod 2147483647
h2 = (h2 * 131 + u) mod 2147483629
```

Both accumulators start at `1`, so a stream beginning with `NUL` is not absorbed.

**Exactness in VBA.** The largest intermediate is
`(2147483647 − 1) × 131 + 65535 = 281,320,423,161`, far inside the `2⁵³ =
9,007,199,254,740,992` exact-integer range of a `Double`. The recurrence is
therefore exact in VBA arithmetic **without** unsigned 64-bit support, and
reproduces bit-identically in Python.

**Hashed stream.** Tags, lengths, the colon and values are **all** hashed — the
entire canonical stream, byte for byte, nothing excluded. This is what makes the
length prefixes structurally meaningful to the digest.

**Section ordering — fixed, not sorted:** `HEADER`, `COST`, `RISK`. Phase 6
appends its sections after these; the analytical sections keep their positions so
the analytical subset stays comparable across phases.

**Driver-record ordering:** ascending by **Permanent ID**, ordinal comparison on
UTF-16 code units (`Option Compare Binary` semantics, `StrComp(..., vbBinaryCompare)`).
Never by row, never by digest.

**Final representation:** `HEX8(h1) & HEX8(h2)` — 16 characters, **uppercase**,
zero-padded to 8 each.

### 11.5 The locked reference vector

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

**Python and VBA must both produce this exact literal.** It is a locked test
vector on both sides of Gate A and is re-asserted on real Excel at Gate B (§24).

### 11.6 Collision probes — locked test set

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

All eight are distinct. Under Revision B's `U+001F` join, rows 1–2, 3–5 and 4–8
would have been indistinguishable.

### 11.7 `calc_state` — successful snapshot separated from last attempt

Revision B conflated the stored status, the on-demand status, a refused attempt
and the fingerprint comparison. **They are now separate concepts.**

| Field | Meaning |
|---|---|
| `Last Successful Fingerprint` | fingerprint of the inputs of the last calculation that **completed and committed** |
| `Last Successful Stamp` | its timestamp |
| `Last Successful Applied Timeline` | the applied triple it used |
| `Fingerprint Version` | `FP_VERSION` at the time of that success |
| `Last Attempt Result` | `SUCCESS` / `REFUSED` — the outcome of the most recent explicit `PCCM_Calculate` |
| `Last Refusal Reason` | populated only when the last attempt was `REFUSED` |
| `Status (last evaluated)` | the derived status as of the last evaluation (below) |

### 11.8 Derived status — LOCKED semantics

Status is **derived from the current input state compared with the last successful
fingerprint**, never from the attempt history:

```
no Last Successful Fingerprint
        and the most recent attempt refused   -> REFUSED   (with reason)
no Last Successful Fingerprint                -> NOT CALCULATED
current inputs cannot resolve                 -> INVALID   (with the resolution failure)
current fingerprint == Last Successful        -> CURRENT
current fingerprint != Last Successful        -> STALE
```

**The locked answer to the review's scenario.** Successful calculation → user
makes an input invalid → `PCCM_Calculate` refuses → user restores the input
*exactly* to the previously successful state → status is queried, with no
recalculation:

> **CURRENT.**

The inputs resolve, and their fingerprint equals the last successful fingerprint,
so the stored snapshot genuinely describes them. **A historical failed attempt
never permanently overrides a currently matching successful snapshot.** The
refusal reason remains visible in `Last Refusal Reason` as attempt history, but it
does not determine status.

`INVALID` is distinct from `STALE`: an unresolvable current state cannot produce a
fingerprint at all, so claiming "stale" would assert a comparison that was never
made. `CURRENT` is never returned in that case.

### 11.9 Status is last-evaluated, not live

**There are no change events** — no `Worksheet_Change`, no
`Workbook_SheetChange`, consistent with the Phase-4 rule that structural state is
never maintained by hidden automation.

The `_Calc` status cell therefore holds a **last-evaluated** status. It is
refreshed by:

- `PCCM_Calculate`,
- `PCCM_CalculationStatus`,
- later point-of-consumption guards (Run Check, Run Simulation, output refresh).

**It does not update spontaneously**, and the `_Calc` block is labelled
`Calculation Status (last evaluated)` so an auditor reading the sheet is not misled
into thinking otherwise.

### 11.10 Callable surface for Gate B

Public, invoked by `Application.Run` — **no button** (§17):

```
PCCM_Calculate                   orchestration; refuses cleanly; transactional (§12)
PCCM_CalculationStatus()         re-evaluates and returns the derived status
PCCM_CalculationFingerprint()    the LAST SUCCESSFUL fingerprint
PCCM_CurrentInputFingerprint()   the fingerprint of the inputs as they are NOW
PCCM_CalculationRefusal()        the last refusal reason, or empty
```

The two fingerprint accessors stay separate so Gate B can show the stored snapshot
**unchanged** while the current one has moved (§25).

---

## 12. Transactional write-back

Revision B required that a refused calculation not overwrite the previous
snapshot. **That is extended to write failures.** The Phase-5 audit blocks are one
snapshot, and a half-old / half-new `_Calc` is not an acceptable outcome.

### Locked orchestration

```
1. resolve everything into memory
2. validate everything
3. calculate everything in memory
4. reconcile everything in memory        (identities I1–I5)
5. build the complete fingerprint in memory
6. ONLY THEN begin workbook write-back
```

**No `_Calc` analytical result is written during steps 1–5.** A refusal at any of
those steps leaves the workbook byte-identical, so the previous successful
snapshot survives untouched — which is the Revision-B guarantee, now stated as a
consequence of ordering rather than as a separate rule.

### Rollback of a mid-write failure

Write-back covers `tblCalcFX`, `tblCalcYears`, `tblCalcInflationFactors`,
`tblCalcDrivers`, `tblCalcAnnual`, `calc_totals`, `calc_state`. A failure part way
through must not leave a mixed snapshot.

The strategy reuses **the Phase-4 logical-rollback mechanism already proven on
target** — `modWorkbook.SnapshotTable` / `RestoreTable`, including its
collision-safe header restoration:

1. before the first write, snapshot every Phase-5 `_Calc` block (the five tables
   plus the two scalar blocks);
2. write all analytical blocks;
3. **verify** the written blocks against the in-memory values;
4. on any failure at 2 or 3, restore every snapshot and report through
   `modAppState`, leaving `Last Successful *` untouched;
5. **write `Last Successful Fingerprint`, `Last Successful Stamp` and
   `Last Successful Applied Timeline` LAST**, only after every analytical block has
   been written and verified.

Because the success marker is written last, an interruption at any earlier point
leaves the snapshot marked with the **previous** success — never with a
half-written one. The commit point is a single scalar write.

`Last Attempt Result` and `Last Refusal Reason` are attempt metadata and are
written on both paths; they are not part of the snapshot and never make a failed
attempt look successful.

### Acceptance

A Windows injected-failure scenario (§25) fails **after** one or more `_Calc`
blocks have been mutated, and proves: previous totals restored · previous driver
audit rows restored · previous annual rows restored · previous successful
fingerprint and stamp restored · no mixed snapshot survives · Excel application
state restored. This is a **real acceptance requirement**, not a diagnostic.

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
magnitude of its net result. Per identity:

| Identity | Conditioning scale |
|---|---|
| I1 `A + B = C` | `max(1, \|A\| + \|B\| + \|C\|)` |
| I2 `C + D = E` | `max(1, \|C\| + \|D\| + \|E\|)` |
| I3a / I4a | `max(1, Σ_y \|annual base\| + \|C\|)` |
| I3b / I4b | `max(1, Σ_y \|annual risk\| + \|D\|)` |
| I3c / I4c | `max(1, Σ_y \|annual total\| + \|E\|)` |

Nominal and PV each use their own scale. Summing **absolute** annual
contributions is the point: it is the only term that still grows when the signed
annual values cancel.

The `max(1, …)` floor keeps the scale from going below unity for a genuinely tiny
model, so the `1e-6` absolute floor remains the binding constraint there.

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

| Row | Label | Value | Format |
|---|---|---|---|
| 13 | Calculation Status (last evaluated) | `NOT CALCULATED` / `CURRENT` / `STALE` / `INVALID` / `REFUSED` | `@` |
| 14 | Last Successful Stamp | timestamp | `yyyy-mm-dd hh:mm:ss` |
| 15 | Last Successful Fingerprint | 16 uppercase hex characters | `@` |
| 16 | Fingerprint Version | integer | `0` |
| 17 | Last Successful Applied Timeline | `base/start/duration` | `@` |
| 18 | Last Attempt Result | `SUCCESS` / `REFUSED` | `@` |
| 19 | Last Refusal Reason | empty unless the last attempt refused | `@` |
| 20 | Status Evaluated At | timestamp of the last status evaluation | `yyyy-mm-dd hh:mm:ss` |

Row 13's label says **"(last evaluated)"** because it is not live (§11.9). Row 20
makes that concrete: an auditor can see how old the status reading is.

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

**A static sweep enforces the boundary:** no worksheet identifier (`ThisWorkbook`,
`Worksheets`, `Range`, `ListObjects`, `Cells`, `modWorkbook.*`) may appear in
`modCalcFactors`, `modCalcAnalytical` or `modCalcFingerprint`. A second sweep
extends the Phase-4 `On Error Resume Next` whitelist across the Phase-5 modules,
so the only error handlers are the documented safe primitives of §19.1.
Mechanically
checkable on Linux, permanent, in the established style.

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
this document (§11.3–§11.6)  +  the tested Python/VBA implementations
                             +  the fixed test vectors 50B6EB0E26857EA7 and §11.6
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
| length-prefixed serialisation is collision-free | the eight §11.6 probes — strings containing `:`, `U+001F`, `U+0000`, `U+000A` — must yield eight distinct digests, asserted against the literals |
| UTF-16 code-unit parity | Python encodes `utf-16-le` and reads 16-bit units; test vectors include a non-BMP character, proving it contributes **two** units, and a character above `U+7FFF`, proving the `AscW` sign normalisation |
| numeric encoding is locale-invariant | the ten §11.3 literals asserted exactly; plus a test that the normalisation step maps a `,` decimal separator to `.` before hashing |
| exact hash constants and digest | `FP_BASE`, both moduli, both initial states asserted as literals; the §11.5 stream asserted at **366 code units** and its digest at **`50B6EB0E26857EA7`** |
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
| **27** | **delimiter-hostile field content** | the eight §11.6 probes | eight **distinct** digests, asserted against the §11.6 literals |
| **28** | **naive-overflow, representable result** | `Min = ML = Max = 1e308`, Triangular and Beta-PERT; midpoint with both `1.5e308` | **accepted**; stable forms return `1e308`, `1e308`, `1.5e308`. A naive implementation overflows and is thereby detected |
| **29** | **discount factor underflow** | `r = 1e10`, duration ≥ 34 | **controlled refusal** at project year 34, where the factor reaches exactly zero — never silently accepted |
| **30** | **cancellation-heavy reconciliation** | large positive and negative unit costs whose net is near zero, all representable | identities I1–I4 **hold**; the conditioning scale of §15 keeps the tolerance proportional to the arithmetic performed, not to the near-zero net |
| **31** | **Base-Year factor row** | `Base 2026, Start 2028, Dur 3` | `tblCalcInflationFactors` contains a `2026` row with **blank rate and cumulative factor `1`**, plus the pre-project rows `2027`, `2028` |
| **32** | **status reverts to CURRENT** | calculate → break an input → refuse → restore the input exactly → query | **`CURRENT`**, with no recalculation (§11.8) |
| **33** | **mid-write failure** | injected failure after `tblCalcDrivers` is mutated | full logical rollback; previous snapshot intact; no mixed state (§12) |

Cases 26–27 and 31 assert **audit content**; 28 and 32 assert **acceptance** where
a naive implementation would fail or over-block; 29 and 33 assert **controlled
refusal and rollback**.

---

## 24. Gate B — real Windows / Excel

Extends the accepted `phase4_functional_test.ps1` matrix; it does not replace it.

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
| **status reverts to CURRENT** | the §23 case 32 sequence, driven end to end |
| **refusal preserves the snapshot** | break an input, `PCCM_Calculate`, assert `REFUSED` with a reason, assert `calc_totals`, `tblCalcDrivers`, `tblCalcAnnual` and `Last Successful *` are **byte-for-byte the previous values** |
| **mid-write failure and full rollback** | inject a failure after one or more `_Calc` blocks have been mutated; assert previous totals, driver rows, annual rows, fingerprint and stamp all restored, no mixed snapshot, Excel application state restored. Uses the Phase-4 `FailPointCheck` mechanism already proven on target |
| **Base-Year factor visible** | `tblCalcInflationFactors` contains the `BaseYear` row with blank rate and cumulative factor `1` (§23 case 31) |
| **cancellation-heavy reconciliation** | a fixture with large offsetting contributions and a near-zero net; identities must hold (§23 case 30) |

Every expected value comes from `build/phase5_cases.json`. **The harness asserts
every calculated value; the user inspects no cells manually.**

Phase-4 harness disciplines carry over unchanged and are non-negotiable:
caller-side `@(...)`, one pipeline object per row, container factories emitted
non-enumerated, `catch` attached to its `try`, keyed-only fixtures, failure-safe
cleanup, per-scenario clean-structure prerequisites, `$excelIdentity`.

---

## 25. Gate-B stale-fingerprint scenario

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
that the previous successful snapshot was not overwritten (§12).

**Revert-to-CURRENT proof** — the locked §11.8 semantics, driven end to end:

12. from the `REFUSED` state above, restore the changed input **exactly** to the
    value it held at the last successful calculation;
13. query status **without** recalculating;
14. assert **`CURRENT`** — the inputs resolve and their fingerprint equals the
    last successful fingerprint, so a historical failed attempt must not keep the
    workbook marked stale;
15. assert `Last Refusal Reason` is still readable as attempt history, and that it
    did **not** determine the status.

**Fingerprint parity proof** — on the golden-case-1 fixture,
`PCCM_CurrentInputFingerprint()` must return the literal **`50B6EB0E26857EA7`**,
the same value the Python oracle produces. This is the assertion that proves the
two implementations agree on real Excel, including `AscW` sign normalisation and
locale-invariant numeric formatting.

**Mid-write rollback proof** — inject a failure after one or more `_Calc` blocks
have been mutated (§12), using the Phase-4 `FailPointCheck` mechanism already
proven on target. Assert previous totals, driver rows, annual rows, fingerprint
and stamp all restored; no mixed snapshot; Excel application state restored.

No `Worksheet_Change` or `Workbook_SheetChange` handler exists; status is computed
on demand (§11.9), and a sweep asserts neither handler was introduced. The `_Calc`
status cell is **last-evaluated**, and the harness reads it only after an explicit
`PCCM_Calculate` or `PCCM_CalculationStatus` call — never treating it as live.

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
   `modAppState`. `PCCM_Calculate` and the four accessors of §11.10. **No button.**
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
   digest is **`50B6EB0E26857EA7`** in Python, in VBA on Linux-side tests, and on
   real Excel; the eight collision probes are distinct; the ten numeric encodings
   match, locale-invariantly;
6. identities I1, I2, I3a–c, I4a–c, I5 hold within the §15 tolerances **using the
   per-identity conditioning scales**, with `B` and `E` independently accumulated.
   `A_pv ≤ A_nom` is a conditional diagnostic, **not** a gate;
7. calculation is refused, with a specific message, for every §18 numerical
   prerequisite and for `STRUCTURE CHANGE PENDING`; **no uncontrolled VBA overflow
   escapes, no overflow becomes a fabricated zero, and no discount factor
   underflows silently to zero**;
8. the stable formulas of §19.2 return the correct mean where the naive form
   overflows;
9. `modCalcFactors`, `modCalcAnalytical` and `modCalcFingerprint` contain no
   worksheet access, and no `On Error Resume Next` exists outside the documented
   safe primitives — both proven by sweep;
10. the fingerprint detects staleness for every covered input and **not** for
    Description, row order, Selected Confidence Level or unreferenced Config;
    status returns to `CURRENT` when inputs are restored exactly, with no
    recalculation; there is no change-event handler anywhere;
11. **write-back is transactional**: a mid-write failure restores the previous
    snapshot completely, no mixed state survives, and the success marker is
    written last;
12. the `tblCalcDrivers` column sums reconstruct A, B, C and D exactly, and no
    column carries two meanings by Driver Kind;
13. the full **35/35** Phase-4 functional matrix still passes; all 8 Phase-4
    modules and all 5 Phase-4 buttons persist; Phase-5 modules persist; **no
    Calculate button exists**;
14. the harness asserts every calculated value with no manual inspection;
15. Excel shuts down naturally with clean transient COM release;
16. **Phase 5 introduces no RNG implementation, no sampling implementation and no
    simulation output, and makes no use of Iterations, Random Seed or Selected
    Confidence Level in any analytical calculation or in the fingerprint;
    `_SimData` remains unchanged and unused** (§22).

---

## 29. Decisions

**D1–D6 are locked** (§4), and Revision C leaves them untouched. Every constant
this design depends on — the canonical encoding, the hash parameters, the `_Calc`
anchors, the tolerances and their conditioning scales, the status semantics and
the commit ordering — is now stated as a value rather than as an intention.

**No open decisions remain.**

Model source, workbook artifacts, contracts, bootstrap, harness and Phase-4 tests
are unchanged by this document. No code has been written.

**PHASE 5 PLAN REVISION C READY FOR REVIEW**
