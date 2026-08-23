# Phase 6 — Stochastic simulation layer

**Status: PLANNING ONLY. Proposed for independent review.
NO IMPLEMENTATION EXISTS. NO WINDOWS/EXCEL RUNTIME HAS BEEN EXECUTED.**

Phase 5 is closed. The accepted executable baseline is
`f571154118083e569e1fb9fbf9bf72852cc2d568`; the closure head is `28fa613`. This
document proposes what Phase 6 should be, from the authorities already in the
repository. It changes no code, no contract and no generated artefact.

---

## 1. Authority matrix

Every existing statement found in the sweep that constrains or anticipates
Phase 6. Classification is deliberate: **LOCKED** means an accepted authority
already decides it and Phase 6 may not reopen it; **inherited invariant** means
an accepted design property Phase 6 must preserve; **placeholder only** means a
reserved surface with no committed semantics; **unresolved** means an authority
explicitly deferred the decision to this phase.

### 1.1 What Phase 6 is

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| Phase 6 is the simulation layer | `6+ \| RNG (MRG32k3a), sampling, simulation, percentiles, contingency, sensitivity, Results, Dashboard, Model Check UI \| not started` | `phase5_plan.md` §2 roadmap table | **LOCKED** — but see §2.1, the row is `6+`, not `6` |
| Phase 5 produces no random numbers | "Phase 5 produces numbers. It produces no random numbers." | `phase5_plan.md` §1 | inherited invariant |
| The whole out-of-scope list of Phase 5 | MRG32k3a implementation · seed derivation · RNG stream identity · random variate generation · Bernoulli occurrence · Triangular/Beta-PERT/Uniform sampling · simulation iterations · percentile ladder · P10/P50/P90 · Selected Px · contingency · histogram · CDF · sensitivity/Spearman · `_SimData` iteration storage · Dashboard finalisation · Results finalisation · Model Check UI · annual percentiles · selected-Px annual profiles · simulation reconciliation · any user-facing Calculate button | `phase5_plan.md` §3 | **LOCKED as deferred work** — this is the candidate Phase-6+ scope, not a Phase-6 mandate |
| Phase 6 has not begun | "**PHASE 6 HAS NOT BEGUN.**" | Gate-A steps 1, 2, 3 | obsolete historical note, once this plan is accepted |

### 1.2 The RNG

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| RNG algorithm | `RNG (MRG32k3a)` named in the locked roadmap; listed again as Phase-5 out-of-scope work | `phase5_plan.md` §2, §3 | **inherited invariant — named, never specified.** No parameters, seeding rule, stream discipline or vectors exist anywhere in the repository |
| Seed admissible domain | "The admissible domain is fixed when the RNG is implemented." | `input_contract.yaml` `random_seed.note` | **unresolved — assigned to this phase** |
| Seed range | "Discount Rate and Random Seed have no range. The seed's admissible domain is fixed when the RNG is implemented; inventing one now would be a guess." | `phase2.md` | **unresolved — assigned to this phase** |
| Blank seed | "Optional. Blank means a new random sequence." | `input_contract.yaml` `random_seed.note` | **LOCKED intent**, mechanism unresolved |
| Seed cell | `inpRandomSeed`, Setup `C21`, integer, `required: false`, `default: null`, `validation: null` | `input_contract.yaml` | **LOCKED** |
| `Rnd` / `Randomize` forbidden | `Rnd(`, `Randomize`, `MRG32k3a`, `NPV`, `Percentile`, `RunSimulation` in `forbidden_constructs` | `structure_contract.yaml` | **LOCKED for Phase 4**; Phase 6 must decide how the list is narrowed rather than deleted (open decision **D6-11**) |

### 1.3 The numerical kernel Phase 6 inherits

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| Zero worksheet access per iteration | "at 100,000 iterations the kernel performs ~30 million multiply-adds and **zero** Excel round trips" | `phase5_plan.md` §26 | **LOCKED architectural invariant** |
| Carry types | `DriverFactors`, `YearFactors`, plus one `Double` array per driver of `w × infl`, length `N`. "Simulation needs **only** these: no worksheet, no ListObject, no Range." | `phase5_plan.md` §26; `modCalcFactors.bas` "The two carry types Phase 6 reuses" | **LOCKED** |
| Kernel purity sweep | `modCalcFactors`, `modCalcAnalytical`, `modCalcFingerprint` may not name `Application.`, `ThisWorkbook`, `ActiveWorkbook`, … "the boundary is what makes the kernel reusable by Phase 6 inside an iteration loop" | `phase5_plan.md` §17 | **LOCKED** |
| Probability excluded from factors | "Probability is handled **separately** — multiplied in for analytical EMV, replaced by a Bernoulli draw in Monte Carlo. It is deliberately **not** folded into `Kpv`." | `phase5_plan.md` §10 | **LOCKED — and it is a Phase-6 instruction** |
| Memory headroom | "`DriverFactors` ≈ 26 KB; per-driver weight arrays 300 × 25 × 8 B = 60 KB … **Under 100 KB resident**, which is why the simulation phase can hold it all for an entire run." | `phase5_plan.md` §26 | inherited invariant |
| Fingerprint is worksheet-free by design | "the fingerprint must be computable from resolved data alone, so Phase 6 can extend it without a worksheet" | `phase5_plan.md` §17 | **LOCKED** |

### 1.4 Distribution semantics

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| Accepted families | Triangular, Beta-PERT, Uniform, from `lstDistributions` | `driver_contract.yaml` | **LOCKED** |
| Uniform is two-point | "Uniform is a TWO-POINT distribution. A populated Most Likely is accepted and IGNORED" (D1) | `modCalcCheck.bas`; `phase5_plan.md` D1 | **LOCKED** |
| Ordering rule | Uniform requires `Min ≤ Max`; Triangular and Beta-PERT require `Min ≤ ML ≤ Max` | `modCalcCheck.bas` | **LOCKED** |
| Negative values legal | "NO POSITIVITY RULE IS INVENTED. A correctly ordered set of negative values is a valid distribution — a credit, a saving, a transfer out" | `modCalcCheck.bas`; `phase5_plan.md` §15 | **LOCKED — sampling must not assume a positive support** |
| PERT mean weighting | `(Min + 4·ML + Max)/6`, formed as four copies of ML to avoid overflow | `modCalcAnalytical.PertMean` | **LOCKED** — Phase-6 sampling must be consistent with this mean |
| Central Basis | `ML` for Triangular and Beta-PERT, `Midpoint` for Uniform; applies to both driver kinds | `calc_contract.yaml`; `phase5_plan.md` §5.1, §10; Run-12 accepted | **LOCKED** |
| Cost-line uncertainty is at unit-cost level | `A = central × Quantity × K`; Quantity is deterministic and enters once, outside the factors | `phase5_plan.md` §5.1, §26 | **LOCKED** |

### 1.5 Setup inputs

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| Iterations default | `10000` | `input_contract.yaml` | **LOCKED** |
| Iterations hard minimum | `whole ≥ 1000`; "1000 is the locked hard minimum" | `input_contract.yaml`; `phase2.md` | **LOCKED** |
| Iterations upper limit | "No upper limit is imposed; the <10000 advisory belongs to Model Check." | `input_contract.yaml` | **LOCKED — there is no cap** |
| 100,000 | "Design target: 200 Cost Lines, 100 Risks, 25 project years, 100,000 iterations — targets, not caps" | `phase5_plan.md` §26, §16.2 | **LOCKED as a target, explicitly not a cap** |
| Selected Confidence Level | "Reporting selector only. It does not drive simulation execution." / "wired to nothing — not to simulation execution, not to staleness, not to any calculation." | `input_contract.yaml`; `phase2.md` | **LOCKED as of Phase 2**; whether it stays true is assigned to this phase |
| Confidence list | `P50 P55 P60 P65 P70 P75 P80 P85 P90 P95` — ten values, `lstConfidenceLevels`, not editable | `input_contract.yaml` | **LOCKED — and it does not contain P10** |

### 1.6 Output surfaces

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| `_SimData` | `veryHidden`, role `technical`, "Raw iteration output. Machine data with no audit value in raw form." Currently "intentionally empty" | `workbook.yaml` | **placeholder only** |
| Results | Run Stamp (model version, run id, timestamp, iterations, seed, applied timeline) · Summary Statistics (percentile ladder, mean, standard deviation, dispersion) · Annual Cash Flow · Reconciliation | `workbook.yaml` | **placeholder only** — section titles and notes, no schema |
| Dashboard | "Read-only; every figure originates on Results." Run Status · Headline Figures · Statistical Summary · Charts · Basis and Assumptions | `workbook.yaml` | **placeholder only — but the read-only-from-Results rule is an inherited invariant** |
| Sensitivity | "Driver ranking by Spearman rank correlation against total project cost." Ranked Drivers: driver id, type, name, Spearman rho, absolute rho, rank, direction | `workbook.yaml` | **placeholder only** |
| Methodology | reserved for "distributions, percentile method, contingency definitions, sensitivity method and stated limitations" | `workbook.yaml` | placeholder only |

### 1.7 State, fingerprint and workflow

| Topic | Existing statement | File / section | Class |
|---|---|---|---|
| Section ordering | "`HEADER`, `COST`, `RISK`. **Phase 6 appends its sections after these**; the analytical sections keep their positions so the analytical subset stays comparable across phases." | `phase5_plan.md` §11.4 | **LOCKED — an explicit Phase-6 instruction** |
| Hash algorithm | `FP_BASE 131`, `FP_MOD_1 2147483647`, `FP_MOD_2 2147483629`, `FP_INIT_1/2 = 1`, `FP_VERSION 1` | `phase5_plan.md` §11.4 | **LOCKED** |
| Fingerprint version | `FP_VERSION` is "the fingerprint ALGORITHM VERSION NUMBER only", stored alongside the digest so a workbook calculated under a future version 2 can be recognised | `calc_contract.yaml`; `phase5_plan.md` §11 | **LOCKED mechanism** |
| Phase-5 excluded from the fingerprint | Iterations, Random Seed and Selected Confidence Level are used in no analytical calculation and in no Phase-5 fingerprint | `phase5_plan.md` §16.1 | **LOCKED for Phase 5** — Phase 6 decides its own sections |
| `calc_state` fields | `last_successful_stamp`, `last_successful_fingerprint`, `fingerprint_version`, `last_successful_applied_timeline`, `last_attempt_result`, `last_attempt_detail` | `calc_contract.yaml` | **LOCKED** — the state philosophy Phase 6 should mirror |
| No Calculate button | "LOCKED for Phase 5: `PCCM_Calculate` — yes. Calculate button — no." "A standalone Calculate button was not part of the locked Dashboard command set" | `phase5_plan.md` §17, §3 | **LOCKED for Phase 5**; the eventual command surface is unresolved |
| Correlation between drivers | **No statement exists anywhere.** The only occurrences of "correlation" in the repository are the Sensitivity sheet's *output* Spearman ranking | sweep result | **absent — see §13** |

### 1.8 Contradictions found

**One, and it is a scoping ambiguity rather than a conflict of fact.**

The roadmap row is labelled **`6+`**, and it bundles nine items —
`RNG, sampling, simulation, percentiles, contingency, sensitivity, Results,
Dashboard, Model Check UI`. Read as "Phase 6", it mandates Dashboard, charts,
sensitivity and the Model Check UI in this phase. Read as "Phase 6 and later",
it is a list of everything still outstanding. The `+` and the fact that the
preceding rows are exact single phases (`4`, `5`) say the second reading is
correct, and §11 proposes the split explicitly rather than leaving it to
inference.

**No other contradiction was found.** In particular the repository is silent on
inter-driver correlation, so §13's out-of-scope recommendation contradicts
nothing.

---

## 2. Proposed Phase-6 scope

### 2.1 In scope

The simulation **engine** and its **persisted results**, with the minimum
Results-sheet output needed to make those results auditable:

- MRG32k3a RNG, seeding, stream discipline, and its locked vectors
- inverse-CDF sampling for Triangular, Uniform and Beta-PERT
- Bernoulli risk occurrence
- the per-iteration total accumulation, nominal and PV
- percentile / moment statistics over the retained totals
- contingency at the selected confidence level
- the simulation fingerprint sections, appended after `HEADER`/`COST`/`RISK`
- simulation run state: CURRENT / STALE / INVALID, effective seed, run identity
- `_SimData` as the persisted results surface
- a minimal Results block: Run Stamp and Summary Statistics
- `PCCM_RunSimulation` and its read accessors — **no button**

### 2.2 Out of scope, with the proposed later-phase boundary

| Deferred to | Item | Why |
|---|---|---|
| **Phase 7** | Sensitivity / Spearman ranking, tornado | Needs per-driver per-iteration information that the engine does not otherwise retain (§12). Correctness of the distribution should be accepted before anything is ranked against it |
| **Phase 7** | Annual simulated distributions, selected-Px annual profiles | A second, larger retention decision; the analytical annual series already exists and is accepted |
| **Phase 8** | Results finalisation, Annual Cash Flow and Reconciliation blocks | Presentation over an accepted engine |
| **Phase 8** | Dashboard, charts, S-curve, histogram | `workbook.yaml` already requires Dashboard to read from Results; it cannot be built before Results is final |
| **Phase 9** | Model Check UI, warning aggregation, the `<10000` iterations advisory | Named as Model Check's job by `input_contract.yaml` |
| **not scheduled** | Correlated uncertainty between drivers | §13 |

---

## 3. Execution model

### 3.1 The invariant, restated

> Resolve Excel inputs once. Run the simulation kernel over plain in-memory
> numerical structures. **No `Range`, `ListObject` or worksheet access inside the
> iteration loop.**

### 3.2 Reused unchanged from Phase 5

`DriverFactors` (all thirteen fields), `YearFactors`, the per-driver
`w × infl` vectors, `modCalcResolve` for worksheet→structure resolution,
`modCalcCheck` for numerical prerequisites, and every safe-arithmetic primitive
in `modCalcFactors`. **Phase 6 adds no field to either carry type.**

### 3.3 New structures Phase 6 requires

```vb
Type RngState                 ' MRG32k3a: six Doubles holding integers
    S10 As Double: S11 As Double: S12 As Double
    S20 As Double: S21 As Double: S22 As Double
End Type

Type SimulationRequest        ' resolved once, before the loop
    Iterations       As Long
    EffectiveSeed    As Double
    SeedWasSupplied  As Boolean
    RngVersion       As Long
    MethodVersion    As Long
End Type

Type SimulationResult         ' produced once, after the loop
    Iterations       As Long
    EffectiveSeed    As Double
    MeanNominal      As Double:  MeanPv      As Double
    StdDevNominal    As Double:  StdDevPv    As Double
    MinNominal       As Double:  MaxNominal  As Double
    MinPv            As Double:  MaxPv       As Double
    ' percentile ladder, nominal and PV, in the locked order of §8.3
End Type
```

plus two `Double` arrays of length `Iterations` — the retained nominal totals and
PV totals — which are what the percentile step sorts.

### 3.4 Phase separation

| When | Work |
|---|---|
| **once, before the loop** | resolve inputs; run the Phase-5 checks; build `DriverFactors`, `YearFactors` and the weight vectors; read Iterations and Random Seed; derive the effective seed; seed the RNG; allocate the two sample arrays |
| **once per driver, before the loop** | nothing new — `Knom`, `Kpv`, `Quantity`, `Probability`, `DistKind`, `Min/ML/Max` are already resolved by Phase 5 |
| **per iteration, per driver** | one uniform for a cost line (unit cost), or one uniform for Bernoulli plus one conditional uniform for severity for a risk; one inverse-CDF evaluation; two multiply-accumulates (nominal and PV) |
| **per iteration** | write the two accumulated totals into the sample arrays at index `i` |
| **after the loop** | sort each sample array; compute moments and percentiles; compute contingency; compute the simulation fingerprint; write `_SimData` and Results in one transactional publish |

Per-iteration cost at the design target: **300 inverse-CDF draws and 600
multiply-accumulates**, no allocation, no COM.

---

## 4. Sampling semantics

All sampling is **inverse-CDF on a single uniform** `u ∈ (0,1)`, which is what
makes the oracle exactly reproducible and makes stream consumption countable.
Rejection sampling is rejected for exactly that reason.

Let `a = Min`, `m = MostLikely`, `b = Max`, with the Phase-5 ordering guarantees
already established. **No positivity is assumed**; every formula below is valid
for negative `a`, `m`, `b`.

### 4.1 Uniform

```
x = a + u · (b − a)
```

`m` is not read (D1). Degenerate `a = b` returns `a`.

### 4.2 Triangular — inverse CDF

With `c = (m − a) / (b − a)`:

```
u ≤ c :  x = a + sqrt( u · (b − a) · (m − a) )
u > c :  x = b − sqrt( (1 − u) · (b − a) · (b − m) )
```

Boundary cases, all stated rather than left to the implementation:

| Case | Result |
|---|---|
| `a = b` (degenerate point) | `x = a`, no `u` consumed? **No — one `u` is consumed anyway.** Stream consumption must not depend on data values (§5.6) |
| `m = a` (right triangle) | `c = 0`; the `u > c` branch always taken |
| `m = b` (left triangle) | `c = 1`; the `u ≤ c` branch always taken |
| `u = 0` | `x = a` |
| `u → 1` | `x → b`; `u` is drawn from the open interval so `x = b` is not produced |

### 4.3 Beta-PERT

Beta-PERT has no closed-form inverse CDF, and its mean is **locked** at
`(a + 4m + b)/6` by `PertMean`. The sampler must be consistent with that mean.
Standard Beta-PERT sets `λ = 4` and

```
α = 1 + λ·(m − a)/(b − a)      β = 1 + λ·(b − m)/(b − a)
x = a + (b − a) · BetaInvCDF(u; α, β)
```

whose mean is exactly `(a + 4m + b)/6`. `BetaInvCDF` requires a numerical
inversion — the incomplete beta function inverse. This is the single hardest
numerical component of Phase 6 and it is called **open decision D6-04**: the
inversion algorithm, its iteration count, its convergence tolerance and its
behaviour at `u → 0` and `u → 1` must all be locked in the plan before code,
because VBA and Python must agree bit-for-bit or the oracle is worthless.

### 4.4 Cost Line

```
unitCost = sample(dist, a, m, b)                 ' ONE draw, at unit-cost level
contribNom = unitCost × Quantity × Knom
contribPv  = unitCost × Quantity × Kpv
```

Quantity is **deterministic** and applied after sampling — the accepted
semantics, unchanged.

### 4.5 Risk

```
occurred = ( u_bernoulli < Probability )         ' ONE draw, always consumed
if occurred:
    severity   = sample(dist, a, m, b)           ' ONE draw, always consumed
    contribNom = severity × Knom
    contribPv  = severity × Kpv
else:
    contribNom = 0 ;  contribPv = 0
```

`Probability` is **not** folded into `Knom`/`Kpv` — the locked Phase-5 rule.

**Strict `<`**, so `Probability = 0` never occurs and `Probability = 1` always
occurs, given `u ∈ (0,1)`.

**The severity draw is consumed whether or not the risk occurred.** This is a
deliberate design commitment: it keeps stream position a function of the model
shape alone, never of the sampled values, which is what makes a run reproducible
and a per-driver substream possible.

---

## 5. RNG — the locked contract

### 5.1 Algorithm

**MRG32k3a** (L'Ecuyer 1999), as named in the locked roadmap. Evaluated against
the criteria the review required:

| Criterion | MRG32k3a | VBA `Rnd` |
|---|---|---|
| Reproducible in VBA | Yes — all arithmetic in `Double`, every intermediate `< 2^53` | Undocumented, version-dependent, not portable |
| Arithmetic safety | Products bounded by `~2^53`; exact in IEEE-754 double | n/a |
| Cross-language oracle | Trivially exact in Python integers | Not implementable |
| Period | `≈ 2^191` | `2^24` — **exhausted well before 100,000 × 300 draws** |
| Speed | ~10 flops + 2 mods per uniform | faster, irrelevant |
| Exact vector verification | Yes, integer state | No |

`Rnd` is not merely inadvisable; its period is **too short for this model's
design target** by four orders of magnitude.

### 5.2 State and constants — proposed LOCKED

```
m1 = 4294967087        a12 =  1403580     a13n =  810728
m2 = 4294944443        a21 =   527612     a23n = 1370589
norm = 2.328306549295727688e-10          ' 1 / (m1 + 1)
```

State is six `Double`s holding integers: `s1 = (s10,s11,s12)` in `[0, m1)`,
`s2 = (s20,s21,s22)` in `[0, m2)`. Not all of `s1` may be zero; not all of `s2`
may be zero.

```
p1 = (a12·s11 − a13n·s10) mod m1        ; s10←s11, s11←s12, s12←p1
p2 = (a21·s22 − a23n·s20) mod m2        ; s20←s21, s21←s22, s22←p2
u  = ((p1 − p2) mod m1 + 1) · norm       ' u ∈ (0, 1], see D6-02
```

**The `+1` puts `u` in `(0, 1]`, not `[0, 1)`.** That matters at the Triangular
and Uniform boundaries, and it is **open decision D6-02**: whether to adopt the
canonical `(0,1]` form or to reflect to `[0,1)`. Recommendation: keep the
canonical form and define every sampler on `(0,1]`.

### 5.3 Seed domain — proposed LOCKED

`inpRandomSeed` is declared `integer`, optional, no validation. Proposed:

- admissible domain **`1 … 2147483646`** (whole numbers)
- `0`, negatives and non-integers are **refused**, not coerced
- blank is legal and means "choose one" (§5.4)

Rationale: a single positive 31-bit integer is what a user can type and record,
it maps cleanly onto the seeding function below, and it excludes the all-zero
state by construction.

### 5.4 Blank seed

A blank seed produces an **effective seed** derived at run time. The derived
value is **published on Results as part of the Run Stamp and stored in the
simulation state**, so a run started with a blank seed is still exactly
reproducible afterwards by typing the published value back into `C21`.

Source of the derived seed is **open decision D6-03**. Recommendation: a
timestamp-derived value folded into the admissible domain, chosen for
auditability rather than entropy quality — this is not a cryptographic context.

### 5.5 Seeding function

A user seed of `1` must not produce a near-zero state. Proposed: expand the
single integer through a small locked splitting function into the six state
words, each reduced into its modulus and forced non-zero. The exact function,
and its first-N vectors, are Gate-A material and are **open decision D6-05**.

### 5.6 Stream discipline — proposed LOCKED

**One global stream, consumed in a fixed order**, and the order is a function of
the model's *shape*, never of its sampled values:

1. drivers are visited in **ascending Permanent ID**, ordinal UTF-16 comparison —
   the same canonical order the Phase-5 fingerprint already locks
2. each Cost Line consumes exactly **one** uniform
3. each Risk consumes exactly **two** uniforms — Bernoulli then severity —
   **always**, whether or not it occurred
4. one iteration therefore consumes exactly `C + 2R` uniforms, a constant

This yields **row-order invariance for free**: reordering rows in the register
changes no Permanent ID, so it changes no draw. It is a testable property, not a
hope, and it is a Gate-B requirement.

Independent substreams per driver were considered and are **not** proposed: they
would add a jump-ahead implementation and a second class of vectors for a
property the fixed consumption order already delivers. Revisit only if Phase 7's
sensitivity work needs it (§12).

### 5.7 Reproducibility guarantee — proposed wording

> The same resolved model, the same effective seed, the same iteration count and
> the same RNG version produce bit-identical nominal and PV sample arrays, and
> therefore bit-identical statistics, on any machine.

"Same resolved model" is precisely what the simulation fingerprint captures
(§10), which is what makes the guarantee checkable rather than aspirational.

### 5.8 Versioning

`RNG_VERSION` and `SIM_METHOD_VERSION` are separate integers, both stored with
the run. The first changes if the generator or seeding changes; the second if a
sampler or an accumulation order changes. Either invalidates a stored result the
way `FP_VERSION` already does.

---

## 6. Iterations

| Concept | Value | Source |
|---|---|---|
| Hard refusal | `< 1000`, or non-whole | `input_contract.yaml`, LOCKED |
| Default | `10000` | LOCKED |
| Advisory | `< 10000` — belongs to **Model Check**, not to Phase 6 | `input_contract.yaml`, LOCKED |
| Tested performance target | `100000` | `phase5_plan.md` §26, target not cap |
| Technical ceiling | see below | proposed |

**No upper business limit is introduced.** A technical ceiling does exist and
must be stated rather than discovered: two `Double` arrays of length `n` cost
`16n` bytes, so 10 million iterations is 160 MB and 100 million is 1.6 GB. The
proposal is to **refuse allocation failure gracefully** with a numerical-range
message naming iterations, rather than to invent a cap. Exact ceiling behaviour
is **open decision D6-08**.

---

## 7. Selected Confidence Level

**Recommendation: the Phase-2 statement stops being true in one respect only.**

- The simulation distribution is generated **independently** of the selected
  confidence level. Changing `P50` to `P80` re-selects from an existing sorted
  sample; it does not re-run anything and does not change one random number.
- Therefore Selected Confidence Level is **not** part of the simulation
  computational fingerprint (§10.2).
- It **is** part of a separate, smaller notion: the *reporting selection*. A
  workbook whose stored results were selected at `P50` and whose cell now reads
  `P80` is not stale — it is *unselected*, and the correct response is to
  re-derive the selected figures from the stored sample, not to re-simulate.

This preserves the Phase-2 intent ("reporting selector only") while making
honest what changes: it now selects from a real distribution instead of nothing.

**Note a gap in the accepted list.** `lstConfidenceLevels` is
`P50 … P95` and **contains no P10**. The Dashboard placeholder nevertheless asks
for "P10 / P50 / selected Px / P90". These are compatible only if P10 is a
*fixed reported statistic* rather than a *selectable* one. That is the proposal,
and it is **open decision D6-06**.

---

## 8. Statistics

### 8.1 Moments

- **Sample mean**: `Σx / n`.
- **Standard deviation**: the **sample** standard deviation, divisor `n − 1`.
  Rationale: the run is a sample from the model's distribution, not the
  population. Stated explicitly because "Excel's `STDEV` vs `STDEVP`" is exactly
  the ambiguity this project refuses elsewhere.
- Computed with the accepted Phase-5 safe-arithmetic primitives, in a numerically
  stable pass — **not** the naive `Σx²` form.

### 8.2 Percentile algorithm — proposed LOCKED

**Nearest-rank on the sorted ascending sample, with linear interpolation**, the
same definition as `PERCENTILE.INC` / NumPy's `linear` method, stated in full so
no library default is inherited:

```
sorted x[0..n-1] ascending
h = (n − 1) · p                       ' p ∈ [0,1], e.g. P70 → p = 0.70
lo = floor(h) ; hi = min(lo + 1, n − 1)
Px = x[lo] + (h − lo) · (x[hi] − x[lo])
```

`n = 1` returns `x[0]`. `p = 0` returns the minimum; `p = 1` the maximum.

**No `WorksheetFunction.Percentile`** — it is already in the forbidden-construct
list, and a worksheet function inside the kernel would break the purity sweep.

### 8.3 The reported ladder — proposed LOCKED order

`P10, P50, P70, P90, selected Px`, each reported for **nominal** and **PV**
separately, plus mean, standard deviation, minimum and maximum for each.

P10, P50, P70 and P90 are **always** reported. Selected Px is whichever of the
ten `lstConfidenceLevels` values `C18` holds.

### 8.4 Deferred

**Annual simulated distributions are deferred to Phase 7**, explicitly. Phase 6
reports the analytical annual series that Phase 5 already produces, and does not
retain per-year samples.

### 8.5 Hand-derived vectors before code

Gate A must carry hand-derived percentile vectors for `n = 1, 2, 3, 4, 10` at
every reported `p`, including the cases where `h` is exactly integral and where
it falls between two samples.

---

## 9. Contingency

Six quantities, each with its own name, and the word "contingency" is never used
for any of the first five:

| # | Quantity | Definition | Origin |
|---|---|---|---|
| 1 | Escalated Deterministic Base Estimate | `A` | Phase 5, accepted |
| 2 | Expected Risk / EMV | `D` | Phase 5, accepted |
| 3 | Analytical Expected Total | `C + D` | Phase 5, accepted |
| 4 | Simulation Mean Total | `Σx/n` over the sample | Phase 6 |
| 5 | Selected Px Total | `Px` from §8.2 | Phase 6 |
| 6 | **Contingency at the selected confidence level** | **`Selected Px Total − Escalated Deterministic Base Estimate`** | Phase 6 |

The baseline for contingency is the **deterministic base estimate (1)**, not the
simulation mean and not the analytical expected total. Proposed and stated
explicitly because all three are defensible and only one can be the definition.
Reported for nominal and PV separately.

**The workbook reports at the confidence level the user selected. It does not
recommend one.** No automatic recommendation, no "suggested" level, no
highlighting of a preferred row.

---

## 10. Simulation fingerprint and state

### 10.1 Extension, not replacement

Sections are appended **after** `HEADER`, `COST`, `RISK` — the locked ordering —
so the analytical subset remains byte-comparable across phases. Proposed new
section: **`SIM`**, appended fourth.

### 10.2 What `SIM` contains

| Field | In? | Why |
|---|---|---|
| Iteration count | **yes** | changes the distribution |
| Effective seed | **yes** | changes the distribution |
| `RNG_VERSION` | **yes** | changes the numbers for identical inputs |
| `SIM_METHOD_VERSION` | **yes** | same |
| Selected Confidence Level | **no** | §7 — it selects from the distribution, it does not generate it |

### 10.3 The three states

Mirroring the accepted Phase-5 attempt/state philosophy rather than inventing a
second machine:

| State | Meaning |
|---|---|
| **CURRENT** | stored analytical fingerprint matches, stored `SIM` fingerprint matches, and the last simulation attempt succeeded |
| **STALE** | a stored successful result exists, but the analytical or `SIM` fingerprint no longer matches |
| **INVALID** | prerequisites refuse, or the last attempt failed |

Plus one Phase-6 addition, **UNSELECTED**: results are CURRENT but `C18` names a
confidence level the stored selection was not derived at. Resolved by
re-selecting from the stored sample — never by re-simulating.

### 10.4 What survives a refusal

The accepted Phase-5 rule, unchanged: a refused simulation **preserves the prior
successful results** and leaves them visible, marked with their own stamp, and
records the failure in the attempt fields. A failed run never publishes a
partial distribution (§16).

### 10.5 New state fields

`last_successful_sim_stamp`, `last_successful_sim_fingerprint`,
`rng_version`, `sim_method_version`, `effective_seed`, `iterations_run`,
`last_sim_attempt_result`, `last_sim_attempt_detail` — the same shape as
`calc_state`, in a `sim_state` block.

---

## 11. Results ownership

| Layer | Phase | Note |
|---|---|---|
| **A. simulation engine** | **6** | |
| **B. persisted results** | **6** | `_SimData` holds the sorted sample arrays and the run stamp |
| **C. Results presentation** | **6, minimal** — Run Stamp + Summary Statistics only | Annual Cash Flow and Reconciliation blocks deferred to Phase 8 |
| D. Dashboard | 8 | must read from Results, never recompute |
| E. Sensitivity / Spearman | 7 | §12 |
| F. Charts | 8 | |

**No duplicate source of truth.** Results derives from `_SimData`; Dashboard
derives from Results; nothing recomputes a Monte Carlo independently. That rule
is already in `workbook.yaml` for Dashboard and is extended here to Results.

---

## 12. Sensitivity boundary — recommended Phase 7

Spearman rank correlation of each driver's sampled contribution against the
total requires **per-driver, per-iteration** values.

- Retaining them at the design target: `300 × 100,000 × 8 B = 240 MB`.
  Against the current `< 100 KB` resident kernel, that is a different program.
- **Online computation is not available for Spearman.** Ranking is inherently a
  whole-sample operation; a rank cannot be accumulated incrementally the way a
  mean can. Pearson could be online; Spearman cannot.
- Options are therefore: retain 240 MB, re-run per driver, or store to
  `_SimData` and rank in a second pass.

**Recommendation: defer to Phase 7**, and get the distribution accepted first.

**But one Phase-6 decision is forced by it.** If Phase 7 must reproduce the exact
per-driver draws of a Phase-6 run, the RNG design must make that possible. The
fixed consumption order of §5.6 does: driver `k`'s draw in iteration `i` is at a
computable stream offset, so Phase 7 can replay it exactly without Phase 6
storing anything. **This is the reason to lock §5.6 now**, and it is why
independent substreams are not needed.

---

## 13. Correlation between drivers — OUT OF SCOPE

The repository contains **no** inter-driver correlation authority. The only
occurrences of the word are the Sensitivity sheet's *output* ranking, which is a
different thing entirely — measuring how a driver co-varies with the total is not
imposing dependence between inputs.

**Proposed: drivers are sampled independently, and correlated uncertainty is
explicitly out of scope and marked as a stated limitation on the Methodology
sheet.** No correlation matrix, no copula, no hidden dependency. Introducing one
would require a new input contract and a new authority, neither of which exists.

No conflicting repository text was found.

---

## 14. Performance model

At the design target — 200 Cost Lines, 100 Risks, 25 years, 100,000 iterations:

| Quantity | Count |
|---|---|
| Uniforms per iteration | `200 + 2×100 = 400` |
| Uniforms per run | `4.0 × 10^7` |
| Inverse-CDF evaluations per run | `3.0 × 10^7` (200 cost + 100 severity) |
| Bernoulli comparisons per run | `1.0 × 10^7` |
| Multiply-accumulates per run | `6.0 × 10^7` (nominal + PV) |
| Sorts | 2, of `10^5` elements — `~2 × 1.7 × 10^6` comparisons |
| **Worksheet / COM calls inside the loop** | **0** |

Memory:

| Item | Bytes |
|---|---|
| Phase-5 resident kernel | `< 100 KB` |
| Retained totals, nominal + PV | `2 × 100,000 × 8 = 1.6 MB` |
| If annual samples were retained (deferred) | `25 × 100,000 × 8 = 20 MB` |
| If driver samples were retained (deferred) | `300 × 100,000 × 8 = 240 MB` |
| **Phase-6 total** | **`< 2 MB`** |

The Beta-PERT inverse (§4.3) is the dominant per-draw cost and is the one place
where the iteration budget could be missed. Its cost must be measured in Gate A
against the Python oracle before the VBA is written.

---

## 15. Oracle and cross-language evidence

A pure-Python oracle, independent of the VBA, extending
`builder/pccm_builder/` alongside the accepted `calc_oracle.py`.

**Independence strategy** — the accepted Phase-5 approach, which is not a
line-for-line port:

- the oracle is written **from this document's equations**, in Python's own
  idiom, using exact integer arithmetic for the RNG where VBA must use `Double`
- agreement between two implementations that made *different* representational
  choices is evidence; agreement between a port and its original is not
- the RNG is the sharpest case: Python computes the recurrence in unbounded
  integers, VBA in `Double`. If VBA's `Double` arithmetic ever exceeded `2^53`
  the two would diverge — which is exactly the property worth testing

Locked vector families for Gate A:

| # | Vectors |
|---|---|
| 1 | RNG state after seeding, for seeds `1`, `2`, `12345`, `2147483646` |
| 2 | first 20 uniforms for each of those seeds, to full 17 significant digits |
| 3 | seed edge cases: `1`, `2147483646`, and the refusals `0`, `−1`, `1.5`, blank |
| 4 | Bernoulli boundaries: `p = 0`, `p = 1`, `p = 0.5` against a known `u` |
| 5 | Triangular inverse at `u` = the branch point `c`, and either side |
| 6 | Uniform inverse, including `a = b` and negative supports |
| 7 | Beta-PERT inverse at `u ∈ {0.01, 0.25, 0.5, 0.75, 0.99}`, plus `m = a` and `m = b` |
| 8 | degenerate shapes: `a = m = b` |
| 9 | one-driver simulation, 1000 iterations, full sample digest |
| 10 | mixed 3 cost + 2 risk simulation, full sample digest |
| 11 | same seed → identical digest |
| 12 | changed seed → different digest |
| 13 | **row reorder → identical digest** |
| 14 | percentile vectors for `n = 1, 2, 3, 4, 10` |
| 15 | nominal and PV totals against a hand-derived case |

---

## 16. Error and interruption behaviour

| Situation | Behaviour |
|---|---|
| Invalid input before simulation | refuse before allocating or drawing; prior results untouched; `INVALID` |
| Numerical failure mid-run | abort; publish nothing; prior results untouched; detail names the iteration and driver |
| Unexpected VBA error | the accepted `modAppState` refusal path; no partial publish |
| User cancellation | **not supported in Phase 6** — a 100,000-iteration run at the target completes in seconds; a cancel path is a second exit route through the transactional publish and is not worth the risk yet |
| Excel state | restored exactly as the accepted Phase-4/5 lifecycle does |
| Partial results | structurally impossible: `_SimData` and Results are written **once, after** the loop, in one transactional publish, commit-last |
| Effective seed on failure | recorded in the attempt detail, so a failing run is reproducible |

> A failed run never publishes a partially completed distribution as a
> successful result.

---

## 17. User-facing command surface

Phase 5 deliberately added no button, and that stands.

| Control | Phase | Why |
|---|---|---|
| `PCCM_RunSimulation` (`Application.Run`) | **6** | the automation entry point, testable by the Gate-B harness |
| Model Check button | 9 | Model Check owns it |
| Run Simulation button | 8 | belongs with the Dashboard command set, which is where the locked architecture puts commands |
| Refresh / report outputs | 8 | same |

Phase 6 adds **no button, no ribbon, no form, no `MsgBox`**.

---

## 18. Gate strategy

**Gate A — pure and static, on Linux.** RNG and sampler contracts in a new
`spec/sim_contract.yaml`; the Python oracle; every vector family of §15;
static VBA source checks including the extended purity sweep over the new
simulation modules; generated artefacts and Stage-A verification; mutation
controls on every new detector. **No Windows runtime claim of any kind.**

**Gate B — real Windows/Excel.** Whole-project compile through the accepted
P5-CMP gate; seeded simulation against the Python expectations; repeatability
(same seed twice, identical digest); different-seed divergence; **row-order
invariance**; the full 100,000-iteration run at the design target; failure
containment; natural Excel shutdown; clean COM release.

**A performance benchmark is not a semantic test.** The 100k run proves the
budget is met; it proves nothing about correctness, and Gate B passes only if the
semantic scenarios pass independently of it.

---

## 19. Implementation sequence

Each step is independently reviewable, and each states what remains
unimplemented.

| # | Step | Files allowed to change | New authority | Gate-A acceptance |
|---|---|---|---|---|
| 1 | **`spec/sim_contract.yaml`** — RNG constants, seed domain, sampler definitions, percentile method, contingency definition, `SIM` fingerprint fields, `sim_state` schema, `RNG_VERSION`, `SIM_METHOD_VERSION`. Loader + validator, fail-loud. **No hash constants, no vectors.** | `spec/`, `builder/`, `tests/` | the sixth contract | validator rejects every malformed shape; mutation controls on each rule |
| 2 | **RNG reference implementation and vectors** — Python MRG32k3a in exact integers, the seeding function, vector families 1–3. Written before any consumer. | `builder/`, `tests/` | the locked vectors | vectors reproduce L'Ecuyer's published stream; seeding is deterministic |
| 3 | **Samplers in Python** — Triangular, Uniform, Beta-PERT inverse, Bernoulli; vector families 4–8; the Beta inversion algorithm locked and its cost measured | `builder/`, `tests/` | — | every boundary case in §4 has a hand-derived expectation |
| 4 | **Simulation oracle** — the whole engine in Python; vector families 9–15; percentile and contingency implementations | `builder/`, `tests/` | — | golden cases written first, from hand derivations |
| 5 | **Stage-A emission** — `build/phase6_cases.json`, `modSimContract.bas` generated; post-build verification extended | `builder/`, `tests/` | — | Stage-A verifier green; emitted corpus complete |
| 6 | **`modSimRng`, `modSimSample`** — the pure VBA kernel, purity sweep active from the first commit, vectors passing against Python | `src/vba/`, `tests/` | — | source suites green; no worksheet name in either module |
| 7 | **`modSimEngine`** — the iteration loop, accumulation, sample retention | `src/vba/`, `tests/` | — | zero worksheet references in the loop, proved statically |
| 8 | **`modSimStats`** — sort, moments, percentiles, contingency | `src/vba/`, `tests/` | — | percentile vectors pass |
| 9 | **`modSimFingerprint`** — the `SIM` section appended after `RISK`; analytical subset proved byte-identical to Phase 5 | `src/vba/`, `tests/` | — | Phase-5 fingerprints unchanged for every accepted case |
| 10 | **`modSimReport`** — `_SimData` and the minimal Results block; transactional publish; `sim_state`; `PCCM_RunSimulation` and accessors | `src/vba/`, `tests/` | — | no partial-publish path exists |
| 11 | **Gate-A source review** | — | — | independent review |
| 12 | **Gate-B harness extension**, then Windows Run 1 | `bootstrap/`, `tests/` | — | §18 |

---

## 20. Open decisions

Only genuinely unresolved items. Anything already locked by an accepted
authority is in §1, not here.

| # | Decision | Options | Existing evidence | Recommended | Consequence | Needed before |
|---|---|---|---|---|---|---|
| **D6-01** | Is the roadmap's `6+` row one phase or many? | (a) all nine items in Phase 6; (b) engine + results in 6, rest in 7–9 | roadmap row says `6+`; rows 4 and 5 are exact | **(b)**, per §2.2 | Sets the whole scope | Step 1 |
| **D6-02** | Uniform interval | (a) `(0,1]` canonical MRG32k3a; (b) reflect to `[0,1)` | none | **(a)** | Boundary behaviour of every sampler | Step 2 |
| **D6-03** | Blank-seed derivation | (a) timestamp-derived; (b) refuse blank; (c) fixed default seed | contract says blank means "a new random sequence" — so (b) contradicts it | **(a)**, published in the Run Stamp | Reproducibility of unseeded runs | Step 1 |
| **D6-04** | Beta inverse-CDF algorithm | (a) Newton on the regularised incomplete beta; (b) bisection to a fixed tolerance; (c) a fixed-iteration continued fraction | none | **(b) or (c)** — whichever is bit-reproducible across VBA and Python | Determines whether the oracle can be exact | Step 3 |
| **D6-05** | Seed-expansion function | (a) splitmix-style; (b) LCG warm-up; (c) fixed offsets per word | none | (a), with locked vectors | Two seeds must not collide | Step 2 |
| **D6-06** | P10 reporting vs `lstConfidenceLevels` | (a) P10 is a fixed reported statistic; (b) add P10 to the selectable list | list is `P50…P95`; Dashboard note asks for P10 | **(a)** — no accepted-list change | Whether an accepted contract is reopened | Step 1 |
| **D6-07** | Contingency baseline | (a) deterministic base `A`; (b) analytical expected total `C+D`; (c) simulation mean | none | **(a)** | The headline number | Step 4 |
| **D6-08** | Iteration technical ceiling | (a) no cap, refuse on allocation failure; (b) a stated representable maximum | "No upper limit is imposed" | **(a)** | Whether Phase 6 invents a cap the contract forbids | Step 7 |
| **D6-09** | Standard deviation divisor | (a) sample `n−1`; (b) population `n` | none | **(a)** | Reported dispersion | Step 4 |
| **D6-10** | Does `UNSELECTED` exist as a state? | (a) yes, per §10.3; (b) treat a changed Px as STALE | Phase-2 "reporting selector only" | **(a)** | Whether changing Px forces a re-run | Step 1 |
| **D6-11** | `forbidden_constructs` narrowing | (a) scope the list per module; (b) delete `MRG32k3a`/`RunSimulation` entries; (c) keep and exempt the new modules | list is Phase-4-scoped by its own comment | **(a)** | A Phase-4 guard must not be weakened to let Phase 6 in | Step 6 |
| **D6-12** | User cancellation | (a) unsupported in Phase 6; (b) supported | none | **(a)**, per §16 | A second exit path through the publish | Step 10 |

---

## 21. What this plan does NOT claim

1. **No Phase-6 code exists.** No VBA, no builder change, no contract, no test.
2. **No Windows or Excel runtime has been executed for Phase 6.**
3. **Nothing here is accepted.** Every recommendation in §20 is a proposal
   awaiting independent review.
4. **The Phase-5 baseline is untouched.** `f571154` remains the accepted
   executable baseline and this document changes no file under `src/`,
   `spec/`, `builder/`, `bootstrap/` or `tests/`.
