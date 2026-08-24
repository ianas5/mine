# Phase 6 — Stochastic simulation layer

**Status: PLANNING ONLY — revision 2, after review round 2.
NOT ACCEPTED. NO IMPLEMENTATION EXISTS. NO WINDOWS/EXCEL RUNTIME HAS BEEN
EXECUTED. NO SIMULATION CONTRACT HAS BEEN COMMITTED.**

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
PV totals, **held in canonical iteration order**. The percentile step sorts
*copies* (§5 of the persistence rules, §11.1); the originals are never permuted,
because their index *is* the iteration identity.

### 3.4 Phase separation

| When | Work |
|---|---|
| **once, before the loop** | resolve inputs; run the Phase-5 checks; build `DriverFactors`, `YearFactors` and the weight vectors; read Iterations and Random Seed; derive the effective seed; seed the RNG; allocate the two sample arrays |
| **once per driver, before the loop** | nothing new — `Knom`, `Kpv`, `Quantity`, `Probability`, `DistKind`, `Min/ML/Max` are already resolved by Phase 5 |
| **per iteration, per driver** | one uniform for a cost line (unit cost), or one uniform for Bernoulli plus one conditional uniform for severity for a risk; one inverse-CDF evaluation; two multiply-accumulates (nominal and PV) |
| **per iteration** | write the two accumulated totals into the sample arrays at index `i` |
| **after the loop** | copy each sample array and sort the copies; compute moments and percentiles; compute contingency; compute the simulation fingerprint; write `_SimData` in iteration order and Results in one transactional publish |

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

### 4.3 Beta-PERT — and why it decides the stream architecture

Beta-PERT has no closed-form inverse CDF, and its mean is **locked** at
`(a + 4m + b)/6` by the accepted `PertMean`. With `λ = 4`:

```
r = (m − a)/(b − a)                    ' shape ratio, in [0,1]
α = 1 + 4r                             ' both in [1,5]
β = 1 + 4(1 − r)                       ' both >= 1
x = a + (b − a) · BetaSample(α, β)
```

whose mean is exactly `(a + 4m + b)/6`, consistent with `PertMean`.

**Revision 1 left this open while already locking one uniform per sample, a
single global stream and fixed `C + 2R` consumption. Those are not independent
decisions, and measuring the options showed the coupling is decisive.**

The parameter family is narrow and benign — `α, β ∈ [1, 5]`, both `≥ 1`, so the
density is bounded and unimodal, peak density `≤ 5.0` over the whole family.
Three architectures were costed at the design target, where the worst case is
200 Cost Lines and 100 Risks all Beta-PERT with `Probability = 1`:
**3 × 10^7 Beta samples per run**.

| | Architecture | Uniforms/sample | Hot-path cost | Exact? | Stream |
|---|---|---|---|---|---|
| **A** | Numerical inverse CDF per sample (bisection on the regularised incomplete beta) | 1 | **2.7 × 10^12 flops** | yes | global, fixed |
| **C** | Per-driver precomputed inverse-CDF table + interpolation | 1 | 3 × 10^8 flops + 5.6 × 10^8 build | **no** | global, fixed |
| **B** | Exact variable-draw generator (Cheng BB/BC, `α,β ≥ 1`) | ~2.2–2.6 | ~3 × 10^9 flops | yes | **substreams required** |

**Option A is infeasible.** The incomplete-beta continued fraction needs up to
**91 iterations** over this parameter family (measured, not estimated), ≈ 1,820
flops per evaluation; a bisection to `1e-15` needs ~50 of them. That is ~91,000
flops per sample and `2.7 × 10^12` per run — hours of VBA, not seconds. Option A
is rejected on arithmetic, not on taste.

**Option C is inaccurate at any table size worth building.** Measured maximum
interpolation error in normalised `[0,1]` units, over PERT shapes
`r ∈ {0.01, 0.25, 0.5, 0.75, 0.99}`:

| Nodes | Max error |
|---|---|
| 256 | `1.6 × 10^-1` |
| 1024 | `1.2 × 10^-1` |
| 4096 | `5.9 × 10^-3` |

The tails are the problem: at `r = 0.01`, `α = 1.04`, and the density behaves
like `x^0.04` near zero, so the inverse CDF has near-infinite slope there and a
uniform node grid cannot follow it. `5.9 × 10^-3` of the support is not a
tolerance this project could defend. Non-uniform nodes or a Newton polish could
rescue it, but both reintroduce cost and a second approximation to justify.

**Option B is therefore the recommendation** — an exact variable-draw generator,
which forces **deterministic per-driver substreams** rather than one global
stream, because consumption is no longer constant.

That is a real cost and it is stated plainly: MRG32k3a was *designed* for it
(L'Ecuyer's RngStreams provides jump-ahead of `2^127` between streams and `2^76`
between substreams, via published 3×3 matrix constants), so the mathematics is
settled and the work is implementation, not invention. Row-order invariance is
then delivered by **stream assignment keyed on Permanent ID** instead of by fixed
consumption order — a different mechanism for the same guarantee, and one that
also gives Phase 7 exact per-driver replay directly.

**D6-04 is therefore a joint architectural decision covering the Beta sampler and
the stream discipline together, and it remains OPEN**, with Option B recommended
and a feasibility proof-of-concept required in Step 0 (§19) before
`sim_contract.yaml` can encode either.

Boundary cases, unchanged whichever option wins:

| Case | Result |
|---|---|
| `a = b` | `x = a` |
| `m = a` | `α = 1`, `β = 5` — a valid Beta, no special case |
| `m = b` | `α = 5`, `β = 1` — likewise |

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

**Consumption is fixed; the transformation is conditional.** Two uniforms are
drawn for every risk on every iteration, whether or not it occurred, so stream
position never depends on a sampled value. The *severity inverse-CDF
transformation* is only evaluated when the risk occurred — there is nothing to
transform otherwise, and evaluating it would be pure waste.

These are different quantities and the performance model (§14) counts them
separately:

```
uniforms consumed              = fixed,       2 per risk per iteration
severity transforms evaluated  = conditional, expected p per risk per iteration
```

Worst case is `Probability = 1` on every risk, where the two coincide.

### 4.6 Sampling numerics over the accepted Double domain

Phase 5 accepted **any finite, correctly ordered** `Min/ML/Max` triple, including
negatives, supports crossing zero, and magnitudes near `Double` maximum, and it
built overflow-safe primitives rather than restricting the domain. **Phase 6 may
not silently narrow that domain**, and the naive formulas in §4.1–§4.3 do narrow
it: `b − a` overflows for `a = −MAX, b = +MAX`, and `(b − a)·(m − a)` overflows
far earlier than that.

**The rule:** for every finite, correctly ordered accepted triple, Phase 6 either
produces the mathematically valid sample using stable arithmetic, **or** issues an
explicitly authorised simulation numerical-range refusal naming the driver and
the stage. There is no third outcome, and no hidden magnitude assumption.

**Stable formulations.** The common device is to work in a normalised space,
mirroring Phase 5's conditioning scale: let `s = max(|a|, |m|, |b|)` and, where
`s > 0`, sample on `a/s, m/s, b/s` and rescale by `s` at the end. Every
intermediate is then bounded by a small multiple of 1.

| Quantity | Naive | Stable |
|---|---|---|
| Uniform interpolation | `a + u·(b − a)` | `(1 − u)·a + u·b` — a convex combination; each term bounded by `max(\|a\|,\|b\|)`, so it cannot overflow where the result does not. This is `StableConvex` territory and should reuse it |
| Triangular branch point | `c = (m − a)/(b − a)` | computed on normalised values, or from the ratio of two same-scaled differences |
| Triangular interpolation | `a + sqrt(u·(b−a)·(m−a))` | `a + sqrt(u)·sqrt(b−a)·sqrt(m−a)` on normalised values, rescaled — the product is never formed at full magnitude |
| Beta-PERT `α`, `β` | `1 + 4(m−a)/(b−a)` | the same ratio on normalised values; `r` is dimensionless and bounded by construction |

**Gate-A vectors required for this section specifically:**

| # | Case |
|---|---|
| 1 | large positive endpoints near `Double` maximum |
| 2 | large negative endpoints |
| 3 | support crossing zero |
| 4 | `a = −MAX_DOUBLE`, `m = 0`, `b = +MAX_DOUBLE`, or the largest authorised equivalent |
| 5 | subnormal / tiny spans |
| 6 | degenerate `a = m = b` |
| 7 | `m = a` |
| 8 | `m = b` |

Cases 1–4 are the ones that decide whether the accepted domain survived. If any
of them can only be handled by refusal, that refusal must be **authorised in the
contract**, not discovered at runtime.

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

### 5.2 State, constants and output — LOCKED (corrected in revision 2)

```
m1 = 4294967087        a12 =  1403580     a13n =  810728
m2 = 4294944443        a21 =   527612     a23n = 1370589
norm = 2.328306549295727688e-10
```

**The norm literal is not a second authority.** `2.328306549295727688e-10` and
`1.0 / (m1 + 1)` evaluate to the *same IEEE-754 double* — verified, both
`0x1.000000d00000bp-32`. Either spelling is admissible and neither can shift the
stream.

State is six `Double`s holding integers: `s1 = (s10,s11,s12)` in `[0, m1)`,
`s2 = (s20,s21,s22)` in `[0, m2)`. Not all of `s1` may be zero; not all of `s2`
may be zero.

**The combination — canonical, corrected.** Revision 1 proposed
`((p1 − p2) mod m1 + 1) · norm` and described the output as `(0,1]`. That is
**not** canonical MRG32k3a: it shifts the stream by one and would not reproduce
L'Ecuyer's published vectors. The canonical form is:

```
p1 = (a12·s11 − a13n·s10) mod m1        ; s10←s11, s11←s12, s12←p1
p2 = (a21·s22 − a23n·s20) mod m2        ; s20←s21, s21←s22, s22←p2

if p1 <= p2:  u = (p1 − p2 + m1) · norm
else:         u = (p1 − p2)      · norm
```

**Output domain: `0 < u < 1`, both ends excluded.**

- `p1 = p2` gives `u = m1·norm = m1/(m1+1) < 1` — the largest value, never `1`.
- `p1 > p2` gives at minimum `u = 1·norm ≈ 2.33e-10 > 0` — never `0`.

MRG32k3a therefore never emits exactly `0` or exactly `1`, and no sampler needs
a guard for either.

**D6-02 is CLOSED: canonical MRG32k3a, output `(0,1)`.**

### 5.2.1 Modular reduction — exact, and not VBA `Mod`

**VBA's `Mod` operator must not be used here.** It coerces its operands to an
integer type; the products below reach `6.03 × 10^15`, which overflows `Long`
outright, and the operator's rounding and result-type rules are the wrong
contract for a `Double`-held integer. The reduction is written out:

```
p = <the signed product difference>          ' exact integer in a Double
k = Fix(p / m)                               ' truncation toward zero
p = p - k * m
If p < 0 Then p = p + m                      ' positive remainder
```

**Every intermediate is exactly representable in IEEE-754 double**, which is what
makes the recurrence exact rather than approximately right. Verified worst cases
against `2^53 = 9,007,199,254,740,992`:

| Intermediate | Worst-case magnitude | Fraction of `2^53` |
|---|---|---|
| `a12 · s11` | `6,028,329,902,567,880` | 0.669 |
| `a13n · s10` | `3,482,050,075,698,608` | 0.387 |
| `a21 · s22` | `2,266,064,226,932,504` | 0.252 |
| `a23n · s20` | `5,886,603,607,816,338` | 0.654 |
| `\|p\|` before reduction | `6,028,329,902,567,880` | 0.669 |
| `k · m1` | `6,028,325,609,004,373` | 0.669 |

Headroom to `2^53` is a factor of 1.49 on the worst term. No intermediate can
lose a bit, so VBA `Double` arithmetic and Python exact-integer arithmetic must
agree exactly — which is what makes guarantee 1 of §5.7 testable.

### 5.2.2 Conformance to published vectors

Gate-A vectors must reproduce **L'Ecuyer's canonical published MRG32k3a
values**, obtained independently of PCCM. A vector set generated only by PCCM's
own implementation would prove self-consistency and nothing else. This is a hard
acceptance criterion for implementation step 2.

### 5.3 Seed domain — proposed LOCKED

`inpRandomSeed` is declared `integer`, optional, no validation. Proposed:

- admissible domain **`1 … 2147483646`** (whole numbers)
- `0`, negatives and non-integers are **refused**, not coerced
- blank is legal and means "choose one" (§5.4)

Rationale: a single positive 31-bit integer is what a user can type and record,
it maps cleanly onto the seeding function below, and it excludes the all-zero
state by construction.

**The seed is a `Long`, the state is not.** `2147483646 < 2^31 − 1`, so the
supplied and effective seed both fit VBA's signed `Long` and are stored and
published as integers. The RNG *state* words cannot be: `m1 − 1 = 4294967086`
exceeds `Long`, so `s10 … s22` are `Double`s holding integers (§5.2). Those are
two different types for two different things, deliberately.

### 5.4 Blank seed — request state versus run identity

Revision 1 put the effective seed in the simulation fingerprint while allowing a
blank input to generate it at run time. **That was circular**: with `C21` still
blank, the effective seed of the previous run cannot be re-derived from current
inputs, so the fingerprint could never be recomputed and every AUTO run would
read as STALE the moment it finished. Writing the generated seed back into
`inpRandomSeed` would "fix" it by silently editing a user input, which is worse.

The two notions are separated:

**A. Simulation request fingerprint** — recomputable from current inputs alone:

```
the accepted Phase-5 analytical fingerprint
iterations
seed_mode                     FIXED | AUTO
supplied_seed                 present only when seed_mode = FIXED
RNG_VERSION
SIM_METHOD_VERSION
```

A blank `C21` **is** canonical input state: it means `seed_mode = AUTO`, and that
is what the fingerprint records. The effective seed is not an input; it is an
outcome.

**B. Successful-run metadata** — what the run produced, stored alongside results:

```
run_id                        §13
request_fingerprint           the digest of A, at the time of the run
effective_seed                the integer actually used
timestamp
iterations
RNG_VERSION
SIM_METHOD_VERSION
```

Consequences, stated so none of them is a surprise:

- A successful AUTO run stays **CURRENT** while `C21` remains blank. Nothing
  about it is stale; the request has not changed.
- Running again with `C21` still blank is a **deliberate new run**: same request,
  new effective seed, new `run_id`. That is exactly what "blank means a new
  random sequence" says, and it is now the only thing it says.
- Typing the published effective seed into `C21` changes `seed_mode` from `AUTO`
  to `FIXED`. That is a **different request**, so the fingerprint changes and the
  prior result becomes STALE — correctly, because the user has asked a different
  question. Re-running then replays the identical stream.
- **A failed AUTO run records its effective seed in the attempt metadata**, once
  selected, so a failure is reproducible by typing that seed back in. A failure
  whose seed was never recorded would be a failure nobody could investigate.

### 5.4.1 What "new" guarantees

The contract says blank means a *new* random sequence, and a timestamp folded
into a 31-bit domain can collide — two runs a millisecond apart, or a modular
fold that happens to land on the same value.

**The rule: an AUTO run may not reuse the immediately preceding effective seed.**
The candidate is derived (D6-03), compared against the stored `effective_seed` of
the last successful run, and if equal it is advanced by a locked deterministic
step until it differs. This is not an entropy claim and does not pretend to be
one — it is a guarantee against *accidental immediate reuse*, which is what an
auditor would actually notice. Cryptographic quality is neither required nor
claimed.

### 5.5 Seeding function — D6-05 remains OPEN

A user seed of `1` must not produce a near-zero state, and two nearby seeds must
not produce correlated streams. The single `Long` seed is expanded into the six
state words, each reduced into its modulus and forced non-zero.

**Revision 1 casually suggested "splitmix-style". That is withdrawn.** SplitMix64
is a 64-bit integer/bitwise design; VBA's `LongLong` exists only on 64-bit
Office, bitwise operators on it behave differently across bitnesses, and the
whole point of §5.2.1 is that this generator's arithmetic stays inside exactly
representable `Double` integers. Importing a 64-bit bitwise dependency to seed a
`Double`-arithmetic generator would be choosing by fashion.

**Selection criteria, in order:** portable in VBA on both Office bitnesses ·
exact under `Double`-held integer arithmetic · straightforward in independent
Python integer arithmetic · easy to vector-lock · no bitwise operators.

**Leading candidate — a modular 31-bit expansion.** Successive application of a
locked Lehmer multiplicative generator, `x ← 48271 · x mod 2147483647`, taking
six successive outputs and reducing each into `m1` or `m2`, forcing any zero to
one. Its worst intermediate is `48271 × 2147483646 ≈ 1.04 × 10^14`, comfortably
inside `2^53`, so it is exact in `Double` and trivial in Python. Its weakness is
that consecutive small seeds produce related first words, which the reduction and
the generator's own mixing must be shown to absorb — that is what the vectors in
§15 family 3 exist to demonstrate.

Alternatives to weigh against it: a second independent MRG recurrence used only
for seeding; or fixed per-word offsets combined with the seed. **D6-05 stays open
and must close in Step 0.**

### 5.6 Stream discipline — depends on D6-04

Two designs, and which one applies follows from the Beta decision (§4.3):

**If a one-uniform sampler survives (Options A/C) — one global stream:**

1. drivers visited in ascending Permanent ID, ordinal UTF-16 comparison — the
   canonical order the Phase-5 fingerprint already locks
2. each Cost Line consumes exactly one uniform
3. each Risk consumes exactly two — Bernoulli then severity — always
4. one iteration consumes exactly `C + 2R` uniforms, a constant

Row-order invariance follows because reordering rows changes no Permanent ID.

**If a variable-draw sampler is adopted (Option B, recommended) — per-driver
substreams:**

1. each driver is assigned a substream by a deterministic function of its
   **Permanent ID**, not of its row
2. each driver advances only its own substream
3. iteration `i` starts each driver's substream at a computable offset

Row-order invariance follows because the assignment is keyed on Permanent ID and
never on position — the same guarantee, obtained structurally rather than by
arithmetic coincidence. Variable consumption inside a driver is then harmless,
because it cannot disturb any other driver.

**This second design is strictly better for Phase 7.** Under the global stream,
replaying driver *k*'s draw requires knowing every earlier driver's consumption;
under substreams it is a direct seek. §12's claim that Phase 7 needs no retained
per-driver samples holds under either, but only trivially under substreams.

### 5.7 Reproducibility — four guarantees, not one

Revision 1 claimed "bit-identical sample arrays on any machine". **That is
withdrawn**: it is one blanket claim covering four different propositions with
four different kinds of evidence, and under Option B's Beta generator the
strongest reading may not be supportable at all. Split:

**G1 — RNG conformance (exact).** MRG32k3a integer state and generated uniforms
agree **exactly** with independent canonical published vectors (§5.2.2). Justified
by §5.2.1: every intermediate is exactly representable, so `Double` and exact
integer arithmetic cannot diverge.

**G2 — same-runtime replay (exact).** The same resolved model, effective seed,
iteration count and method versions reproduce the **identical** VBA run digest
over the sample arrays. This is a determinism claim about one implementation and
needs no cross-language argument.

**G3 — row-order invariance (exact).** Reordering register rows without changing
any Permanent ID or model meaning reproduces the **identical** VBA run digest.
Delivered structurally by §5.6 under either stream design.

**G4 — cross-language sampler agreement (tolerance-bounded).** Python and VBA
sampler outputs agree under a **locked numerical comparison policy**:

- Uniform and Triangular: **exact**, being closed-form on the same exactly-shared
  uniform
- Bernoulli: **exact**
- Beta-PERT: **a locked tolerance**, expressed in ULPs or as a relative bound,
  because an exact match would require the two implementations to perform the
  same floating-point operations in the same order — which is a line-for-line
  port, and a port is not independent evidence

**The oracle stays genuinely independent.** Forcing bit-identical Beta results by
transcribing the VBA into Python would manufacture agreement and destroy the only
reason to have two implementations. The tolerance is the honest price of
independence, and it must be *stated in the contract*, not discovered by relaxing
a failing test.

### 5.8 Versioning

`RNG_VERSION` and `SIM_METHOD_VERSION` are separate integers, both stored with
the run. The first changes if the generator or seeding changes; the second if a
sampler or an accumulation order changes. Either invalidates a stored result the
way `FP_VERSION` already does.

---

## 6. Iterations

Three separate limits, never conflated:

| Concept | Value | Kind | Source |
|---|---|---|---|
| Hard refusal | `< 1000`, or non-whole | **business rule** | `input_contract.yaml`, LOCKED |
| Default | `10000` | business rule | LOCKED |
| `< 10000` advisory | belongs to **Model Check**, not Phase 6 | advisory | LOCKED |
| Tested performance target | `100000` | benchmark | `phase5_plan.md` §26 — target, not cap |
| **Business-rule maximum** | **NONE** | — | `input_contract.yaml`: "No upper limit is imposed" |
| **Simulation representation ceiling** | see §6.1 | **technical** | proposed |

**No business maximum is introduced.** The contract forbids one and Phase 6 does
not invent one.

### 6.1 The technical storage ceiling is real and must be computed, not awaited

Revision 1 proposed "no cap; refuse on allocation failure". **That is wrong once
`_SimData` persists one worksheet row per iteration** (§11.1): Excel's worksheet
capacity is a deterministic limit reached long before memory allocation fails,
and discovering it as a COM error mid-publish — after the RNG has been consumed
and the whole run computed — would be indefensible.

An Excel worksheet holds **1,048,576 rows**. With the `_SimData` header block
occupying the first `H` rows, the capacity is:

```
max_iterations_representable = 1048576 − H
```

`H` follows from the final `_SimData` layout, which is Step 1 contract work; with
a five-row header the ceiling is `1,048,571`.

**The refusal happens before anything is consumed.** A request above the
representable capacity is refused **before** sample allocation, **before** any
RNG draw and **before** the iteration loop, and the message names the *technical
storage limit* and the computed capacity — never a business rule, and never a
number that looks like one.

**D6-08 is narrowed but remains OPEN**: the exact value of `H`, and therefore of
the ceiling, cannot be fixed until the `_SimData` layout is settled in Step 1.
The *rule* above is proposed as locked; only the constant is outstanding.

**A contradiction worth recording.** The input contract says "no upper limit is
imposed" while a row-per-iteration persistence design imposes one. They are
reconcilable only by keeping the two limits in different categories — business
versus technical — and by never letting the technical one be presented as
validation. If independent review prefers no technical ceiling at all, the
consequence is summary-only persistence, which costs everything §11.1 lists.

---

## 7. Selected Confidence Level — LOCKED as a reporting selector

**The Phase-2 statement stands, unqualified.** Selected Confidence Level:

- does **not** change the generated distribution
- does **not** belong in the simulation request fingerprint
- does **not** make a successful simulation STALE
- does **not** make it INVALID
- does **not** require a re-run

Revision 1 proposed an `UNSELECTED` fourth simulation state. **That is rejected.**
A reporting selector has no business in the simulation state machine, and adding
a state to model "the user changed a dropdown" would make the machine describe
two unrelated things. The simulation state remains exactly **CURRENT / STALE /
INVALID**. If a separate reporting-state notion is ever genuinely needed, it will
be defined outside this machine.

**How the selection works instead.** Phase 6 computes and stores the *whole*
percentile ladder a selection could ask for:

- the ten selectable levels of `lstConfidenceLevels` — `P50 P55 P60 P65 P70 P75
  P80 P85 P90 P95`
- plus the fixed headline percentiles the Dashboard placeholder names — `P10`,
  and `P70`/`P90` which are already in the list

Selected Px is then a **deterministic lookup** from stored statistics, not a
computation over samples. Preferred implementation: a Results-side formula keyed
on `C18`, so changing the dropdown updates the reported figure with **no VBA
event, no recalculation and no re-run**. That also keeps the accepted "no change
handler" rule intact.

**P10 is a fixed reported statistic**, not a selectable confidence level.
`lstConfidenceLevels` remains `P50…P95` and no accepted input contract is
reopened.

**D6-06 is CLOSED. D6-10 is CLOSED by rejection: there is no `UNSELECTED`
state.**

---

## 8. Statistics

### 8.1 Moments

- **Sample mean**: `Σx / n`.
- **Standard deviation — LOCKED: the sample standard deviation, divisor
  `n − 1`.** This is a **reporting-method decision taken here**, not a default
  inherited from Excel or from any library. Rationale: the run is a sample drawn
  from the model's distribution, not the population itself, so `n − 1` is the
  consistent estimator of that distribution's spread. Stated explicitly because
  "`STDEV` vs `STDEVP`" is exactly the ambiguity this project refuses elsewhere.
  **D6-09 is CLOSED** and removed from the open table.
- Computed with the accepted Phase-5 safe-arithmetic primitives, in a numerically
  stable pass — **not** the naive `Σx²` form.

### 8.2 Percentile algorithm — LOCKED

**Linear-interpolation percentile, Hyndman–Fan Type 7**, equivalent to Excel's
`PERCENTILE.INC` and NumPy's `linear` method.

Revision 1 called this "nearest-rank with linear interpolation". That name is
wrong and is corrected: **nearest-rank is a different algorithm** — it selects an
actual order statistic and interpolates nothing. Only one name is used here, and
the equation below is the authority regardless of the name:

```
sorted x[0..n-1] ascending
h = (n − 1) · p                       ' p ∈ [0,1], e.g. P70 → p = 0.70
lo = floor(h) ; hi = min(lo + 1, n − 1)
Px = x[lo] + (h − lo) · (x[hi] − x[lo])
```

`n = 1` returns `x[0]`. `p = 0` returns the minimum; `p = 1` the maximum.

**No `WorksheetFunction.Percentile`** — it is already in the forbidden-construct
list, and a worksheet function inside the kernel would break the purity sweep.
The equation above is the authority; no Excel or library default is inherited,
and neither is the name.

### 8.3 The reported ladder — LOCKED

Stored for **nominal** and **PV** separately:

| Group | Values |
|---|---|
| Fixed headline percentiles | `P10`, `P50`, `P70`, `P90` |
| Full selectable ladder | `P50 P55 P60 P65 P70 P75 P80 P85 P90 P95` |
| Moments | mean, sample standard deviation |
| Extremes | minimum, maximum |

Eleven distinct percentiles per measure (`P10` plus the ten selectable). The
whole ladder is stored so that Selected Px is a lookup rather than a
recomputation (§7).

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

### 10.2 What the `SIM` section contains

The `SIM` section is the simulation half of the **request** fingerprint of
§5.4 A, appended after `RISK`:

| Field | In? | Why |
|---|---|---|
| Iteration count | **yes** | changes the distribution |
| `seed_mode` (`FIXED` / `AUTO`) | **yes** | canonical input state; a blank cell *is* a value |
| `supplied_seed` | **yes, only when `seed_mode = FIXED`** | it is an input then, and only then |
| `RNG_VERSION` | **yes** | changes the numbers for identical inputs |
| `SIM_METHOD_VERSION` | **yes** | same |
| **Effective seed** | **no** | it is an *outcome*, not an input — §5.4. Including it made the AUTO fingerprint unrecomputable |
| Selected Confidence Level | **no** | §7 — it selects from the distribution, it does not generate it |

The analytical `HEADER`/`COST`/`RISK` sections keep their positions and their
bytes, so the Phase-5 subset stays comparable across phases.

### 10.3 The three states

Mirroring the accepted Phase-5 attempt/state philosophy rather than inventing a
second machine:

| State | Meaning |
|---|---|
| **CURRENT** | the stored **simulation request fingerprint** matches the recomputed one (§5.4 A), and the last simulation attempt succeeded |
| **STALE** | a stored successful result exists, but the request fingerprint no longer matches — a changed model, iteration count, seed mode, supplied seed or method version |
| **INVALID** | prerequisites refuse (including the Phase-5 prerequisite of §10.6), or the last attempt failed |

**There is no fourth state.** Revision 1's `UNSELECTED` is rejected (§7).

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

### 10.6 The Phase-5 analytical prerequisite — D6-14

**Nothing in revision 1 said how the two result layers relate**, and without a
rule Phase 6 could publish a CURRENT simulation computed from one input snapshot
while the visible Phase-5 analytical outputs came from another. Two published
result sets describing different models, both marked current, is precisely the
class of defect this project has spent Phase 5 eliminating.

Options:

| | Design |
|---|---|
| **A** | `PCCM_RunSimulation` **requires** Phase-5 calculation status `CURRENT`; refuses otherwise |
| **B** | `PCCM_RunSimulation` silently refreshes the Phase-5 calculation first |
| **C** | a single fully transactional operation publishing both layers together |

**Recommendation: A.**

- **B** would make a simulation endpoint silently mutate accepted Phase-5
  publication semantics — a Phase-6 call rewriting Phase-5 outputs is exactly the
  boundary violation the module split exists to prevent.
- **C** is defensible but doubles the transactional surface of an already
  transactional publish, for a convenience that belongs in the presentation
  phase.

Under **A**: if the analytical state is STALE or INVALID, **the simulation
refuses**, the prior successful simulation is left untouched, and the refusal
detail names the analytical state as the reason. A later user-facing *Run
Simulation* control (Phase 8) may orchestrate Calculate-then-Simulate; that is a
presentation decision and does not belong in the engine.

**The proof that both layers describe the same model** is direct: the simulation
request fingerprint (§5.4 A) *contains* the Phase-5 analytical fingerprint as its
first component. A simulation is CURRENT only if that embedded analytical
component still matches, which is the same test Phase 5 uses for its own
currency. One fingerprint, checked twice, rather than two mechanisms that could
disagree.

**D6-14 is OPEN**, with A recommended.

### 10.7 Run identity — D6-15

Results already reserves a Run ID and nothing defines it. It is **audit
metadata**: it is not a source of randomness, it does not enter the fingerprint,
and it does not influence one sampled value.

Proposed semantics:

| Question | Proposal |
|---|---|
| When allocated | at the moment a run **succeeds**, as the last step before publish |
| Do failed attempts consume one? | **No.** A failed run has no identity to cite; its attempt metadata carries the effective seed and the failure detail instead |
| Unique within the workbook? | Yes |
| Monotonic? | Yes — a persisted counter, incremented on success |
| Survives Save/Reopen? | Yes; it is stored in the workbook, not derived at load |
| Relation to effective seed | Independent. One `run_id` has exactly one `effective_seed`; the same seed may legitimately recur under a different `run_id` |
| External dependencies | **None.** No GUID, no COM identity source, no `Now()`-derived uniqueness — a persisted monotonic integer is portable, auditable and trivially testable |

**D6-15 is OPEN**, with the above recommended.

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

### 11.1 `_SimData` persists iteration order — LOCKED

Revision 1 proposed persisting the sorted sample. **That is rejected: sorting
destroys iteration identity**, and iteration identity is what makes a run
auditable. Losing it costs replay, exact run digests, Phase-7 sensitivity pairing
(which must pair *the same iteration's* driver value with *the same iteration's*
total), and any diagnosis of an individual iteration.

`_SimData` therefore stores, in **canonical iteration order**:

| Column | Meaning |
|---|---|
| `iteration_index` | `1 … n`, the identity |
| `total_nominal` | the iteration's total, SAR nominal |
| `total_pv` | the iteration's total, SAR present value |

plus the run-identity block of §5.4 (`run_id`, `effective_seed`,
`request_fingerprint`, `iterations`, versions, timestamp).

**Statistics sort copies in memory.** The persisted arrays are never permuted.
Saving a few megabytes by discarding iteration identity is not a trade this
architecture should make.

**D6-13 is CLOSED by adoption: raw iteration-ordered totals are persisted.**

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

| Quantity | Count | Note |
|---|---|---|
| Uniforms consumed per iteration | `200 + 2×100 = 400` | **fixed**, independent of outcomes |
| Uniforms consumed per run | `4.0 × 10^7` | |
| Bernoulli comparisons per run | `1.0 × 10^7` | one per risk per iteration |
| Cost-line inverse-CDF evaluations | `2.0 × 10^7` | always evaluated |
| Severity inverse-CDF evaluations | `≤ 1.0 × 10^7` | **conditional** — expected `Σp` per iteration; worst case `p = 1` throughout |
| Multiply-accumulates per run | `≤ 6.0 × 10^7` | nominal + PV, on evaluated contributions |
| Sorts | 2, of `10^5` elements | `~3.4 × 10^6` comparisons total |
| **Worksheet / COM calls inside the loop** | **0** | the locked invariant |

Under Option B (§4.3) the uniform count becomes variable within a driver's own
substream; the *architecture* is unchanged and the worst case rises to roughly
`2.6 ×` the Beta samples drawn.

**Memory, by category.** Revision 1's "Phase-6 total < 2 MB" was not a defensible
peak statement: it counted the retained samples and nothing else. Corrected, at
100,000 iterations:

| Category | Contents | Bytes |
|---|---|---|
| **Hot-kernel resident** | Phase-5 structures (`< 100 KB`) + the two iteration-ordered sample arrays | `~1.7 MB` |
| **Statistics working** | two sorted *copies* of the sample arrays | `1.6 MB` |
| **Publication buffer** | a `Variant` array for the bulk `_SimData` write: `100,000 × 3 × 16 B` | `4.8 MB` |
| **Workbook persisted** | 300,000 numeric cells on `_SimData`, compressed in the `.xlsm` | `~1–3 MB` on disk |
| **Peak resident** | hot + statistics + publication, coexisting during the publish | **`~8.1 MB`** |

Deferred retentions, for comparison, and the reason they are deferred:

| Deferred item | Bytes at target |
|---|---|
| annual samples, `25 × 100,000 × 8` | 20 MB |
| per-driver samples, `300 × 100,000 × 8` | 240 MB |

The publication buffer can be **chunked** if the peak is judged too high —
writing `_SimData` in blocks of, say, 20,000 rows reduces that 4.8 MB to under
1 MB at the cost of a handful of extra COM calls, all of them **after** the loop
and therefore outside the hot-path rule. That is a Step-10 implementation choice,
not an architectural one.

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

**Step 0 comes first, and revision 1 did not have it.** The old sequence created
`spec/sim_contract.yaml` in step 1 while the semantics that contract must encode
were still scheduled to be decided in steps 3, 4 and 7. A contract cannot be the
authority for decisions taken after it exists.

| # | Step | Files allowed to change | New authority | Gate-A acceptance |
|---|---|---|---|---|
| **0** | **Authority closure.** Close every decision the contract must encode: canonical MRG32k3a semantics · seed domain · AUTO-seed source and freshness · seeding expansion · **Beta sampler and stream discipline jointly** · percentile method · standard-deviation method · contingency · persistence model · technical iteration ceiling · simulation request fingerprint · analytical prerequisite · run-id semantics · state model. Includes a **feasibility proof-of-concept** where an algorithm genuinely cannot be chosen on paper — currently D6-04 alone | `docs/` only | the closed decisions | every §20 item resolved and independently reviewed; **no contract file yet** |
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

**The feasibility proof-of-concept in Step 0 is planning work, not
implementation.** It exists to measure the Beta sampler options — accuracy,
operation count, achievable cross-language tolerance — so D6-04 can be closed on
evidence. It is throwaway measurement code under `docs/` or a scratch path, it
ships nothing, and it must not become the Phase-6 implementation by default.

---

## 20. Open decisions — revised

Closed by review round 2 and removed from this table: **D6-01** (scope split),
**D6-02** (canonical MRG32k3a, `(0,1)`), **D6-06** (P10 fixed, list unchanged),
**D6-07** (contingency baseline `A`), **D6-09** (sample standard deviation, a
reporting-method decision taken here), **D6-10** (rejected — no `UNSELECTED`
state), **D6-11** (scope the Phase-4 guards, never weaken them), **D6-12** (no
user cancellation), **D6-13** (raw iteration-ordered persistence adopted).

Still genuinely unresolved:

| # | Decision | Options | Existing evidence | Recommended | Consequence | Must close before |
|---|---|---|---|---|---|---|
| **D6-03** | AUTO-seed source and freshness mechanism | (a) timestamp folded into the domain, with the §5.4.1 anti-reuse rule; (b) counter-derived from the persisted `run_id`; (c) refuse blank | contract says blank means "a new random sequence", so (c) contradicts it | **(b)** — a persisted counter is deterministic, auditable, and cannot collide by construction, where a folded timestamp can | Whether "new" is guaranteed or merely likely | **Step 0** |
| **D6-04** | **Beta sampler and stream discipline, jointly** | (a) one-uniform inverse CDF + global fixed-consumption stream; (b) exact variable-draw generator + per-driver substreams; (c) precomputed table + global stream | measured: (a) `2.7 × 10^12` flops/run — infeasible; (c) `5.9 × 10^-3` normalised error at 4096 nodes — indefensible; (b) `~3 × 10^9` flops/run, exact | **(b)**, subject to the Step-0 proof-of-concept | Decides the stream architecture, the jump-ahead work, the oracle's tolerance policy and Phase 7's replay mechanism | **Step 0** |
| **D6-05** | Seed-expansion algorithm | (a) modular 31-bit Lehmer expansion; (b) a second MRG recurrence used only for seeding; (c) fixed per-word offsets | none in repository; splitmix withdrawn as non-portable | **(a)**, if the vectors show nearby seeds decorrelate | Two seeds must not produce related streams | **Step 0** |
| **D6-08** | The technical iteration ceiling constant | (a) `1048576 − H` from the final `_SimData` layout; (b) a lower stated ceiling | Excel row capacity; the rule is settled in §6.1, only `H` is not | **(a)** | The exact refusal threshold | **Step 1** (needs the layout) |
| **D6-14** | Analytical prerequisite / orchestration | (a) require Phase-5 `CURRENT`; (b) simulation refreshes the calculation; (c) one transactional operation for both | Phase-5 publication semantics are accepted and must not be silently changed | **(a)** | Whether Phase 6 can mutate Phase-5 outputs | **Step 0** |
| **D6-15** | Run-ID semantics | (a) persisted monotonic success counter; (b) GUID; (c) timestamp-derived | Results reserves a Run ID; nothing defines it | **(a)** | Audit identity; no computational effect | **Step 0** |

---

## 21. Change ledger — revision 1 to revision 2

| Round-2 item | Section changed | What changed |
|---|---|---|
| §1 canonical MRG32k3a | **§5.2 rewritten**, **§5.2.1 and §5.2.2 new** | The `((p1−p2) mod m1 + 1)·norm` form is withdrawn; the canonical two-branch combination is locked with output `(0,1)`. VBA `Mod` is forbidden for the recurrence and an exact positive-remainder reduction is specified. Every intermediate proved exactly representable, with measured worst cases against `2^53`. Conformance to *published* vectors made an acceptance criterion. Verified that the `norm` literal and `1/(m1+1)` are the same double |
| §2 seed / fingerprint circularity | **§5.4 rewritten, §5.4.1 new, §10.2 rewritten** | Request fingerprint and successful-run metadata separated; `seed_mode = FIXED\|AUTO` is canonical input state; effective seed moved out of the fingerprint into run metadata; AUTO runs stay CURRENT while blank; failed AUTO runs retain their effective seed; anti-reuse rule added |
| §2 seed type | **§5.3** | `EffectiveSeed` is a `Long`; the *state* words stay `Double` because `m1 − 1` exceeds `Long`, and the reason is stated |
| §3 no `UNSELECTED` | **§7 rewritten, §10.3** | Fourth state rejected; the full selectable ladder plus fixed headlines is stored; selection becomes a Results-side deterministic lookup with no event and no re-run |
| §4 percentile naming | **§8.2** | Renamed to Hyndman–Fan Type 7 / `PERCENTILE.INC`-equivalent; the "nearest rank" misnomer removed and the distinction stated |
| §5 `_SimData` order | **§3.3, §11.1 new** | Sorted persistence rejected; iteration order persisted with `iteration_index`; statistics sort copies |
| §5 memory model | **§14 rewritten** | "< 2 MB" withdrawn; four categories plus a peak of `~8.1 MB`; chunked publication noted as a Step-10 option |
| §6 iteration ceiling | **§6 rewritten, §6.1 new** | Business maximum stays NONE; a computed technical storage ceiling `1048576 − H` is added, refused *before* allocation and any RNG draw, and never presented as validation |
| §7 Beta + stream jointly | **§4.3 rewritten, §5.6 rewritten** | Three options costed with measured numbers; A rejected as infeasible, C rejected as inaccurate; **Option B recommended, which changes the recommendation from a global stream to per-driver substreams** |
| §8 risk draw wording | **§4.5, §14** | Consumption (fixed) separated from severity transformation (conditional); counts split |
| §9 numerical stability | **§4.6 new** | Stable formulations for all four quantities; normalisation by `max(\|a\|,\|m\|,\|b\|)`; eight extreme-domain vector cases; refusal must be contract-authorised |
| §10 reproducibility | **§5.7 rewritten** | Blanket bit-identical claim withdrawn; four guarantees G1–G4 with distinct evidence; Beta gets a locked tolerance rather than a manufactured exact match |
| §11 seed expansion | **§5.5 rewritten** | Splitmix withdrawn as non-portable across Office bitnesses; modular 31-bit Lehmer expansion is the leading candidate; D6-05 stays open |
| §12 analytical prerequisite | **§10.6 new** | D6-14 added; option A recommended; the proof is the analytical fingerprint embedded in the simulation request fingerprint |
| §13 run identity | **§10.7 new** | D6-15 added; persisted monotonic success counter; failures consume none; no GUID or COM dependency |
| §14 blank-seed freshness | **§5.4.1 new** | What "new" guarantees, stated: no reuse of the immediately preceding effective seed |
| §16 standard deviation | **§8.1** | Locked as a reporting-method decision taken here; D6-09 closed and removed |
| §17 Step 0 | **§19** | Authority-closure step added before the contract, with the feasibility proof-of-concept scoped as throwaway measurement |

**Decisions closed:** D6-01, D6-02, D6-06, D6-07, D6-09, D6-10, D6-11, D6-12,
D6-13. **Still open:** D6-03, D6-04, D6-05, D6-08, D6-14, D6-15.

### New contradiction discovered

**One**, recorded in §6.1: `input_contract.yaml` states that no upper limit is
imposed on iterations, while a row-per-iteration persistence design imposes a
deterministic one at Excel's worksheet capacity. They reconcile only by keeping
the limits in different categories — **business** versus **technical** — and by
never letting the technical ceiling be presented as validation. If review prefers
no technical ceiling at all, the consequence is summary-only persistence and the
loss of everything §11.1 lists.

### Revised Gate-A / Gate-B evidence implications

| | Change from revision 1 |
|---|---|
| **Gate A** | Vectors must match **published** MRG32k3a values, not self-generated ones (§5.2.2). Eight extreme-domain sampling vectors added (§4.6). The comparison policy becomes **per-sampler**: exact for Uniform, Triangular and Bernoulli; a locked tolerance for Beta (§5.7 G4). Under Option B, substream assignment and jump-ahead need their own vectors |
| **Gate B** | Row-order invariance is now a **structural** property of substream assignment rather than of fixed consumption, so the test is unchanged but what it proves is stronger. A pre-flight refusal test is added for a request above the technical storage ceiling — it must refuse **before** any RNG draw. `_SimData` iteration order must be asserted, not just its row count |

---

## 22. What this plan does NOT claim

1. **No Phase-6 code exists.** No VBA, no builder change, no contract, no test,
   no simulation code in any language.
2. **No Windows or Excel runtime has been executed for Phase 6.**
3. **`spec/sim_contract.yaml` does not exist** and may not be created until
   Step 0 closes every decision it must encode.
4. **Nothing here is accepted.** Six decisions remain open, and the largest of
   them — D6-04 — may still change the stream architecture.
5. **The Phase-5 baseline is untouched.** `f571154` remains the accepted
   executable baseline; this document changes no file under `src/`, `spec/`,
   `builder/`, `bootstrap/` or `tests/`.
