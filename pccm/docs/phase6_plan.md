# Phase 6 — Stochastic simulation layer

**Status: PLANNING ONLY — revision 3, after review round 3.
NOT ACCEPTED. NO IMPLEMENTATION EXISTS. NO WINDOWS/EXCEL RUNTIME HAS BEEN
EXECUTED. `spec/sim_contract.yaml` DOES NOT EXIST.**

Phase 5 is closed. The accepted executable baseline is
`f571154118083e569e1fb9fbf9bf72852cc2d568`; the closure head is `28fa613`. This
document proposes what Phase 6 should be, from the authorities already in the
repository. It changes no code, no contract and no generated artefact.

**Revision 3 describes ONE candidate architecture throughout.** Revision 2
recommended a variable-draw Beta sampler while several sections still described
the one-uniform fixed-consumption design it replaced. Every such statement has
been rewritten; the stale-statement ledger is §21.

**The candidate architecture, in one paragraph.** Beta-PERT is sampled by an
acceptance/rejection method (Cheng BB/BC), so RNG consumption is variable. Each
*simulation component* — one per Cost Line, two per Risk — therefore owns its own
deterministic MRG32k3a **stream**, advanced sequentially across iterations.
Row-order invariance comes from assigning streams in canonical Permanent-ID
order, not from counting draws. Because consumption is variable, there is no
arithmetic seek to iteration *i*; replay is by resetting a stream to its initial
state and re-running.

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

### 3.3 New structures

Revision 2 left a stale `SimulationRequest` carrying `EffectiveSeed` and
`SeedWasSupplied`, which contradicted its own §5.4. Corrected: **request** and
**run context** are different objects with different lifetimes.

```vb
Type RngState                 ' MRG32k3a: six Doubles holding integers
    S10 As Double: S11 As Double: S12 As Double
    S20 As Double: S21 As Double: S22 As Double
End Type

Type SimulationRequest        ' what was ASKED FOR; recomputable from inputs
    Iterations    As Long
    SeedMode      As Long     ' SEED_MODE_FIXED | SEED_MODE_AUTO
    SuppliedSeed  As Long     ' meaningful only when SeedMode = FIXED
    RngVersion    As Long
    MethodVersion As Long
End Type

Type SimulationRunContext     ' what HAPPENED; outcome, never an input
    RunId              As Long
    EffectiveSeed      As Long
    AutoNonce          As Long   ' 0 when SeedMode = FIXED
    RequestFingerprint As String
    ResultDigest       As String
    Timestamp          As Date
    IterationsRun      As Long
End Type

Type ComponentStream          ' one per simulation component (§5.6)
    ComponentKind As Long     ' COST_SAMPLE | RISK_OCCURRENCE | RISK_SEVERITY
    PermanentId   As String
    State         As RngState
End Type
```

plus two `Double` arrays of length `Iterations` — the retained nominal totals and
PV totals, **held in canonical iteration order**. The percentile step sorts
*copies* (§11.1); the originals are never permuted, because their index *is* the
iteration identity.

### 3.4 Phase separation

| When | Work |
|---|---|
| **once, before the loop** | resolve inputs; run the Phase-5 checks; build `DriverFactors`, `YearFactors` and the weight vectors; read Iterations and Random Seed; derive the effective seed; build the `C + 2R` component streams by jump-ahead from the base state; allocate the two sample arrays |
| **once per driver, before the loop** | Cheng dispatch and its shape-dependent constants, which depend only on `α, β` and are therefore fixed for the whole run (§4.3) |
| **per iteration, per Cost Line** | draw from its sampling stream — **a variable number of uniforms**, distribution-dependent; one Cheng evaluation; two multiply-accumulates |
| **per iteration, per Risk** | one uniform from its **occurrence** stream, always; a Cheng draw from its **severity** stream **only if it occurred**; two multiply-accumulates if it occurred |
| **per iteration** | write the two accumulated totals into the sample arrays at index `i` |
| **after the loop** | copy each sample array and sort the copies; compute moments and percentiles; compute contingency; compute the result digest; compute the request fingerprint; write `_SimData` in iteration order and Results in one transactional publish |

**Consumption is not constant and the plan no longer pretends it is.** Only the
*occurrence* draw is one-per-risk-per-iteration; everything else varies with the
sampler's acceptance path.

---

## 4. Sampling semantics

Let `a = Min`, `m = MostLikely`, `b = Max`, with the Phase-5 ordering guarantees
already established. **No positivity is assumed**; every formula below is valid
for negative `a`, `m`, `b`.

Uniform and Triangular are **inverse-CDF on a single uniform**. Beta-PERT is
**not** — it uses an acceptance/rejection method and consumes a variable number
(§4.3). Revision 2's blanket statement that "all sampling is inverse-CDF on a
single uniform" was false under its own recommendation and is withdrawn.

### 4.1 Uniform — one uniform

```
x = (1 − u)·a + u·b                     ' stable convex form, §4.6
```

`m` is not read (D1). Degenerate `a = b` returns `a`.

### 4.2 Triangular — one uniform, inverse CDF

With `c = (m − a)/(b − a)`, computed on normalised values (§4.6):

```
u ≤ c :  x = a + sqrt( u · (b − a) · (m − a) )
u > c :  x = b − sqrt( (1 − u) · (b − a) · (b − m) )
```

Boundary cases:

| Case | Result |
|---|---|
| `a = b` | `x = a`; one uniform still drawn, so the stream position is unaffected by the data |
| `m = a` | `c = 0`; the `u > c` branch always taken |
| `m = b` | `c = 1`; the `u ≤ c` branch always taken |
| `u → 0` | `x → a`; `u = 0` is never produced (§5.2) |
| `u → 1` | `x → b`; `u = 1` is never produced |

### 4.3 Beta-PERT — Cheng BB/BC, a variable number of uniforms

With `λ = 4`:

```
r = (m − a)/(b − a)                     ' shape ratio in [0,1], normalised (§4.6)
α = 1 + 4r                              ' in [1,5]
β = 1 + 4(1 − r)                        ' in [1,5]
x = (1 − y)·a + y·b   where  y ~ Beta(α, β)
```

whose mean is exactly `(a + 4m + b)/6`, consistent with the accepted `PertMean`.

**Dispatch rule — LOCKED, including the boundary:**

```
if min(α, β) > 1  →  Cheng BB
else              →  Cheng BC          ' the equality case belongs to BC
```

`α = 1` occurs exactly when `m = a`; `β = 1` exactly when `m = b`. Both therefore
dispatch to **BC**, and neither is a special case bolted on afterwards.

**Terminology, stated precisely.** Cheng BB/BC is a
**distribution-exact acceptance/rejection method**: in exact arithmetic it samples
the target Beta with no inverse-CDF approximation. That is **not** a claim that
its floating-point output is mathematically exact, and **not** a claim that two
independent implementations produce bit-identical values. Revision 2's
unqualified "exact" is withdrawn; the consequences are §7.

**Why not the alternatives** — measured, and the measurement artefacts are Step-0
deliverables (§17), not accepted numbers:

| | Architecture | Measured | Verdict |
|---|---|---|---|
| A | one-uniform numerical inverse CDF | incomplete-beta continued fraction needs up to **91 iterations** over this family; ~91,000 flops/sample; `2.7 × 10^12` flops at the design target | infeasible |
| C | precomputed per-driver table | max normalised error `5.9 × 10^-3` at 4096 nodes — at `r = 0.01`, `α = 1.04` and the inverse CDF has near-infinite tail slope | indefensible |
| **B** | **Cheng BB/BC** | peak density `≤ 5.0` over the family, so acceptance is bounded well away from zero | **recommended** |

**D6-04 remains OPEN** and now covers the sampler, the component-stream scheme
and the jump arithmetic **as one decision**, closing only on the Step-0
feasibility package.

### 4.4 Cost Line

```
unitCost = sample(dist, a, m, b)        ' 1 uniform if Uniform/Triangular,
                                        ' variable if Beta-PERT
contribNom = unitCost × Quantity × Knom
contribPv  = unitCost × Quantity × Kpv
```

Quantity is **deterministic** and applied after sampling — accepted semantics,
unchanged.

### 4.5 Risk — two streams, not two draws from one

```
occurred = ( u_occurrence < Probability )      ' 1 uniform, OCCURRENCE stream,
                                               ' always, every iteration
if occurred:
    severity   = sample(dist, a, m, b)         ' SEVERITY stream, variable count
    contribNom = severity × Knom
    contribPv  = severity × Kpv
else:
    contribNom = 0 ;  contribPv = 0            ' severity stream not advanced
```

`Probability` is **not** folded into `Knom`/`Kpv` — the locked Phase-5 rule.
Strict `<`, so `Probability = 0` never occurs and `Probability = 1` always
occurs, given `u ∈ (0,1)`.

**Revision 2 said every risk consumes exactly two uniforms while also saying the
severity transform is conditional. Both cannot hold under a variable-draw
sampler, and the contradiction is resolved by separating the streams**, not by
drawing a uniform nobody uses:

- the **occurrence** stream advances by exactly one uniform per iteration,
  unconditionally, so the occurrence path is a function of `Probability` and the
  seed alone
- the **severity** stream advances only when the risk occurred, by however many
  uniforms the sampler needed

Because they are different streams, variable severity consumption **cannot**
perturb the occurrence path. That is the property §5.6 exists to guarantee.

### 4.6 Numerics over the accepted Double domain

Phase 5 accepted **any finite, correctly ordered** triple, including negatives,
supports crossing zero, and magnitudes near `Double` maximum, and built
overflow-safe primitives rather than restricting the domain. **Phase 6 may not
silently narrow that domain**, and the naive formulas do: `b − a` overflows for
`a = −MAX, b = +MAX`, and `(b − a)·(m − a)` overflows far earlier.

**The rule:** for every finite, correctly ordered accepted triple, Phase 6 either
produces the mathematically valid result using stable arithmetic, **or** issues an
explicitly contract-authorised numerical-range refusal naming the driver and the
stage. No third outcome, no hidden magnitude assumption.

**The common device** is to work in a normalised space, mirroring Phase 5's
conditioning scale: with `s = max(|a|, |m|, |b|)` and `s > 0`, operate on
`a/s, m/s, b/s` and rescale by `s` at the end.

**This discipline extends to the whole pipeline, not only the sampler.**
Revision 2 stopped at §4.6 and said "a numerically stable pass" for the
statistics, which is not a specification:

| Stage | Naive hazard | Stable strategy |
|---|---|---|
| Uniform / Triangular / Beta rescale | `a + u·(b − a)` | convex form `(1 − u)·a + u·b`; reuse `StableConvex` |
| Triangular branch point and interpolation | `(b−a)·(m−a)` | normalised values, rescaled after the `sqrt` |
| `α`, `β` parameterisation | `(m−a)/(b−a)` | the same ratio on normalised values; `r` is dimensionless |
| Per-driver contribution | `x × Quantity × K` | the accepted `TripleProduct` / `SafeProduct`, already proven in Phase 5 |
| Accumulating 300 contributions | running sum overflow with cancelling signs | the accepted `SafeSignedSum` / `SafeAccumulate` |
| Sample mean | `Σx / n` | scale-aware accumulation, or a shifted mean about a running scale |
| Sample standard deviation | naive `Σx²`; **also Welford's `δ = x − mean`, which overflows for opposite-sign near-max samples** | a scale-normalised two-pass formulation, or an authorised refusal |
| Type-7 percentile interpolation | **`x_lo + f·(x_hi − x_lo)` overflows at the difference even where the convex result is finite** | convex form `(1 − f)·x_lo + f·x_hi` |
| Contingency subtraction | `Px − A` | the accepted `SafeSubtract` |

**Gate-A extreme-domain vectors must exercise the complete pipeline**, sampler
through published statistic — not the sampler alone:

| # | Case |
|---|---|
| 1 | large positive endpoints near `Double` maximum |
| 2 | large negative endpoints |
| 3 | support crossing zero |
| 4 | `a = −MAX_DOUBLE`, `m = 0`, `b = +MAX_DOUBLE`, or the largest authorised equivalent |
| 5 | subnormal / tiny spans |
| 6 | degenerate `a = m = b` |
| 7 | `m = a` — and therefore Cheng **BC** |
| 8 | `m = b` — and therefore Cheng **BC** |
| 9 | a sample set of opposite-sign near-max totals, driven through mean, standard deviation and every percentile |
| 10 | contingency where `Px` and `A` are both near-max and of opposite sign |

Cases 9 and 10 are new in revision 3 and are the ones that decide whether the
*statistics* preserved the accepted domain.

---
## 5. RNG — the locked contract

### 5.1 Algorithm

**MRG32k3a** (L'Ecuyer 1999), as named in the locked roadmap. `Rnd` is not
merely inadvisable: its period is `2^24`, and the design target needs
`≥ 4 × 10^7` uniforms — exhausted by more than four orders of magnitude, quite
apart from being unreproducible and version-dependent.

### 5.2 State, constants and output — LOCKED

```
m1 = 4294967087        a12 =  1403580     a13n =  810728
m2 = 4294944443        a21 =   527612     a23n = 1370589
norm = 2.328306549295727688e-10
```

**The norm literal is not a second authority.** `2.328306549295727688e-10` and
`1.0 / (m1 + 1)` are the *same IEEE-754 double* — verified, both
`0x1.000000d00000bp-32`.

State is six `Double`s holding integers: `s1 = (s10,s11,s12)` in `[0, m1)`,
`s2 = (s20,s21,s22)` in `[0, m2)`; `s1` not all zero, `s2` not all zero.

```
p1 = (a12·s11 − a13n·s10) mod m1        ; s10←s11, s11←s12, s12←p1
p2 = (a21·s22 − a23n·s20) mod m2        ; s20←s21, s21←s22, s22←p2

if p1 <= p2:  u = (p1 − p2 + m1) · norm
else:         u = (p1 − p2)      · norm
```

**Output domain: `0 < u < 1`, both ends excluded.** `p1 = p2` gives
`u = m1/(m1+1) < 1`; `p1 > p2` gives at minimum `u = norm ≈ 2.33e-10 > 0`.

**D6-02 is CLOSED: canonical MRG32k3a, output `(0,1)`.**

### 5.2.1 Modular reduction — exact, and not VBA `Mod`

**VBA's `Mod` must not be used here.** It coerces operands to an integer type;
the products reach `6.03 × 10^15`, which overflows `Long` outright.

```
p = <the signed product difference>          ' exact integer in a Double
k = Fix(p / m)                               ' truncation toward zero
p = p - k * m
If p < 0 Then p = p + m                      ' positive remainder
```

Verified against `2^53 = 9,007,199,254,740,992`:

| Intermediate | Worst case | Fraction of `2^53` |
|---|---|---|
| `a12 · s11` | `6,028,329,902,567,880` | 0.669 |
| `a13n · s10` | `3,482,050,075,698,608` | 0.387 |
| `a21 · s22` | `2,266,064,226,932,504` | 0.252 |
| `a23n · s20` | `5,886,603,607,816,338` | 0.654 |
| `k · m1` | `6,028,325,609,004,373` | 0.669 |

Headroom 1.49× on the worst term. **This guarantee covers the base recurrence
only** — the jump arithmetic is a different problem entirely (§5.9).

### 5.2.2 Conformance to published vectors

Gate-A vectors must reproduce **L'Ecuyer's canonical published MRG32k3a values**,
obtained independently of PCCM. A vector set generated only by PCCM's own
implementation proves self-consistency and nothing else. Hard acceptance
criterion for step 2.

### 5.3 Seed domain — proposed LOCKED

- admissible domain **`1 … 2147483646`**, whole numbers
- `0`, negatives and non-integers are **refused**, not coerced
- blank is legal and means `seed_mode = AUTO` (§5.4)

**The seed is a `Long`; the state words are not.** `2147483646 < 2^31 − 1` fits
VBA's signed `Long`. `m1 − 1 = 4294967086` does not, so `s10 … s22` are `Double`s
holding integers. Two types for two different things, deliberately.

### 5.4 Blank seed — request state versus run identity

Putting the effective seed in the fingerprint while allowing a blank input to
generate it is circular: with `C21` blank the previous effective seed cannot be
re-derived, so every AUTO run would read as STALE the moment it finished.
Writing the generated seed back into `inpRandomSeed` would "fix" that by silently
editing a user input, which is worse.

**A. Simulation request fingerprint** — recomputable from current inputs alone:
the canonical `HEADER/COST/RISK` prefix plus a `SIM` section carrying
`iterations`, `seed_mode`, `supplied_seed` (FIXED only), `RNG_VERSION`,
`SIM_METHOD_VERSION`. See §10.1 for how the two digests relate.

**B. Successful-run metadata** — §11 identities: `run_id`, `effective_seed`,
`auto_nonce`, `request_fingerprint`, `result_digest`, timestamp, iterations,
versions.

Consequences:

- A successful AUTO run stays **CURRENT** while `C21` remains blank. The request
  has not changed.
- Running again while blank is a **deliberate new run**: same request
  fingerprint, new nonce, new effective seed, new `run_id`, **new result
  digest**. §11 explains why that is not a contradiction.
- Typing the published effective seed into `C21` switches `seed_mode` to `FIXED`.
  A different request, so the fingerprint changes and the prior result becomes
  STALE — correctly, because a different question is being asked. Re-running then
  replays the identical streams.
- **A failed AUTO run records its effective seed and nonce in the attempt
  metadata**, so the failure is reproducible.

### 5.5 Scalar seed to six-word state — D6-05

Revision 2 proposed a bespoke modular Lehmer expander and asserted that nearby
seeds "must decorrelate". **Both are withdrawn.** The second is not a contract
unless it is given a quantitative statistical statement, and inventing a scalar
mixer to carry statistical authority that belongs to the generator is the wrong
place to put it.

**Candidate A — the canonical repeated scalar** (now preferred):

```
s10 = s11 = s12 = s20 = s21 = s22 = seed
```

For `seed ∈ 1 … 2147483646` this is a valid state: every word lies inside both
moduli (`2147483646 < m2 < m1`), and neither component is all zero. It is what
L'Ecuyer's own examples do (`12345` repeated), so a PCCM run with seed `12345`
lands on a stream a reviewer can check against published material. **No new
algorithm, no new portability surface, no new vectors to invent.**

**Candidate B — modular expansion**, retained only as a fallback if Step-0
evidence shows a real defect in A across the admissible domain.

**D6-05 is OPEN**, with **A preferred**. The statistical-quality authority is
MRG32k3a itself, not a scalar mixer.

### 5.6 Component streams — the architecture

Revision 2 said "keyed on Permanent ID", which is not a contract. This is.

**Each simulation component owns one deterministic MRG32k3a stream:**

| Component | Count at target | Purpose |
|---|---|---|
| `COST_SAMPLE` | 200 | unit-cost sampling for one Cost Line |
| `RISK_OCCURRENCE` | 100 | the Bernoulli draw for one Risk |
| `RISK_SEVERITY` | 100 | severity sampling for one Risk |
| **Total** | **400** | trivial against MRG32k3a's stream space |

**Terminology, used precisely from here on.** A **stream** is a state separated
from its neighbour by the canonical `2^127` jump. A **substream** is the `2^76`
RngStreams concept. **Phase 6 proposes streams only and does not use substreams**
— revision 2's loose use of "substream" is corrected throughout.

**Why occurrence and severity are separate streams.** If one Risk drew its
Bernoulli and its Beta severity from a single stream, the variable number of
severity draws in iteration `i` would move where the Bernoulli sits in iteration
`i+1`. Changing a severity distribution — or merely improving the sampler — would
then change that Risk's *occurrence path* although `Probability` never changed.
That is an unacceptable coupling in a model whose whole point is attribution.
Separating them makes:

- Beta consumption unable to perturb any other driver, **and** unable to perturb
  its own Risk's occurrence path
- row-order invariance provable per component
- Phase-7 replay a per-component operation
- occurrence and severity independently testable in the oracle

### 5.7 Stream assignment — D6-16

Requirements: deterministic · collision-free over every representable Permanent
ID · independent of row position · derivable from resolved model data alone ·
portable to Python · unaffected by locale or collation.

**Family A — canonical sorted order → sequential streams** *(recommended)*

Components are ordered by `(ComponentKind, PermanentId)` with Permanent ID
compared **ordinally on UTF-16 code units** — the same comparison the accepted
Phase-5 fingerprint already locks, so no new collation authority is created.
Stream `k` is the base state advanced by `k` applications of the `2^127` jump.

- **Row-order invariance: guaranteed**, because the order is a sort, not a
  position.
- **Cost:** inserting or deleting a driver shifts every later component to a
  different stream, so an unrelated driver's samples change. Defensible — adding
  a driver changes the model, so the run is a new run — but it must be *stated*,
  because a user may expect an unrelated driver's numbers to hold still.

**Family B — direct ID-derived stream index**

The index is computed from `ComponentKind` plus the numeric part of the Permanent
ID (`CL-001` → 1, `R-001` → 1), e.g.
`index = kind_offset + numeric_id`.

- **Preserves stream identity across insertion and deletion** — a driver keeps
  its stream for the life of the model.
- **Cost:** requires a precisely bounded mapping and a guaranteed-collision-free
  numeric range, and it inherits whatever the Permanent-ID format guarantees. If
  IDs are ever reissued or the format widens, the bound must widen with it.

**Recommendation: A**, on the grounds that it introduces no new authority over
Permanent-ID structure. **B is the better user experience** and should be
reconsidered in Step 0 if the ID format's numeric bound can be stated as a
contract. **D6-16 is OPEN.**

### 5.8 No arithmetic seek to iteration *i*

**Revision 2 claimed each stream could be started "at a computable offset" for
iteration `i`. That is false under a variable-draw sampler** — the number of
advances before iteration `i` depends on every acceptance decision that came
before it, so it is not an arithmetic function of `i`. It would only be true if
each iteration got its own fixed substream, which would mean `400 × 100,000`
jumps and is not proposed.

**The design instead:** each component's stream advances **sequentially** across
iterations, and replay means *resetting the stream to its initial state and
re-running iterations `1 … i`*. That is deterministic and sufficient. **Direct
random access to an arbitrary iteration is not a Phase-6 requirement**, and §12 no
longer claims it.

### 5.9 Jump-ahead is a second, harder numerical contract — D6-04

The base recurrence keeps its products under `2^53` (§5.2.1). **The jump matrices
do not, and by a wide margin.** Matrix entries and state residues are both of
order `4 × 10^9`:

```
naive term   ≈ 4294967086 × 4294967086 = 18,446,742,269,823,331,396
2^53                                   =      9,007,199,254,740,992
overflow                               = 2048×
```

**A naive `sum(A(i,j)·state(j)) Mod m` in `Double` is therefore wrong**, and would
be wrong silently — it would produce plausible states that are not the canonical
ones. RngStreams uses dedicated modular-multiplication machinery for exactly this
reason.

**The decomposition (L'Ecuyer `MultModM`, `H = 2^17`)** splits `a = a1·H + a0` and
reduces in halves. Verified worst-case terms for `a = s = m1 − 1`:

| Term | Value | Fraction of `2^53` |
|---|---|---|
| `a1 · s` | `140,733,186,506,962` | 0.016 |
| `a0 · s` | `562,047,982,808,132` | 0.062 |
| `H · (a1·s mod m)` | `562,945,631,191,040` | 0.062 |

Every term is under 7% of `2^53` — comfortably exact, where the naive form was
2048× over.

**Step 0 must lock:** the canonical `A1p127` / `A2p127` matrices (`A1p76`/`A2p76`
only if substreams are ever actually used, which §5.6 does not propose) · the
modular matrix-vector algorithm · the `MultModM`-equivalent primitive · and
cross-language jump vectors.

**Gate-A jump vectors required:** base seed state · stream 0 initial state ·
stream 1 initial state · several later stream initial states (including one
beyond 400, to prove the recurrence rather than a table) · the Python
exact-integer result · the VBA-safe-arithmetic expected result.

**D6-04 cannot close until the jump arithmetic is part of the feasibility
result.**

### 5.10 Versioning

`RNG_VERSION` and `SIM_METHOD_VERSION` are separate integers stored with the run.
The first changes if the generator, the seeding or the jump/assignment scheme
changes; the second if a sampler or an accumulation order changes. Either
invalidates a stored result exactly as `FP_VERSION` already does.

---
## 6. Iterations

| Concept | Value | Kind | Source |
|---|---|---|---|
| Hard refusal | `< 1000`, or non-whole | **business rule** | `input_contract.yaml`, LOCKED |
| Default | `10000` | business rule | LOCKED |
| `< 10000` advisory | **Model Check**, not Phase 6 | advisory | LOCKED |
| Tested performance target | `100000` | benchmark | `phase5_plan.md` §26 — target, not cap |
| **Business-rule maximum** | **NONE** | — | "No upper limit is imposed" |
| **Simulation representation ceiling** | §6.1 | **technical** | proposed |

### 6.1 The technical storage ceiling is computed, not awaited

Once `_SimData` persists one row per iteration (§11.1), Excel's worksheet
capacity is a deterministic limit reached long before memory allocation fails.
Discovering it as a COM error mid-publish — after the whole run has been computed
— would be indefensible.

```
max_iterations_representable = 1048576 − H          ' H = _SimData header rows
```

**The refusal happens before anything is consumed:** before sample allocation,
before any stream is built, before a single uniform is drawn. The message names
the *technical storage limit* and the computed capacity, never a business rule.

**D6-08 is narrowed but OPEN**: `H` follows from the `_SimData` layout settled in
Step 1. The *rule* is proposed as locked; only the constant is outstanding.

**A contradiction, recorded.** The input contract imposes no upper limit while
row-per-iteration persistence imposes one. They reconcile only by keeping the two
in different categories — business versus technical — and never letting the
technical one be presented as validation. If review prefers no technical ceiling,
the consequence is summary-only persistence and the loss of everything §11.1
lists.

---

## 7. Selected Confidence Level — LOCKED as a reporting selector

Selected Confidence Level does **not** change the generated distribution, does
**not** belong in the request fingerprint, does **not** make a successful
simulation STALE or INVALID, and does **not** require a re-run.

Revision 1's `UNSELECTED` fourth state is **rejected**. The simulation state
remains exactly **CURRENT / STALE / INVALID**.

Phase 6 stores the whole ladder a selection could ask for — the ten selectable
levels of `lstConfidenceLevels` (`P50 … P95`) plus the fixed headline `P10` — so
Selected Px is a **deterministic lookup**, not a computation over samples.
Preferred implementation: a Results-side formula keyed on `C18`.

**Corrected wording.** Revision 2 said changing `C18` involves "no VBA event, no
recalculation and no re-run". A worksheet formula *does* recalculate under normal
Excel calculation. What is meant, precisely:

> **no VBA event, no simulation execution, and no new random draws.**
> Ordinary worksheet formula recalculation is expected and allowed.

**P10 is a fixed reported statistic**, not selectable. `lstConfidenceLevels`
remains `P50…P95`; no accepted input contract is reopened.

**D6-06 CLOSED. D6-10 CLOSED by rejection.**

---

## 8. Statistics

### 8.1 Moments

- **Sample mean**: `Σx / n`, accumulated scale-safely (§4.6).
- **Standard deviation — LOCKED: the sample standard deviation, divisor
  `n − 1`.** A **reporting-method decision taken here**, not inherited from Excel
  or any library: the run is a sample from the model's distribution, not the
  population. **D6-09 CLOSED.**
- Both computed with the accepted Phase-5 safe primitives and the scale-safe
  strategies of §4.6 — **not** naive `Σx²`, and **not** unguarded Welford.

### 8.2 Percentile algorithm — LOCKED

**Linear-interpolation percentile, Hyndman–Fan Type 7**, equivalent to Excel's
`PERCENTILE.INC` and NumPy's `linear` method. *Nearest-rank is a different
algorithm and the name is not used here.*

```
sorted x[0..n-1] ascending
h  = (n − 1) · p                      ' p ∈ [0,1]
lo = floor(h) ; hi = min(lo + 1, n − 1) ; f = h − lo
Px = (1 − f)·x[lo] + f·x[hi]          ' convex form, §4.6
```

`n = 1` returns `x[0]`; `p = 0` the minimum; `p = 1` the maximum. The convex form
is used rather than `x_lo + f·(x_hi − x_lo)` because the difference can overflow
where the result cannot.

**No `WorksheetFunction.Percentile`** — already forbidden, and a worksheet
function in the kernel would break the purity sweep. The equation is the
authority; no library default and no name is inherited.

### 8.3 The reported ladder — LOCKED

Stored for **nominal** and **PV** separately: fixed headline `P10, P50, P70, P90`;
the full selectable ladder `P50 … P95`; mean and sample standard deviation;
minimum and maximum. Eleven distinct percentiles per measure.

### 8.4 Deferred

**Annual simulated distributions are deferred to Phase 7.** Phase 6 reports the
analytical annual series Phase 5 already produces and retains no per-year samples.

### 8.5 Hand-derived vectors before code

Hand-derived percentile vectors for `n = 1, 2, 3, 4, 10` at every reported `p`,
including exactly-integral `h` and interpolated `h`.

---

## 9. Contingency

| # | Quantity | Definition |
|---|---|---|
| 1 | Escalated Deterministic Base Estimate | `A` — Phase 5, accepted |
| 2 | Expected Risk / EMV | `D` — Phase 5, accepted |
| 3 | Analytical Expected Total | `C + D` — Phase 5, accepted |
| 4 | Simulation Mean Total | `Σx/n` |
| 5 | Selected Px Total | `Px` from §8.2 |
| 6 | **Contingency at the selected confidence level** | **`Selected Px Total − A`** |

The baseline is the **deterministic base estimate**, not the simulation mean and
not the analytical expected total. Nominal and PV separately, using
`SafeSubtract`. **D6-07 CLOSED.**

**The workbook reports at the confidence level the user selected. It does not
recommend one.**

---

## 10. Fingerprints and state

### 10.1 One canonical stream, two digests over a prefix and its extension

Revision 2 left this ambiguous. Specified now, and it avoids hash-of-hash
duplication:

```
canonical field stream:   HEADER · COST · RISK · SIM
                          └──── analytical prefix ────┘
                          └──────── request extension ────────┘

analytical_fingerprint = digest( HEADER · COST · RISK )
request_fingerprint    = digest( HEADER · COST · RISK · SIM )
```

**The analytical digest is not inserted into `SIM` as a field.** The two digests
are taken over a prefix and over that same prefix plus an extension, using the
identical accepted hash (`FP_BASE 131`, the two moduli, `FP_INIT 1`). Because the
prefix bytes are shared, a simulation whose `request_fingerprint` is current is
*by construction* describing the same model prefix Phase 5 evaluated — the proof
is structural, not a comparison of two stored strings.

`FP_VERSION` continues to govern the algorithm; `HEADER/COST/RISK` keep their
positions and bytes, so the Phase-5 subset stays comparable across phases.

### 10.2 What `SIM` contains

| Field | In? | Why |
|---|---|---|
| Iteration count | **yes** | changes the distribution |
| `seed_mode` (`FIXED`/`AUTO`) | **yes** | canonical input state; a blank cell *is* a value |
| `supplied_seed` | **yes, FIXED only** | it is an input then, and only then |
| `RNG_VERSION` | **yes** | changes the numbers for identical inputs |
| `SIM_METHOD_VERSION` | **yes** | same |
| Effective seed | **no** | an *outcome*; including it made the AUTO fingerprint unrecomputable |
| `auto_nonce` | **no** | audit state, not model input |
| `run_id` | **no** | audit sequencing |
| Selected Confidence Level | **no** | §7 |

### 10.3 The three states

| State | Meaning |
|---|---|
| **CURRENT** | the recomputed `request_fingerprint` matches the stored one, and the last attempt succeeded |
| **STALE** | a stored successful result exists, but `request_fingerprint` no longer matches |
| **INVALID** | prerequisites refuse (including §10.4), or the last attempt failed |

**There is no fourth state.**

### 10.4 The Phase-5 analytical prerequisite — D6-14

Without a rule, Phase 6 could publish a CURRENT simulation from one input
snapshot while the visible Phase-5 outputs came from another — two published
result sets describing different models, both marked current.

| | Design |
|---|---|
| **A** | `PCCM_RunSimulation` **requires** Phase-5 status `CURRENT`; refuses otherwise |
| **B** | it silently refreshes the Phase-5 calculation first |
| **C** | one transactional operation publishing both layers |

**Recommendation: A.** **B** would let a Phase-6 endpoint silently mutate
accepted Phase-5 publication semantics — the boundary violation the module split
exists to prevent. **C** doubles an already transactional surface for a
convenience belonging to the presentation phase.

Under **A**: STALE or INVALID analytical state ⇒ the simulation **refuses**, the
prior successful simulation is untouched, and the detail names the analytical
state. The proof that both layers describe one model is §10.1's shared prefix.
**D6-14 OPEN**, A recommended.

### 10.5 What survives a refusal

A refused simulation **preserves the prior successful results** and leaves them
visible with their own stamp, recording the failure in the attempt fields. A
failed run never publishes a partial distribution (§16).

---

## 11. Simulation identities

Revision 2 had two identities and needed five. **"Simulation fingerprint" is no
longer used to mean both a request and a result.**

| # | Identity | Over what | Changes when |
|---|---|---|---|
| 1 | `analytical_fingerprint` | `HEADER·COST·RISK` | the model changes |
| 2 | `request_fingerprint` | `HEADER·COST·RISK·SIM` | the model, iterations, seed mode, supplied seed or a version changes |
| 3 | `run_id` | — | every successful run (§11.2) |
| 4 | `effective_seed` | — | every AUTO run; fixed by input in FIXED mode |
| 5 | **`result_digest`** | the canonical iteration-ordered outputs | any sampled value changes |

**Why identity 5 is required.** Under AUTO, the same model, iterations, seed mode
and versions legitimately produce **many successful runs sharing one
`request_fingerprint`** but with different effective seeds and different samples.
A request fingerprint is therefore *not* a run identity, and nothing in
revision 2 could distinguish two such runs. For an AUTO re-run:

```
request_fingerprint   unchanged        ← the same question was asked
run_id                changes
effective_seed        changes
result_digest         changes          ← a different answer came back
```

That is correct, not contradictory.

**`result_digest` definition — D6-17.** The accepted Phase-5 canonical encoder
and hash, applied to the iteration-ordered outputs: for `i = 1 … n`, the
canonical `Double` encoding of `total_nominal[i]` then `total_pv[i]`, with the
iteration index and `n` in the stream so a truncated run cannot collide with a
short one. The exact field order and tagging is Step-0 work. **This is what G2
and G3 compare** (§15).

### 11.1 `_SimData` persists iteration order — LOCKED

Sorting destroys iteration identity, and iteration identity is what makes a run
auditable — it costs replay, the result digest, Phase-7 sensitivity pairing, and
diagnosis of any single iteration. `_SimData` stores, in **canonical iteration
order**:

| Column | Meaning |
|---|---|
| `iteration_index` | `1 … n`, the identity |
| `total_nominal` | SAR nominal |
| `total_pv` | SAR present value |

plus the run-identity block of §11. **Statistics sort copies in memory; the
persisted arrays are never permuted.** **D6-13 CLOSED by adoption.**

### 11.2 Run ID and the AUTO nonce — D6-15, jointly with D6-03

Revision 2 recommended deriving the AUTO seed from `run_id` **and** allocating
`run_id` only on success. **Those conflict**: repeated failed AUTO attempts would
re-derive the same candidate seed, so "blank means a new random sequence" would
be false exactly when a user is retrying after a failure.

Resolved by separating two counters:

| Counter | Advances | Purpose |
|---|---|---|
| **`run_id`** — successful-run sequence | only when a result **commits** | human/audit sequencing |
| **`auto_nonce`** — attempt sequence | whenever an AUTO effective seed is **allocated**, including attempts that later fail | guarantees a fresh seed by construction |

The `auto_nonce` is **audit state, not model input**: it is not in the request
fingerprint, it is recorded in attempt metadata, and a failure leaves the prior
successful simulation untouched.

**The mapping `auto_nonce → effective_seed ∈ 1 … 2147483646`** must be
deterministic and collision-free over the representable nonce range, or carry an
explicit collision-resolution rule. A full-period multiplicative cycle over the
same modulus is the leading candidate — every nonce in one period yields a
distinct seed, so accidental immediate reuse is impossible *by construction*
rather than by comparison. **D6-03 OPEN**, jointly with this.

`run_id` semantics: allocated at commit · failed attempts consume none · unique
within the workbook · monotonic · persisted, so it survives Save/Reopen ·
independent of `effective_seed` · **no GUID, no COM identity source, no
`Now()`-derived uniqueness**. **D6-15 OPEN**, as above.

---
## 12. Results ownership and the Phase-7 boundary

| Layer | Phase | Note |
|---|---|---|
| **A. simulation engine** | **6** | |
| **B. persisted results** | **6** | `_SimData`, iteration-ordered (§11.1) |
| **C. Results presentation** | **6, minimal** — Run Stamp + Summary Statistics | Annual Cash Flow and Reconciliation deferred to Phase 8 |
| D. Dashboard | 8 | must read from Results, never recompute |
| E. Sensitivity / Spearman | 7 | below |
| F. Charts | 8 | |

**No duplicate source of truth.** Results derives from `_SimData`; Dashboard
derives from Results; nothing recomputes a Monte Carlo independently.

### 12.1 Sensitivity — Phase 7

Spearman rank correlation of a driver's sampled contribution against the total
requires **per-driver, per-iteration** values: `300 × 100,000 × 8 B = 240 MB` at
the design target, against a `< 2 MB` engine. **Online computation is not
available for Spearman** — ranking is inherently a whole-sample operation, unlike
a mean.

**Recommendation: defer to Phase 7**, and get the distribution accepted first.

**How Phase 7 replays without Phase 6 retaining anything.** Revision 2 claimed a
"direct seek" to iteration `i`; **that claim is withdrawn** (§5.8). What the
component-stream architecture actually provides is better than nothing and weaker
than random access: Phase 7 can **reset any component's stream to its initial
state and replay iterations `1 … i` deterministically**, in isolation, without
disturbing or needing any other component. That is sufficient for sensitivity —
which needs whole columns, not arbitrary single cells — and it is a direct
consequence of §5.6's separation.

---

## 13. Correlation between drivers — OUT OF SCOPE

The repository contains **no** inter-driver correlation authority; the only
occurrences of the word are the Sensitivity sheet's *output* ranking, which is a
different thing — measuring how a driver co-varies with the total is not imposing
dependence between inputs.

**Drivers are sampled independently.** Correlated uncertainty is explicitly out
of scope and is a stated limitation on the Methodology sheet. No correlation
matrix, no copula, no hidden dependency; introducing one needs a new input
contract and a new authority, neither of which exists. **No conflicting
repository text was found.**

The component-stream architecture reinforces this: independent streams per
component make accidental coupling structurally difficult rather than merely
unintended.

---

## 14. Performance and memory model

**Operation counts are architecture-dependent and no longer assume fixed
consumption.** At the design target (200 Cost Lines, 100 Risks, 25 years,
100,000 iterations), worst case = all drivers Beta-PERT with `Probability = 1`:

| Quantity | Count | Note |
|---|---|---|
| Occurrence uniforms | `1.0 × 10^7` | **fixed** — one per risk per iteration |
| Cost-line sampling draws | `2.0 × 10^7` samples | **variable** uniform count per sample |
| Severity sampling draws | `≤ 1.0 × 10^7` samples | **conditional** on occurrence; worst case `p = 1` |
| Total Beta samples, worst case | `3.0 × 10^7` | |
| Uniforms consumed | **not a fixed number** | bounded by acceptance rate; the Step-0 package measures mean and high-percentile consumption per shape (§17) |
| Multiply-accumulates | `≤ 6.0 × 10^7` | nominal + PV on evaluated contributions |
| Sorts | 2 × `10^5` elements | `~3.4 × 10^6` comparisons |
| Stream jumps | `400`, **once**, before the loop | §5.9 |
| **Worksheet / COM calls inside the loop** | **0** | the locked invariant |

**No flop total is asserted here.** Revision 2 quoted `~3 × 10^9` for Option B;
that was a desk estimate, and per §17 no such number is an accepted input to
D6-04 unless Step 0 returns the measurement artefact and the counting method
alongside it.

**Memory, by category:**

| Category | Contents | Bytes |
|---|---|---|
| Hot-kernel resident | Phase-5 structures (`< 100 KB`) + two iteration-ordered sample arrays | `~1.7 MB` |
| **Component stream states** | `400 × 6 × 8 B` | **`19 KB`** |
| Statistics working | two sorted *copies* | `1.6 MB` |
| Publication buffer | `Variant` array, `100,000 × 3 × 16 B` | `4.8 MB` |
| Workbook persisted | 300,000 numeric cells, compressed | `~1–3 MB` on disk |
| **Peak resident** | hot + streams + statistics + publication | **`~8.2 MB`** |

Deferred retentions: annual samples 20 MB; per-driver samples 240 MB.

The publication buffer can be **chunked** — 20,000-row blocks reduce 4.8 MB to
under 1 MB for a handful of extra COM calls, all **after** the loop and therefore
outside the hot-path rule. A Step-10 implementation choice, not architectural.

---

## 15. Oracle and evidence — layered by what each layer can prove

**This is the section Option B changes most.** With an acceptance/rejection
sampler, a one-ULP difference in an acceptance test flips accept to reject,
changes how many uniforms that stream consumed, and **desynchronises every
subsequent draw on that stream**. After divergence the two implementations are no
longer sampling the same random path, so comparing sample `k` to sample `k`
compares unrelated numbers. "Python and VBA Beta outputs agree within tolerance"
is therefore *not* a meaningful statement about a full seeded simulation, and
revision 2's G4 was too strong in exactly that way.

The oracle is not weakened. What each layer *can* prove is stated honestly.

| Layer | Subject | Evidence | Strength |
|---|---|---|---|
| **A** | MRG32k3a backbone | canonical published state and uniform vectors | **EXACT** — justified by §5.2.1 |
| **B** | jump-ahead and stream assignment | initial state of streams 0, 1, and several later, Python exact-integer vs VBA-safe arithmetic | **EXACT** — justified by §5.9 |
| **C** | Uniform / Triangular transforms | **injected** uniforms, not generated ones; expected outputs | locked numeric comparison policy (§15.1) |
| **D** | Bernoulli decisions | locked uniforms × locked probabilities, including `p = 0`, `p = 1`, and `u` either side of `p` | **EXACT** — a comparison, not arithmetic |
| **E** | Cheng BB/BC | locked deterministic test streams; branch-covering vectors for BB and BC; **expected draw counts**; outputs under §15.1; independent theoretical mean/variance checks | mixed: draw counts exact, outputs tolerance-bounded, distribution independently verified |
| **F** | full seeded simulation, **no Beta drivers** | Python vs VBA end-to-end | **EXACT and strong** — no rejection path to desynchronise |
| **G** | full seeded simulation, **with Beta drivers** | **not** required to match sample-for-sample across languages | see below |

**Layer G is the honest one.** Cross-language sample-for-sample identity with a
rejection sampler is only defensible if Step-0 evidence shows the floating
acceptance path is stable enough to make it so. Until then, confidence in a full
Beta simulation comes from combining:

- **G1** RNG conformance (layer A) — exact
- **G2** same-runtime replay: identical `result_digest` in VBA — exact
- **G3** row-order invariance: identical `result_digest` in VBA — exact
- deterministic Cheng vector cases (layer E) — exact draw counts
- independent statistical checks: simulated moments and quantiles against
  independently computed Beta-PERT theory, at a stated sample size and tolerance
- layer F, which exercises the *whole engine* end to end with no rejection path

**The Python oracle stays genuinely independent.** Forcing bit-identical Beta
results by transcribing the VBA into Python would manufacture agreement and
destroy the only reason to keep two implementations.

### 15.1 The numeric comparison policy

| Subject | Policy |
|---|---|
| RNG state and uniforms | **exact equality** |
| Bernoulli decisions | **exact equality** |
| Jump-ahead states | **exact equality** |
| Cheng draw counts, for a locked injected stream | **exact equality** |
| Transformed floating samples (Uniform, Triangular, Beta) | **a locked tolerance** — ULP or relative, fixed in the contract |
| Individual hand vectors whose result is exactly representable | exact equality **case by case**, where stated |

**Revision 2 said "Uniform and Triangular: exact" as a universal cross-language
claim. That is withdrawn.** The RNG uniform can be exact; a transform involving
floating multiply, add and especially `Sqr` may not be bit-identical across
independent runtimes, and asserting it without proof would put an unprovable
claim in the contract. Exactness is asserted per case, where the arithmetic
justifies it — never universally.

### 15.2 Vector families

| # | Family |
|---|---|
| 1 | canonical MRG32k3a state after seeding, seeds `1`, `2`, `12345`, `2147483646` |
| 2 | first 20 uniforms per seed, to 17 significant digits |
| 3 | seed edge cases: `1`, `2147483646`, refusals `0`, `−1`, `1.5`, blank |
| 4 | jump-ahead: streams 0, 1, 7, 399, 400+ |
| 5 | stream assignment for a model whose rows are then reordered |
| 6 | Bernoulli boundaries `p = 0`, `p = 1`, `u` either side of `p` |
| 7 | Triangular inverse at the branch point `c` and either side; injected uniforms |
| 8 | Uniform inverse including `a = b` and negative supports |
| 9 | Cheng **BB** branch vectors, `min(α,β) > 1` |
| 10 | Cheng **BC** branch vectors, `min(α,β) = 1` — i.e. `m = a` and `m = b` |
| 11 | Cheng draw counts for locked injected streams |
| 12 | Beta-PERT simulated moments vs independent theory |
| 13 | degenerate `a = m = b` |
| 14 | one-driver simulation, 1000 iterations, `result_digest` |
| 15 | mixed 3 cost + 2 risk, **no Beta**, full Python-vs-VBA end-to-end |
| 16 | same seed → identical `result_digest` |
| 17 | changed seed → different `result_digest` |
| 18 | **row reorder → identical `result_digest`** |
| 19 | percentile vectors for `n = 1, 2, 3, 4, 10` |
| 20 | extreme-domain pipeline vectors, §4.6 cases 1–10 |

---

## 16. Error and interruption behaviour

| Situation | Behaviour |
|---|---|
| Phase-5 analytical state not CURRENT | refuse before anything (§10.4) |
| Iterations above the storage ceiling | refuse before allocation and before any stream is built (§6.1) |
| Invalid input before simulation | refuse; prior results untouched; `INVALID` |
| Numerical failure mid-run | abort; publish nothing; prior results untouched; detail names the iteration, component and stage |
| Unexpected VBA error | the accepted `modAppState` refusal path; no partial publish |
| User cancellation | **not supported in Phase 6** — see below |
| Excel state | restored exactly as the accepted Phase-4/5 lifecycle does |
| Partial results | structurally impossible: `_SimData` and Results are written **once, after** the loop, in one transactional publish, commit-last |
| Effective seed and nonce on failure | recorded in attempt metadata, so a failing AUTO run is reproducible |

> A failed run never publishes a partially completed distribution as a
> successful result.

**Cancellation — the rationale, corrected.** Revision 2 justified this with "a
100,000-iteration run completes in seconds". **That is withdrawn: no runtime has
been measured, and the variable-draw architecture makes an unproven duration
claim less defensible, not more.** The real reasons are architectural:
cancellation adds a second transactional exit path through a publish designed to
have one; Phase 6 should establish correctness and the benchmark first; and it can
be reconsidered only if measured Gate-B runtime justifies it. **D6-12 CLOSED**,
on that basis.

---

## 17. User-facing command surface

| Control | Phase | Why |
|---|---|---|
| `PCCM_RunSimulation` (`Application.Run`) | **6** | the automation entry point, testable by the Gate-B harness |
| Model Check button | 9 | Model Check owns it |
| Run Simulation button | 8 | belongs with the Dashboard command set |
| Refresh / report outputs | 8 | same |

Phase 6 adds **no button, no ribbon, no form, no `MsgBox`**.

---

## 18. Gate strategy

**Gate A — pure and static, on Linux.** RNG, jump, sampler and stream contracts
in `spec/sim_contract.yaml`; the Python oracle; every vector family of §15.2;
static VBA source checks including the extended purity sweep; generated
artefacts and Stage-A verification; mutation controls on every new detector.
**No Windows runtime claim of any kind.**

**Gate B — real Windows/Excel.** Whole-project compile through the accepted
P5-CMP gate; seeded simulation against Python expectations *at the layer each can
prove* (§15); same-runtime replay digest; **row-order invariance digest**;
different-seed divergence; the full 100,000-iteration run at the design target;
the pre-flight storage-ceiling refusal; failure containment; natural Excel
shutdown; clean COM release.

**A performance benchmark is not a semantic test.** The 100k run proves the
budget is met and nothing about correctness; Gate B passes only if the semantic
scenarios pass independently of it.

---

## 19. Implementation sequence

**Step 0 comes first.** A contract cannot be the authority for decisions taken
after it exists.

| # | Step | Files | New authority | Gate-A acceptance |
|---|---|---|---|---|
| **0** | **Authority closure**, including the **D6-04 feasibility package** of §20.1 | `docs/` only | the closed decisions | every open §20 item resolved and independently reviewed; **no contract file yet** |
| 1 | **`spec/sim_contract.yaml`** — RNG and jump constants, seed domain, AUTO nonce mapping, stream assignment, Cheng dispatch, percentile method, standard-deviation method, contingency, `SIM` fields, `result_digest` canonicalisation, `sim_state` schema, `_SimData` layout and therefore `H`, versions. Loader + validator, fail-loud | `spec/`, `builder/`, `tests/` | the sixth contract | validator rejects every malformed shape; mutation controls per rule |
| 2 | **RNG + jump reference in Python**, exact integers; families 1–5 | `builder/`, `tests/` | locked vectors | canonical published vectors reproduced |
| 3 | **Samplers in Python** — Uniform, Triangular, **Cheng BB/BC**, Bernoulli; families 6–13 | `builder/`, `tests/` | — | every boundary in §4 has a hand-derived expectation; BB and BC branches both covered |
| 4 | **Simulation oracle** — the engine, `result_digest`, percentiles, contingency; families 14–20 | `builder/`, `tests/` | — | golden cases written first |
| 5 | **Stage-A emission** — `build/phase6_cases.json`, `modSimContract.bas`; verification extended | `builder/`, `tests/` | — | Stage-A verifier green |
| 6 | **`modSimRng`** — recurrence, reduction, jump, stream assignment; purity sweep active from the first commit | `src/vba/`, `tests/` | — | families 1–5 pass against Python |
| 7 | **`modSimSample`** — the three samplers | `src/vba/`, `tests/` | — | families 6–13 pass under §15.1 |
| 8 | **`modSimEngine`** — the iteration loop and accumulation | `src/vba/`, `tests/` | — | zero worksheet references in the loop, proved statically |
| 9 | **`modSimStats`** — sort, moments, percentiles, contingency, all scale-safe | `src/vba/`, `tests/` | — | families 19–20 pass |
| 10 | **`modSimFingerprint`** — `SIM` appended after `RISK`; `result_digest` | `src/vba/`, `tests/` | — | Phase-5 digests unchanged for every accepted case |
| 11 | **`modSimReport`** — `_SimData`, the minimal Results block, transactional publish, `sim_state`, `PCCM_RunSimulation` | `src/vba/`, `tests/` | — | no partial-publish path exists |
| 12 | **Gate-A source review** | — | — | independent review |
| 13 | **Gate-B harness extension**, then Windows Run 1 | `bootstrap/`, `tests/` | — | §18 |

The Step-0 feasibility work is **throwaway measurement code** under `docs/` or a
scratch path. It ships nothing and must not become the implementation by default.

---

## 20. Open decisions

**Closed:** D6-01 (scope split) · D6-02 (canonical MRG32k3a, `(0,1)`) · D6-06
(P10 fixed) · D6-07 (contingency baseline `A`) · D6-09 (sample standard
deviation) · D6-10 (no `UNSELECTED`) · D6-11 (scoped guards) · D6-12 (no
cancellation) · D6-13 (raw iteration-ordered persistence).

| # | Decision | Options | Evidence | Recommended | Consequence | Close before |
|---|---|---|---|---|---|---|
| **D6-03** | AUTO-seed source and freshness, **jointly with D6-15** | (a) timestamp folded into the domain; (b) **`auto_nonce` counter → full-period multiplicative cycle**; (c) refuse blank | (c) contradicts "blank means a new random sequence"; (a) can collide | **(b)** — collision-freedom by construction, and it settles the failed-attempt case | whether "new" is guaranteed or merely likely | **Step 0** |
| **D6-04** | **Beta sampler + component-stream scheme + jump arithmetic, as ONE decision** | (a) one-uniform inverse CDF + fixed consumption; (b) **Cheng BB/BC + component streams + `2^127` jumps**; (c) precomputed table | (a) infeasible; (c) `5.9e-3` normalised error at 4096 nodes; (b) needs the §20.1 package | **(b)**, subject to §20.1 | the stream architecture, the jump work, the oracle's layer boundaries, Phase-7 replay | **Step 0** |
| **D6-05** | scalar seed → six-word state | (a) **repeated scalar**, canonical; (b) modular expansion | (a) is valid over the whole domain and matches published examples; no defect is known | **(a)** | one more portability surface, or none | **Step 0** |
| **D6-08** | technical storage ceiling constant | (a) `1048576 − H` from the final layout; (b) a lower stated ceiling | rule settled in §6.1; only `H` outstanding | **(a)** | the exact refusal threshold | **Step 1** |
| **D6-14** | analytical prerequisite | (a) require Phase-5 `CURRENT`; (b) refresh first; (c) one transaction | Phase-5 publication semantics are accepted | **(a)** | whether Phase 6 can mutate Phase-5 outputs | **Step 0** |
| **D6-15** | `run_id` semantics, **jointly with D6-03** | (a) **monotonic success counter, separate from `auto_nonce`**; (b) GUID; (c) timestamp | Results reserves a Run ID; nothing defines it | **(a)** | audit identity; no computational effect | **Step 0** |
| **D6-16** | stream assignment rule | (a) **canonical sorted order → sequential streams**; (b) direct ID-derived index | (a) needs no new authority; (b) preserves stream identity across insert/delete but needs a bounded ID contract | **(a)**, revisit (b) if the ID numeric bound can be contracted | whether an unrelated driver's samples move when a driver is added | **Step 0** |
| **D6-17** | `result_digest` canonicalisation | (a) accepted encoder over iteration-ordered totals with `n` and index in the stream; (b) a separate scheme | reuses the accepted Phase-5 hash | **(a)** | what G2/G3 actually compare | **Step 0** |

### 20.1 The D6-04 Step-0 feasibility package

D6-04 closes on **evidence, not prose**. Required deliverables:

1. **Cheng dispatch** implemented and demonstrated: BB when `min(α,β) > 1`, BC
   when `min(α,β) ≤ 1`.
2. **Parameter grid spanning the whole PERT family**, `α, β ∈ [1,5]`, including
   `1/5`, `5/1`, both near-1 boundaries, symmetric `3/3`, and interior shapes.
3. **Per shape:** acceptance rate · mean uniforms consumed · high-percentile
   uniforms consumed · numerical output checks · independent theoretical
   mean/variance checks.
4. **The component-stream architecture** demonstrated: Cost → one stream;
   Risk → occurrence + severity streams.
5. **The exact stream-assignment rule** (D6-16), demonstrated row-order invariant.
6. **The exact jump matrices and modular arithmetic design** (§5.9), with
   cross-language vectors.
7. **Memory for all stream states** at the design target.
8. **A worst-case operation model** — 200 Cost Lines, 100 Risks, all Beta-PERT,
   `p = 1`, 100,000 iterations — **with the measurement artefact and the counting
   method returned alongside it.** No desk estimate is an accepted input.

---
## 21. Stale revision-2 statements corrected

Every statement inherited from the fixed-consumption architecture, and every
other correction from review round 3:

| # | Where | Stale statement | Corrected to |
|---|---|---|---|
| 1 | §3.4 | one uniform per Cost Line, two per Risk, one inverse-CDF evaluation per driver | variable consumption per component; occurrence fixed at one, severity conditional and variable |
| 2 | §4 opening | "All sampling is inverse-CDF on a single uniform" | Uniform and Triangular are; **Beta-PERT is not** |
| 3 | §4.4 | `unitCost = sample(...)` described as "ONE draw" | one uniform only for Uniform/Triangular; variable for Beta-PERT |
| 4 | §4.5 | "every risk consumes exactly two uniforms" alongside a conditional transform | occurrence stream advances by one unconditionally; severity stream advances only on occurrence, by a variable amount — **two streams, not two draws** |
| 5 | §5.6 | "substreams" used loosely | **stream** (`2^127`) and **substream** (`2^76`) distinguished; Phase 6 proposes streams only |
| 6 | §5.6 | "iteration `i` starts at a computable offset" | **withdrawn** — false under variable draws; replay is stream-reset plus re-run (§5.8) |
| 7 | §5.6 | "keyed on Permanent ID" | two named assignment families with their trade-offs; D6-16 |
| 8 | §5.7 (old) | "Uniform and Triangular: exact" cross-language | **withdrawn**; per-case exactness only, under the §15.1 policy |
| 9 | §5.9 (new) | jump arithmetic not mentioned at all | naive form overflows `2^53` by **2048×**; `MultModM` decomposition holds every term under 0.062 |
| 10 | §3.3 | `SimulationRequest` carrying `EffectiveSeed` / `SeedWasSupplied` | split into `SimulationRequest` (`SeedMode`, `SuppliedSeed`) and `SimulationRunContext` (`EffectiveSeed`, `RunId`, `AutoNonce`, digests) |
| 11 | §5.5 | "splitmix-style"; "nearby seeds must decorrelate" | repeated-scalar canonical mapping preferred; the unquantified decorrelation claim **withdrawn** |
| 12 | §10.1 | how the two digests relate left ambiguous | prefix and prefix+extension over one canonical stream; **no hash-of-hash** |
| 13 | §11 | two identities where five were needed | `analytical_fingerprint`, `request_fingerprint`, `run_id`, `effective_seed`, **`result_digest`** |
| 14 | §11 table | "`_SimData` holds the sorted sample arrays" | iteration-ordered, contradicting §11.1; **fixed** |
| 15 | §11.2 | `run_id` derived AUTO seed **and** allocated only on success | `auto_nonce` (advances on allocation, including failures) separated from `run_id` (advances on commit) |
| 16 | §7 | "no VBA event, no recalculation and no re-run" | "no VBA event, no simulation execution, no new random draws"; **worksheet recalculation is expected and allowed** |
| 17 | §14 | fixed 40-million-uniform count, then an Option-B caveat | counts stated per component kind; **uniforms consumed is explicitly not a fixed number** |
| 18 | §14 | `~3 × 10^9` flops presented as a figure | **withdrawn** — no flop total is asserted without the Step-0 artefact |
| 19 | §15 | "Beta-PERT inverse at fixed `u` values" | Cheng BB/BC branch vectors and draw counts; there is no inverse to test |
| 20 | §15 | cross-language tolerance treated as sufficient for a full seeded run | **layered evidence** — divergence in an acceptance test desynchronises the stream, so layer G is explicitly not required to match sample-for-sample |
| 21 | §16 | "completes in seconds" | **withdrawn**; the rationale is architectural, and no runtime is claimed before Gate B |
| 22 | §19 step 3 | "Beta-PERT inverse" | Cheng BB/BC |
| 23 | §12 | "fixed consumption order … independent substreams are not needed" | **deleted**; replaced by stream-reset replay |
| 24 | §4.6 | numerics covered the sampler only; "a numerically stable pass" for statistics | extended to contribution, accumulation, mean, standard deviation (**including the Welford overflow**), percentile interpolation and contingency; vector cases 9–10 added |
| 25 | §4.3 | Cheng described as "exact" | **distribution-exact acceptance/rejection, subject to floating-point implementation error**; dispatch boundary stated |

## 22. Remaining unresolved blocker

**One, and it is D6-04.** The architecture in this document is coherent and
internally consistent, but it rests on a sampler whose cost and floating-point
stability have not been measured. Specifically:

- if the Step-0 package shows Cheng BB/BC acceptance rates or consumption are
  worse than the bounded-density argument suggests, the performance case weakens
  and options A and C — both already rejected on measured grounds — would have to
  be revisited, which would reopen the stream architecture with them;
- if the floating acceptance path proves *stable* across implementations, layer G
  of §15 could be strengthened from statistical to exact, which would materially
  improve the evidence model;
- the jump arithmetic must be demonstrated exact against canonical vectors before
  any of it can be locked.

Everything else is either closed or closes on a decision that does not depend on
measurement.

---

## 23. What this plan does NOT claim

1. **No Phase-6 code exists.** No VBA, no builder change, no contract, no test,
   no simulation or oracle code in any language.
2. **No Windows or Excel runtime has been executed for Phase 6.**
3. **`spec/sim_contract.yaml` does not exist** and may not be created until
   Step 0 closes every decision it must encode.
4. **No runtime duration is claimed.** No benchmark has been run.
5. **No flop total is asserted** without its measurement artefact.
6. **Nothing here is accepted.** Eight decisions remain open, and D6-04 could
   still change the architecture this revision describes.
7. **The Phase-5 baseline is untouched.** `f571154` remains the accepted
   executable baseline; this document changes no file under `src/`, `spec/`,
   `builder/`, `bootstrap/` or `tests/`.
