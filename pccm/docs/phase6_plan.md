# Phase 6 — Stochastic simulation layer

**Status: PLANNING ONLY — revision 6, after review round 6.
NOT ACCEPTED. NO IMPLEMENTATION EXISTS. NO WINDOWS/EXCEL RUNTIME HAS BEEN
EXECUTED. `spec/sim_contract.yaml` DOES NOT EXIST.**

Phase 5 is closed. The accepted executable baseline is
`f571154118083e569e1fb9fbf9bf72852cc2d568`; the closure head is `28fa613`. This
document proposes what Phase 6 should be, from the authorities already in the
repository. It changes no code, no contract and no generated artefact.

**This document describes ONE candidate architecture throughout.** Revision 2
recommended a variable-draw Beta sampler while several sections still described
the one-uniform fixed-consumption design it replaced. Every such statement has
been rewritten; the stale-statement ledger is §21.

### How to read the status labels in this document

Revision 3 used **LOCKED** for two different things: authority inherited from an
accepted contract, and a decision this plan itself proposes. Those are not the
same standing, and the document now distinguishes three:

| Label | Meaning |
|---|---|
| **INHERITED** | Locked by an accepted prior authority — a contract, the accepted plan, or an accepted runtime result. Phase 6 may not reopen it. Everything so labelled in §1 is of this kind |
| **PROPOSED** | A Phase-6 design choice put forward here. It has no authority at all until this plan is accepted |
| **SETTLED-IN-PLAN** | A question this revision closes internally, so the rest of the document can be consistent. **Still only a proposal**; it is not authoritative until the plan is accepted, and it is not a contract until Step 1 encodes it |

**Nothing invented in this document is authoritative.** The plan is NOT ACCEPTED,
so every `SETTLED-IN-PLAN` item remains a recommendation to independent review.

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
Phase 6. **Every classification in this section is INHERITED authority** in the
sense of the legend above — these are statements already accepted, not Phase-6
proposals. Within that, the sub-classifications mean: **LOCKED** — an accepted
authority decides it and Phase 6 may not reopen it; **inherited invariant** — an
accepted design property Phase 6 must preserve; **placeholder only** — a reserved
surface with no committed semantics; **unresolved** — an authority explicitly
deferred the decision to this phase. The word LOCKED in this section always
refers to prior accepted authority, never to anything this plan invents.

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
| `Rnd` / `Randomize` forbidden | `Rnd(`, `Randomize`, `MRG32k3a`, `NPV`, `Percentile`, `RunSimulation` in `forbidden_constructs` | `structure_contract.yaml` | **LOCKED for Phase 4.** How Phase 6 narrows the list without deleting protections is decided in **§10.4a (D6-11)**, which also records that the extension reopens this accepted authority |

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
- **Uniform** sampling by affine transform of one uniform
- **Triangular** sampling by inverse CDF of one uniform
- **Beta-PERT** sampling by Cheng BB/BC acceptance/rejection — **not** inverse
  CDF, and a variable number of uniforms
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
| **per iteration, per Risk** | one uniform from its **occurrence** stream, **always**; then per D6-18 — **Option A:** invoke the severity sampler only if the risk occurred. **Option B:** invoke it every iteration and use the value only if the risk occurred. Two multiply-accumulates when it occurred |
| **per iteration** | write the two accumulated totals into the sample arrays at index `i` |
| **after the loop** | copy each sample array and sort the copies; compute moments and percentiles; compute contingency; compute the result digest; compute the request fingerprint; write `_SimData` in iteration order and Results in one transactional publish |

**Consumption is not constant and the plan no longer pretends it is.** Only the
*occurrence* draw is one-per-risk-per-iteration; everything else varies with the
sampler's acceptance path.

### 3.5 Canonical evaluation and accumulation order — SETTLED-IN-PLAN

**Stream assignment invariance is necessary but not sufficient.** Revision 3
made stream assignment row-order invariant and stopped there. Floating-point
addition is **not associative**: if contributions are accumulated in resolved-row
order, physically reordering two Cost Lines changes the order of the additions,
changes the rounding, and therefore changes the iteration totals and the
`result_digest` — **even though every driver received exactly the same stream and
exactly the same sampled value**. Revision 3's G3 would have been false for a
reason it never mentioned.

**The rule:**

> Simulation driver evaluation and accumulation use **canonical Permanent-ID
> order**, independent of physical worksheet row order.

Stated exactly:

| Question | Rule |
|---|---|
| Cost Line order | ascending Permanent ID, ordinal UTF-16 comparison |
| Risk order | ascending Permanent ID, ordinal UTF-16 comparison |
| Order between kinds | **all Cost Lines, then all Risks** — a fixed kind precedence, so a model gaining its first Risk does not reshuffle cost-line arithmetic |
| Do the occurrence and severity components affect evaluation order? | **No.** They are RNG components, not accumulation terms. A Risk contributes exactly one nominal and one PV term, at the position its Permanent ID gives it |
| Nominal and PV accumulation | **two independent accumulators**, each accumulating its own term in the same canonical driver order. Neither is derived from the other |

This is the identical ordering the Phase-5 fingerprint already locks for driver
records, so no new collation authority is created.

**Gate-A control required, on a rounding-sensitive fixture.** A mutation that
**keeps every stream identity fixed** and reverses only the accumulation order
must be caught.

**The difference is not universal and the test must not assume it is.** Many
legal contribution sets sum to the same binary64 value in either order — small
integers, values sharing an exponent, anything that happens not to round. The
control therefore uses a **deliberately constructed non-associative fixture**:

| Requirement | |
|---|---|
| 1 | contributions chosen so that canonical order and reversed order give **different binary64 totals**, proved by the independent oracle before the control is written |
| 2 | every RNG stream and every sampled driver value held **identical** between the two orders |
| 3 | the mutation detector must catch the reversed order on that fixture |

Vector family 21 states that the difference is required **only** for this
constructed fixture.

**Row-order invariance itself remains universal**, and is a different claim:

> a physical worksheet reorder leaves the canonical execution order unchanged,
> and therefore produces an identical same-runtime `result_digest`

— which holds for every model, rounding-sensitive or not, because the additions
happen in the same order regardless of where the rows sit.

**G3 and the Gate-B row-order test must therefore prove both:** that stream
assignment is invariant, and that accumulation order is invariant. One without
the other is not row-order invariance.

---

## 4. Sampling semantics

Let `a = Min`, `m = MostLikely`, `b = Max`, with the Phase-5 ordering guarantees
already established. **No positivity is assumed**; every formula below is valid
for negative `a`, `m`, `b`.

Uniform and Triangular are **inverse-CDF on a single uniform**. Beta-PERT is
**not** — it uses an acceptance/rejection method and consumes a variable number
(§4.3). Revision 2's blanket statement that "all sampling is inverse-CDF on a
single uniform" was false under its own recommendation and is withdrawn.

### 4.0 The degenerate case, decided before parameterisation — SETTLED-IN-PLAN

`a = m = b` is legal under the accepted ordering rule, and revision 3 handled it
inconsistently: Triangular said one uniform was consumed, Uniform said only that
it returns `a`, and Beta-PERT would have computed `r = (m − a)/(b − a) = 0/0`
before anyone noticed. **This is an RNG and replay contract, not an
implementation detail** — the same model and seed must advance each component
stream the same way, always.

**The rule:**

> A degenerate driver is detected **before** sampler dispatch and before any
> parameterisation. It returns `a`, enters no sampler, and **consumes zero
> uniforms**.

> **POST-ACCEPTANCE AUTHORITY CORRECTION — Phase-6 Step-1 review.**
>
> The *outcome* above is unchanged. The **detection predicate** was wrong for
> Uniform, and the error was inherited from writing one condition for three
> families.
>
> Revision 6 used `a = m = b` for every family. **Accepted Phase-5 D1 states
> that Uniform's Most Likely is ignored numerically and excluded from the
> calculation fingerprint.** A legal Uniform may therefore have `Min = Max` with
> Most Likely blank, or with Most Likely populated and unrelated — and under the
> common predicate the second case was **not** degenerate. It would enter the
> sampler and consume a uniform.
>
> That let an input the model explicitly ignores change **RNG consumption**, and
> therefore the stream position and every subsequent draw on that component.
> An ignored input must not be able to do that.
>
> **Detection is family-specific:**
>
> | Family | Degenerate when |
> |---|---|
> | **Uniform** | `a = b` — Most Likely is not read |
> | **Triangular** | `a = m = b` |
> | **Beta-PERT** | `a = m = b` |
>
> For Triangular and Beta-PERT the accepted ordering `a ≤ m ≤ b` already makes
> `a = b` imply `m = a`, so the three-way condition states the semantic rather
> than adding a restriction. §4.1 is corrected with it.

| Family | Degenerate when | Returned sample | Uniforms consumed | Dispatch entered? | Stream position |
|---|---|---|---|---|---|
| Uniform | **`a = b`** | `a` | **0** | no | unchanged |
| Triangular | `a = m = b` | `a` | **0** | no | unchanged |
| Beta-PERT | `a = m = b` | `a` | **0** | no — `r` is never formed, so `0/0` cannot arise | unchanged |

**Why zero rather than one.** Revision 3 justified consuming one uniform for a
degenerate Triangular on the grounds that it kept "stream position unaffected by
the data". **That justification no longer holds under this architecture**: with
Cheng, consumption already depends on acceptance outcomes, which depend on drawn
values. Since each component owns its own stream (§5.6), a component consuming
nothing cannot perturb anything else, and consuming zero is both simpler and
cheaper. Degeneracy is a property of the *model shape*, fixed for the whole run,
so consumption stays constant across iterations either way.

**Consequence, stated:** a driver that changes from degenerate to non-degenerate
changes its own stream's consumption from iteration 1 onward. That is a model
change, so it is a new run, and the request fingerprint changes with it.

**Vectors required** (§15.2 families 13a–13c): one degenerate driver of each
family, proving the returned value, **the zero-consumption claim** — the
component's stream state after `n` iterations must equal its initial state — and
that no Beta parameterisation was attempted.

### 4.1 Uniform — one uniform

```
x = (1 − u)·a + u·b                     ' stable convex form, §4.6
```

`m` is not read (D1) — **not for the value, and not for degeneracy either**.
Degenerate `a = b` is handled by §4.0 — returns `a`, consumes nothing, never
reaches this formula — **whatever Most Likely holds, including a populated,
unrelated value**. Two Uniforms with the same `Min` and `Max` and different
ignored Most Likely values have identical sampling semantics and identical RNG
consumption.

### 4.2 Triangular — one uniform, inverse CDF

With `c = (m − a)/(b − a)`, computed on normalised values (§4.6):

```
u ≤ c :  x = a + sqrt( u · (b − a) · (m − a) )
u > c :  x = b − sqrt( (1 − u) · (b − a) · (b − m) )
```

Boundary cases:

| Case | Result |
|---|---|
| `a = b` | handled by §4.0 before dispatch: returns `a`, **consumes zero uniforms**, never reaches this formula |
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

**Dispatch rule — SETTLED-IN-PLAN, including the boundary:**

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

**Why not the alternatives — PROVISIONAL / PRIOR MEASUREMENTS.**

The figures below were obtained during planning, by scripts that were **not
retained**. Under §20.2 that makes them **prior indications, not evidence**:

| | Architecture | Prior indication (unretained) | Provisional verdict |
|---|---|---|---|
| A | one-uniform numerical inverse CDF | incomplete-beta continued fraction up to **~91 iterations** over this family; ~91,000 flops/sample; `~2.7 × 10^12` flops at the design target | *provisionally* infeasible |
| C | precomputed per-driver table | max normalised error `~5.9 × 10^-3` at 4096 nodes — at `r = 0.01`, `α = 1.04`, near-infinite tail slope | *provisionally* indefensible |
| **B** | **Cheng BB/BC** | prior exploration observed target peak density `~≤ 5.0` over the family. **This is NOT evidence of Cheng acceptance efficiency** — see below | *provisionally* recommended |

**The density observation does not establish the acceptance rate.** Revision 5
inferred that a bounded peak density means "acceptance is bounded away from
zero". That inference is invalid: acceptance efficiency is a property of the
**proposal and the algorithm**, not of the target's peak density. A bound of that
kind would apply to a *uniform-envelope* rejection method, which is not what
Cheng BB/BC does. The observation is retained as context and carries no weight.

**The retained Step-0 package must measure Cheng's behaviour directly** (§20.2):
acceptance rate · uniforms consumed per accepted sample · the mean · the
high-percentile and tail of that consumption · **by shape, across the whole PERT
family**. **D6-04 closes on measured Cheng behaviour, never on a density
proxy.**

**No option is finally accepted or rejected on these numbers.** A decision that
changes the architecture may not rest on measurements nobody can re-run. The
retained Step-0 package (§20.2) must **reproduce** them — including the
counting methodology — before D6-04 closes, and if it does not reproduce them the
verdicts change with it.

**D6-04 remains OPEN** and now covers the sampler, the component-stream scheme
and the jump arithmetic **as one decision**, closing only on the retained Step-0
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

**Shared invariant, whichever option D6-18 settles on:**

> the occurrence stream advances **exactly once per Risk per iteration**,
> unconditionally.

```
' --- OPTION A: severity sampler invoked only on occurrence -----------------
occurred = ( u_occurrence < Probability )      ' 1 uniform, OCCURRENCE stream
if occurred:
    severity   = sample(dist, a, m, b)         ' SEVERITY stream, variable count
    contribNom = severity × Knom
    contribPv  = severity × Kpv
else:
    contribNom = 0 ;  contribPv = 0            ' severity stream NOT advanced

' --- OPTION B: severity sampler invoked every iteration --------------------
occurred = ( u_occurrence < Probability )      ' 1 uniform, OCCURRENCE stream
severity = sample(dist, a, m, b)               ' SEVERITY stream, ALWAYS advanced
if occurred:
    contribNom = severity × Knom
    contribPv  = severity × Kpv
else:
    contribNom = 0 ;  contribPv = 0            ' value discarded, stream advanced
```

**Neither is chosen here.** Revision 4 recommended B in the decision table while
its pseudocode silently implemented A; both are now written out, and the main
execution path stays neutral until D6-18 closes.

`Probability` is **not** folded into `Knom`/`Kpv` — the locked Phase-5 rule.
Strict `<`, so `Probability = 0` never occurs and `Probability = 1` always
occurs, given `u ∈ (0,1)`.

**Revision 2 said every risk consumes exactly two uniforms while also saying the
severity transform is conditional. Both cannot hold under a variable-draw
sampler, and the contradiction is resolved by separating the streams**, not by
drawing a uniform nobody uses:

- the **occurrence** stream advances by exactly one uniform per iteration,
  **unconditionally, under both D6-18 options**, so the occurrence path is a
  function of `Probability` and the seed alone
- the **severity** stream advances per D6-18: under **Option A** the sampler is
  invoked only when the risk occurred; under **Option B** it is invoked every
  iteration and its value used only when the risk occurred. In both cases the
  number of uniforms one invocation consumes is variable

Because they are different streams, variable severity consumption **cannot**
perturb the occurrence path. That is the property §5.6 exists to guarantee.

#### D6-18 — when does the severity stream advance?

The design above insulates the *occurrence* path from severity consumption. It
does **not** insulate the *severity* path from occurrence history: under
conditional advancement, changing `Probability` changes which iterations advance
the severity stream, and therefore which severity value lands on iteration `i`.

| | Design | Cost | What is preserved |
|---|---|---|---|
| **A** | **Conditional** — invoke the severity sampler only when the risk occurred | lower when `p < 1`; at `p = 0.1`, one tenth the severity work | occurrence path only |
| **B** | **Unconditional** — invoke the severity sampler every iteration, use the value only when the risk occurred | full severity cost regardless of `p`; identical to A at `p = 1` | occurrence path **and** the iteration-indexed severity path |

For Beta-PERT "one severity sample" may consume several uniforms under Cheng;
the question is whether the sampler is *invoked* every iteration, not how many
uniforms one invocation takes.

**Which invariance does PCCM actually want?** The concrete case is a user asking
what changing one Risk's `Probability` does to the contingency. Under **A** that
comparison is confounded: the probability changed *and* every later severity
value moved to a different iteration, so two runs differ for two reasons and
neither can be attributed. Under **B** the severity path is stable and the
comparison isolates the probability change. PCCM exists to support exactly that
kind of decision question, so **B is the semantically correct answer.**

**Recommendation: B, contingent on measurement.** Its cost is a full severity
sampler invocation per risk per iteration whatever `p` is — at the design target
`1.0 × 10^7` Beta samples that A would have skipped at low probabilities. That
cost is unmeasured until Step 0 returns Cheng's actual throughput. **If B proves
prohibitive, A is acceptable**, and the confounding above must then be documented
as a stated limitation rather than discovered by a user.

**D6-18 is OPEN and must close in Step 0**, because it determines performance,
stream-consumption semantics, replay, cross-run comparability and Phase-7
sensitivity together.

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
## 5. RNG — the proposed contract

### 5.1 Algorithm

**MRG32k3a** (L'Ecuyer 1999), as named in the locked roadmap.

**`Rnd` is unacceptable, and revision 5 overstated by how much.** It said the
period was "exhausted by more than four orders of magnitude". The arithmetic:

```
Rnd period          2^24 = 16,777,216
design-target draws      ≈ 40,000,000
ratio                    ≈ 2.4× the period      (0.38 orders of magnitude)
```

**A single run consumes roughly 2.4 times the period** — more under Option B,
but nothing like 10,000×. The conclusion is unchanged and does not need the
exaggeration: *a generator whose period is below a plausible single-run
consumption is unacceptable*, because the run would revisit the same values
within itself. `Rnd` and `Randomize` fail independently on reproducibility — no
seeding contract, no explicit stream, undocumented and version-dependent
behaviour — which alone disqualifies them from an evidence model built on exact
vectors.

### 5.2 State, constants and output — INHERITED algorithm, SETTLED-IN-PLAN parameters

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

**D6-02 is SETTLED-IN-PLAN: canonical MRG32k3a, output `(0,1)`.** The algorithm
is INHERITED (§1.2); its correct combination is what this revision settles.

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

### 5.3 Seed admissible domain — D6-20

Revision 5 added **D6-19**, which asks *who owns* this rule. It did not classify
*what the rule is*: §5.3 proposed a domain that appeared in no decision table and
carried no authority status. An accepted plan may not contain an unclassified
semantic rule, so the domain is now **D6-20**.

**Proposed domain:**

- admissible **`1 … 2147483646`**, whole numbers
- `0`, negatives and non-integers are **refused**, not coerced
- blank is legal and means `seed_mode = AUTO` (§5.4)

**Is there a credible alternative?** The constraints are tight enough to examine:

| Candidate | Verdict |
|---|---|
| `1 … 2147483646` | fits signed `Long` with a spare value; every member is a valid six-word state under D6-05 candidate A (`2147483646 < m2 < m1`); excludes the all-zero state by construction |
| include `0` | **rejected** — under D6-05 candidate A a seed of `0` produces the all-zero state, which MRG32k3a forbids. Admitting it would require a special case for exactly one value |
| allow negatives | **rejected** — they would need folding into the positive range, and two distinct user seeds would then collide silently |
| widen beyond `Long` | **rejected** — the cell is declared `integer`, `LongLong` is 64-bit-Office only, and nothing needs more than `2.1 × 10^9` distinct seeds |
| `1 … m2 − 1` (`4294944442`) | technically valid for the state, but exceeds `Long` and buys nothing |

**No credible alternative survives**, so D6-20 is **SETTLED-IN-PLAN** rather than
left open — but its standing is now explicit rather than implied, and it is a
proposal until the plan is accepted.

**D6-20 states the rule. D6-19 states which file owns it.** They are separate
questions and neither answers the other.

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
  fingerprint, new nonce, new effective seed, new `run_id`, and a **result digest
  that may differ and normally does** — though on a degenerate model it legally
  will not (§11, §15.3). §11 explains why that is not a contradiction.
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
- Phase-7 replay possible at **driver** granularity without replaying unrelated
  drivers — a Cost Line through its one sampling component, a Risk through its
  occurrence and severity components **together** (§12.1)
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

## 7. Selected Confidence Level — INHERITED as a reporting selector

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

**D6-06 SETTLED-IN-PLAN. D6-10 SETTLED-IN-PLAN by rejection** — there is no
`UNSELECTED` state.

---

## 8. Statistics

### 8.1 Moments

- **Sample mean**: `Σx / n`, accumulated scale-safely (§4.6).
- **Standard deviation — SETTLED-IN-PLAN: the sample standard deviation, divisor
  `n − 1`.** A **reporting-method decision taken here**, not inherited from Excel
  or any library: the run is a sample from the model's distribution, not the
  population. **D6-09 SETTLED-IN-PLAN.**
- Both computed with the accepted Phase-5 safe primitives and the scale-safe
  strategies of §4.6 — **not** naive `Σx²`, and **not** unguarded Welford.

### 8.2 Percentile algorithm — SETTLED-IN-PLAN

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

### 8.3 The reported ladder — SETTLED-IN-PLAN

Stored for **nominal** and **PV** separately: fixed headline `P10, P50, P70, P90`;
the full selectable ladder `P50 … P95`; mean and sample standard deviation;
minimum and maximum. Eleven distinct percentiles per measure.

### 8.4 Annual output — out of scope, and not "reported" either

Revision 3 said Phase 6 "reports the analytical annual series Phase 5 already
produces", while also deferring annual simulated distributions to Phase 7 and the
Annual Cash Flow presentation to Phase 8. Reporting an annual series *is*
presentation, so that sentence claimed a scope the same document denied twice.

Corrected:

> **Phase 6 adds no annual stochastic output and no new Annual Cash Flow
> presentation.** The existing accepted Phase-5 analytical annual data remains
> unchanged and available to later presentation phases. Phase 6 retains no
> per-year samples and writes no annual block.

Annual simulated distributions: **Phase 7**. Annual Cash Flow presentation:
**Phase 8**.

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
`SafeSubtract`. **D6-07 SETTLED-IN-PLAN.**

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

> **POST-ACCEPTANCE AUTHORITY CORRECTION — Phase-6 Step-1 review.**
>
> Revision 6's predicates were **neither mutually exclusive nor collectively
> complete**, and they contradicted the inherited Phase-5 orthogonality this
> document's own authority matrix records: *current state decides; the
> historical attempt result is a separate axis*.
>
> **The hole.** Success `A` → an invalid edit → `RunSimulation` REFUSED → the
> user restores exactly request `A`. The fingerprint now matches and the last
> attempt is `REFUSED`, so CURRENT was false, STALE was false and INVALID was
> false. **There was no state.**
>
> **The overlap.** A stored success, valid but changed inputs, and a `FAILED`
> last attempt made STALE and INVALID both true.
>
> The three LABELS are unchanged and no fourth state is created. What is
> corrected is the DERIVATION, which no longer reads the attempt history at all.

**The derived status, in order. The first matching rule wins:**

| # | Condition | Status |
|---|---|---|
| 1 | current simulation prerequisites do not resolve | **INVALID** |
| 2 | prerequisites resolve **and** no successful simulation snapshot exists | **BLANK** — see below |
| 3 | prerequisites resolve, a snapshot exists, and the recomputed `request_fingerprint` **equals** the stored successful one | **CURRENT** |
| 4 | prerequisites resolve, a snapshot exists, and the fingerprints **differ** | **STALE** |

The rules are ordered, so they are **mutually exclusive** by construction, and
they are **total**: rule 1 or rule 2 catches everything rules 3 and 4 do not.

**The attempt result — `NONE`, `SUCCESS`, `REFUSED`, `FAILED` — MUST NOT
participate in the derivation.** It is an orthogonal audit axis, exactly as
`last_attempt_result` is in the accepted Phase-5 calculation state. A refusal
that the user has since corrected does not make a matching request stale, and a
failure does not make a valid changed request invalid.

**BLANK is the absence of a successful-comparison state, not a fourth label.** A
status can only be CURRENT or STALE *relative to a stored success*, so with no
success there is nothing to be relative to. `status_evaluated_at` may still be
populated while the status is blank, which is what distinguishes *never
evaluated* from *evaluated, and there is no successful simulation to compare
against*.

**Worked cases:**

| Case | Status | `last_attempt_result` |
|---|---|---|
| success `A` → REFUSED invalid edit → restore `A` | **CURRENT** | `REFUSED` |
| success `A` → valid changed request `B` | **STALE** | any |
| success `A` → FAILED on `B`, rolled back; viewing `B` | **STALE** | `FAILED` |
| success `A` → FAILED on `B`, rolled back; restored to `A` | **CURRENT** | `FAILED` |
| current prerequisites invalid | **INVALID** | any |
| no successful simulation, current request valid | **BLANK** | `NONE`/`REFUSED`/`FAILED` |
| no successful simulation, current request invalid | **INVALID** | any |

**There is still no fourth state.**

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

### 10.4a D6-11 — narrowing the Phase-4 guards without deleting them

Revision 4 called D6-11 OPEN in §1.2 and SETTLED-IN-PLAN in §20 and specified no
decision at all. Specified here.

**The mechanism that actually exists.** `structure_contract.yaml` declares
`vba.forbidden_constructs` as a **flat list of strings**, and it is enforced in
two places, both **global**:

| Enforcement | Scope |
|---|---|
| `test_15_no_forbidden_construct_appears_in_phase_4_vba` | every hand-written module, no exceptions |
| the Gate-B P5-EV sweep over the persisted project | every component, over *executable* code with comments and string literals stripped |

**The existing schema cannot express "allowed in this module only".** There is no
per-module field, so any narrowing is a **schema extension to an accepted
Phase-4 authority**, and that reopening must be authorised rather than assumed.

**Construct by construct:**

| Construct | Phase-6 need | Resolution |
|---|---|---|
| `Rnd(` | none — Phase 6 never uses it | **remains globally forbidden**, unchanged |
| `Randomize` | none | **remains globally forbidden**, unchanged |
| `NPV` | none | **remains globally forbidden**, unchanged. Not loosened merely because Phase 6 begins |
| `Percentile` | Phase 6 computes percentiles | **resolvable by naming discipline, with no contract change**: the internal implementation is named `…Quantile…`, never `…Percentile…`, so the guard never fires. Worksheet and library percentile calls stay forbidden, which is what the guard was protecting |
| `MRG32k3a` | the RNG module must name it | **requires scoping** — no naming trick avoids it |
| `RunSimulation` | the endpoint is `PCCM_RunSimulation`, which contains it as a substring | **requires scoping** — renaming the endpoint would contradict the Phase-5 out-of-scope list, which names `RunSimulation` as the future entry point |

**Proposed extension, minimal and backward-compatible.** An entry may optionally
carry an `allowed_in` module list; an entry without one behaves exactly as today:

```yaml
forbidden_constructs:
  - "Rnd("                      # global, unchanged
  - "Randomize"                 # global, unchanged
  - "NPV"                       # global, unchanged
  - "Percentile"                # global, unchanged - naming discipline avoids it
  - construct: "MRG32k3a"
    allowed_in: ["modSimRng"]
  - construct: "RunSimulation"
    allowed_in: ["modSimReport"]
```

**This narrows nothing that was protecting anything.** A duplicate RNG
implementation in a second module still fails; a stray `PCCM_RunSimulation` call
outside the declared owner still fails. What changes is that the declared owner
is allowed to exist.

**Consequences that must be carried in Step 1:** the loader and validator must
accept both entry shapes; `test_15` and the Gate-B P5-EV sweep must both learn
the scoped form; and a mutation control must prove that an `allowed_in` module
list cannot be widened silently — an entry whose `allowed_in` names a module that
does not exist, or names every module, must fail.

**D6-11 is SETTLED-IN-PLAN** — one classification, not two — with the schema
extension flagged as a reopening of an accepted Phase-4 authority that Step 1
must carry out explicitly.

### 10.4b D6-19 — who owns the Random Seed admissible domain

Revision 4 proposed putting the seed domain in `sim_contract.yaml` while
`input_contract.yaml` already owns `inpRandomSeed` and says its admissible domain
"is fixed when the RNG is implemented". **Without an ownership rule that is two
authorities for one semantic rule**, which is the defect this project has spent
five phases avoiding.

| | Architecture |
|---|---|
| **A** | **Input-contract ownership.** Step 1 updates `input_contract.yaml` with the now-resolved admissibility (`1 … 2147483646`, whole). `sim_contract.yaml` **references** that rule and owns only RNG and seeding *semantics* — never a duplicate range |
| **B** | **Split ownership.** `input_contract` stays the authority for what may physically exist in `C21`; `sim_contract` owns a stricter execution-time admissibility checked at `PCCM_RunSimulation` |

**Recommendation: A.** The input contract's own note says the domain is to be
fixed when the RNG exists — it is *waiting* for this decision, not declining it.
Resolving it there discharges that note, keeps one rule in one file, and needs no
argument about why two files disagree.

**B is defensible only with an explicit justification** for why `validation:
null` should remain — for instance, that Excel data validation cannot express the
rule cleanly and a storable-but-invalid value must be refused at simulation time
instead. If B is chosen, that justification must be written down, not implied.

**The same discipline applies to every other Phase-6 input whose business rule
already belongs to `input_contract.yaml`** — `monte_carlo_iterations` most of
all. Its `≥ 1000` minimum stays where it is; Phase 6 adds only the *technical*
storage ceiling (§6.1), which is a different category and belongs in
`sim_contract.yaml`. **No semantic rule is maintained in two files.**

**D6-19 is OPEN**, with A recommended.

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
run_id                changes          ← on successful commit
effective_seed        changes
result_digest         MAY change       ← and normally does; see below
```

That is correct, not contradictory.

**`result_digest` inequality is not guaranteed even here.** Consistently with
§15.3, a different effective seed always produces a different RNG stream, but it
does not always produce different retained outputs: a fully degenerate model, or
one whose Risks all sit at `Probability = 0`, maps every seed to the same totals.
Equal digests across two AUTO runs of such a model are **legal**, not a defect.

**Run identity survives that.** Two AUTO runs remain distinguishable by `run_id`
and `effective_seed` whether or not their digests coincide — which is exactly why
those are separate identities from the digest.

**`result_digest` definition — D6-17.** The accepted Phase-5 canonical encoder
and hash, applied to the iteration-ordered outputs: for `i = 1 … n`, the
canonical `Double` encoding of `total_nominal[i]` then `total_pv[i]`, with the
iteration index and `n` in the stream so a truncated run cannot collide with a
short one. The exact field order and tagging is Step-0 work. **This is what G2
and G3 compare** (§15).

### 11.1 `_SimData` persists iteration order — SETTLED-IN-PLAN

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
persisted arrays are never permuted.** **D6-13 SETTLED-IN-PLAN by adoption.**

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

#### Exhaustion — defined, not assumed unreachable

Both counters are `Long`. A full-period mapping over `1 … 2147483646` yields
that many distinct seeds and no more, and `run_id` is bounded by `Long` itself.
**Practically unreachable — 2.1 billion runs — but "unreachable" is not a
specification**, and a silent wrap would turn a reused seed into something the
workbook presents as new.

| Requirement | Rule |
|---|---|
| Silent wrap | **forbidden** |
| Negative rollover | **forbidden** — the refusal fires before `Long` can overflow |
| Seed reuse presented as "new" | **forbidden** — this is the failure the rule exists to prevent |
| Reset on Save/Reopen | **forbidden** — both counters are persisted workbook state |
| Behaviour at exhaustion | an **explicit, auditable refusal** naming the exhausted counter and its limit; the prior successful simulation is untouched |

`auto_nonce` exhaustion refuses the AUTO run and states that the representable
unique-seed domain is exhausted — the user can still run with an explicit `FIXED`
seed, so the model is not bricked. `run_id` exhaustion refuses any further
commit, because a result that cannot be identified cannot be published.

**Both belong to D6-03 and D6-15 respectively and must close with them.**

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
"direct seek" to iteration `i`; that claim is withdrawn (§5.8). Revision 4 then
over-corrected in the other direction, saying **any component's** stream can be
replayed "in isolation, without needing any other component". **That is too broad
for a Risk**, and the correct statement is at driver granularity:

> Phase 7 can replay **one driver** without replaying unrelated drivers:
> a **Cost Line** needs its one sampling stream; a **Risk** needs **both** its
> occurrence and severity streams, paired by iteration.

A Risk's contribution is not a function of the severity stream alone — the
occurrence draw decides whether the severity value contributes at all, so
reconstructing a Risk's per-iteration column requires both, advanced together.

**And the replay procedure differs by D6-18:**

| | Replay of one Risk |
|---|---|
| **Option A** | The severity stream's iteration mapping **depends on the occurrence path**, because it advances only when the risk occurred. Both streams must be replayed **in lockstep from iteration 1** — the severity stream cannot be positioned without knowing how many occurrences preceded |
| **Option B** | Both streams are independently deterministic per iteration, so each can be advanced separately; the contribution still requires **pairing them by iteration index** |

Under either option the guarantee is per-driver isolation, not per-component
isolation — sufficient for sensitivity, which needs whole driver columns.

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
| Severity sampler invocations | **depends on D6-18** — see below | |
| Total Beta samples, worst case | `3.0 × 10^7` | |
| Uniforms consumed | **not a fixed number** | bounded by acceptance rate; the Step-0 package measures mean and high-percentile consumption per shape (§17) |
| Multiply-accumulates | `≤ 6.0 × 10^7` | nominal + PV on evaluated contributions |
| Sorts | 2 × `10^5` elements | `~3.4 × 10^6` comparisons |
| Stream jumps | `400`, **once**, before the loop | §5.9 |
| **Worksheet / COM calls inside the loop** | **0** | the locked invariant |

**Severity sampler invocations, by D6-18 option**, at 100 Risks × 100,000
iterations:

| | Invocations | Depends on `Probability`? |
|---|---|---|
| **Option A** | the number of **occurrences**: `0 … 1.0 × 10^7`, determined by the probabilities and the occurrence path | **yes** |
| **Option B** | **exactly `1.0 × 10^7`** | **no** |

They coincide only at `p = 1`. **The `p = 1` case is therefore useless for
comparing A against B** — it is the one point where the options are identical.
§20.2 requires at least one representative **lower-probability** operation
comparison, because that is where the cost difference lives: at `p = 0.1`,
Option A performs one tenth the severity work Option B performs, targeting the
same probability law but following a different iteration-indexed severity path.

*(Wording corrected after Revision 6 was accepted: this sentence previously read
"for identical statistical output". Both options draw severities from the same
intended distribution, but they do not produce identical output — they consume
the stream differently, so the realised severity sequence differs. The
correction is **non-semantic**: it changes no decision, no number and no
requirement, and the surrounding argument — that the options differ in work and
in severity path, and coincide only at `p = 1` — is unchanged.)*

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
| **F1** | full seeded simulation, **no Beta drivers**, general fixture | Python vs VBA end-to-end, per-iteration and summary outputs | **tolerance-bounded** under §15.1 — no rejection path to desynchronise, but a transformed sample may still differ by an ULP |
| **F2** | full seeded simulation, **no Beta drivers**, **exact-friendly fixture** | Python vs VBA `result_digest` | **EXACT**, for that fixture only — see §15.4 |
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
- layers F1 and F2, which exercise the *whole engine* end to end with no
  rejection path — F1 broadly under tolerance, F2 exactly on one constructed
  fixture

**The Python oracle stays genuinely independent.** Forcing bit-identical Beta
results by transcribing the VBA into Python would manufacture agreement and
destroy the only reason to keep two implementations.

### 15.0 Why layer F had to be split

Revision 4 said layer F was "EXACT and strong" while §15.1 said transformed
Uniform and Triangular samples are **not** universally bit-identical across
independent runtimes. **Both cannot be true.** A single-ULP difference in one
transformed sample changes that sample's canonical `Double` encoding, which
changes the accumulated total, which changes `result_digest` — so a
cross-language digest comparison over an arbitrary no-Beta fixture is asserting
exactly the universal exactness §15.1 declines to claim.

Split into **F1** (general, tolerance-bounded) and **F2** (one constructed
fixture where exactness is *justified*, not assumed). Nothing about the
same-runtime guarantees changes: **G2 and G3 remain exact**, because they compare
one implementation against itself.

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

### 15.1a The exact-friendly fixture — F2

Cross-language `result_digest` equality is required **only** for a fixture
deliberately constructed so that every value is binary-exact at every step:

| Requirement | Why |
|---|---|
| Uniform and Triangular drivers only | no rejection path |
| Endpoints, quantities and factors chosen as **small dyadic rationals** — values with short binary expansions | `(1 − u)·a + u·b` is then exact |
| Uniforms whose products with those endpoints are exactly representable | no rounding in the transform |
| Triangular cases restricted to those whose `sqrt` argument is a **perfect square** in binary64, or avoided entirely | `Sqr` is the least portable operation in §15.1 |
| Few drivers and few iterations, so the accumulation cannot round | the sum is exact term by term |
| The exactness of each step **demonstrated in the fixture's own documentation**, not assumed | this is the justification that makes the claim admissible |

If a candidate fixture cannot be shown exact at every step, it belongs to **F1**
and is compared under tolerance. **F2 is a narrow, justified exception; it is not
a general claim about the engine.**

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
| 15a | mixed 3 cost + 2 risk, **no Beta**, general fixture — full Python-vs-VBA end-to-end **under the §15.1 tolerance policy**; no cross-language digest equality required |
| 15b | the **exact-friendly** fixture of §15.1a — cross-language `result_digest` equality **required**, with the step-by-step exactness justification attached |
| 16 | same seed → identical `result_digest` |
| 17 | changed seed → different RNG stream, **always**; and on a **deliberately non-degenerate stochastic fixture** whose sampled uncertainty affects the retained total, different `result_digest`. See §15.3 |
| 18 | **row reorder → identical `result_digest`** |
| 19 | percentile vectors for `n = 1, 2, 3, 4, 10` |
| 20 | extreme-domain pipeline vectors, §4.6 cases 1–10 |
| 13a–c | degenerate `a = m = b` for **each** family: returned value, **zero consumption** proved by the component's stream state after `n` iterations equalling its initial state, and no Beta parameterisation attempted (§4.0) |
| 21 | **accumulation-order control** on a **constructed non-associative fixture** (§3.5): stream identities and sampled values held fixed, accumulation order reversed, the oracle having first proved the two orders give different binary64 totals. The difference is required only for this fixture; row-order invariance itself is universal and is family 18 |
| 22 | counter exhaustion: `auto_nonce` and `run_id` at their limits refuse explicitly, without wrap (§11.2) |

### 15.3 Digest inequality is not a universal invariant

Revision 3's "changed seed → different `result_digest`" is **too strong as a model
invariant**. A different seed guarantees a different RNG stream under the seeding
contract; it does **not** guarantee different published totals for every valid
model. Counterexamples that are entirely legal:

- every driver degenerate (`a = m = b`) — no sampled value can vary
- every Risk at `Probability = 0` — no severity ever contributes
- any model whose sampled variability cannot reach the retained total

For such fixtures, **equal `result_digest` across different seeds is correct
behaviour**, and a test asserting inequality would be asserting a defect.

The test is therefore scoped in two parts:

| Claim | Scope |
|---|---|
| Different seed → different RNG stream / state | **universal**, over the seeding contract |
| Different seed → different `result_digest` | **only** on a fixture deliberately constructed to be non-degenerate, with sampled uncertainty that demonstrably reaches the retained total |

Gate-B wording changes with it: divergence is asserted on the non-degenerate
fixture, not on any model.

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
be reconsidered only if measured Gate-B runtime justifies it. **D6-12
SETTLED-IN-PLAN**, on that basis.

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
prove* (§15); same-runtime replay digest; **row-order invariance digest** (both
stream assignment **and** accumulation order, §3.5); the full 100,000-iteration
run at the design target; the pre-flight storage-ceiling refusal; failure
containment; natural Excel shutdown; clean COM release.

**Different-seed evidence, scoped.** Revision 5's Gate-B line said only
"different-seed divergence", which reads as a universal digest claim that §15.3
explicitly denies. What Gate B proves is two separate things:

| | Claim | Scope |
|---|---|---|
| **A** | two different accepted **FIXED** seeds produce the expected different RNG initial stream and state | **universal** |
| **B** | on the **designated non-degenerate divergence fixture** — sampled uncertainty demonstrably reaching the retained total — the retained distribution and `result_digest` diverge as expected | **that fixture only** |

**Equal `result_digest` across different seeds on a fully degenerate legal model
remains valid**, and no Gate-B wording may imply otherwise.

**A performance benchmark is not a semantic test.** The 100k run proves the
budget is met and nothing about correctness; Gate B passes only if the semantic
scenarios pass independently of it.

---

## 19. Implementation sequence

**Step 0 comes first.** A contract cannot be the authority for decisions taken
after it exists.

| # | Step | Files | New authority | Gate-A acceptance |
|---|---|---|---|---|
| **0** | **Authority closure**, including the **retained D6-04 feasibility package** of §20.2 | `docs/`, plus a non-production evidence area for the retained package | the closed decisions | **every Class-1 and Class-2 decision closed and independently accepted**; the retained evidence package accepted; **no unresolved semantic choice remains for `sim_contract.yaml`**. **D6-08 is explicitly not a Step-0 blocker** — it is a Class-3 constant derived in Step 1. **No contract file yet** |
| 1 | **`spec/sim_contract.yaml`** — RNG and jump constants, AUTO nonce mapping, stream assignment, Cheng dispatch, percentile method, standard-deviation method, contingency, `SIM` fields, `result_digest` canonicalisation, `sim_state` schema, `_SimData` layout and therefore `H`, versions. **Plus the seed-admissibility change in whichever contract D6-19 names as owner** — see below. Loader + validator, fail-loud | `spec/`, `builder/`, `tests/` | the sixth contract; possibly an update to an existing one | validator rejects every malformed shape; mutation controls per rule; **no duplicated seed range** |
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

**Step 1 is ownership-neutral on the seed range until D6-19 closes.** Revision 5
listed "seed domain" among `sim_contract.yaml`'s contents, which pre-empts D6-19's
own recommendation. The rule instead:

- the contract **selected by D6-19** becomes the single owner of seed
  admissibility and is the file Step 1 updates;
- the other contract **references** that authority and **must not duplicate the
  range**;
- `sim_contract.yaml` owns RNG and seeding **execution semantics** regardless of
  which file owns the range.

If D6-19 closes on **A**: `input_contract.yaml` owns `1 … 2147483646` and
`sim_contract` references it. If on **B**: the split is stated explicitly,
together with the written justification for why `validation: null` remains
intentional.

**The same one-authority rule applies to every Phase-6 Setup input.**
`monte_carlo_iterations` keeps its `≥ 1000` minimum in `input_contract.yaml`;
only the *technical* storage ceiling (§6.1) — a different category — belongs to
`sim_contract.yaml`.

**The Step-0 feasibility work is retained evidence, not throwaway code.**
Revision 3 called it throwaway; that is withdrawn. D6-04 changes the architecture
on the strength of measurement, and evidence that closes an authority decision may
not disappear once the decision is taken — otherwise the decision rests on numbers
nobody can re-run. The package is **committed and auditable** (§20.2).

It lives in a **non-production evidence area**, and it **must not become the
Phase-6 implementation by default**: it is written to measure, not to ship, and
step 2 onward implements against the contract rather than against the probe.

---

## 20. Open decisions

**Settled in this plan** (proposals, not authority until the plan is accepted):
D6-01 · D6-02 · D6-06 · D6-07 · D6-09 · D6-10 · **D6-11 (§10.4a)** · D6-12 ·
D6-13 · **D6-20 (§5.3)**. Each is SETTLED-IN-PLAN in exactly one place.

**D6-20 is settled rather than open** because the analysis in §5.3 found no
credible alternative — but it is stated as a decision with an explicit status,
which revision 5 failed to do at all. **D6-19, which asks who *owns* that rule,
remains OPEN** in Class 2; the two are separate questions.

Revision 3 said "eight open" in §20 and "one blocker" in §22 without saying how
those relate. They are different things, and the open items now fall into three
classes:

**Every decision listed in Classes 1, 2 and 3 below is OPEN.** No item appears in
both the settled list above and a class table below.

### Class 1 — architectural feasibility blockers

Cannot close on argument. They need measurement, and they can still change the
architecture this document describes.

| # | Decision | Options | Recommended | Consequence |
|---|---|---|---|---|
| **D6-04** | **Beta sampler + component-stream scheme + jump arithmetic, as ONE decision** | (a) one-uniform inverse CDF + fixed consumption; (b) **Cheng BB/BC + component streams + `2^127` jumps**; (c) precomputed table | **(b)**, subject to §20.2 | the stream architecture, the jump work, the oracle's layer boundaries, Phase-7 replay |
| **D6-18** | **When the Risk severity stream advances** | (a) conditional on occurrence; (b) **unconditional, value used only on occurrence** | **(b)** for correct attribution, contingent on Cheng's measured cost | performance, consumption semantics, replay, cross-run comparability, Phase-7 |

### Class 2 — ordinary Step-0 authority decisions

Close on argument and review; no measurement required.

| # | Decision | Options | Recommended | Consequence |
|---|---|---|---|---|
| **D6-03** | AUTO-seed source, freshness and **exhaustion**, jointly with D6-15 | (a) timestamp folded into the domain; (b) **`auto_nonce` → full-period multiplicative cycle**; (c) refuse blank | **(b)** | whether "new" is guaranteed or merely likely |
| **D6-05** | scalar seed → six-word state | (a) **repeated scalar**, canonical; (b) modular expansion | **(a)** | one more portability surface, or none |
| **D6-14** | analytical prerequisite | (a) **require Phase-5 `CURRENT`**; (b) refresh first; (c) one transaction | **(a)** | whether Phase 6 can mutate Phase-5 outputs |
| **D6-15** | `run_id` semantics and **exhaustion**, jointly with D6-03 | (a) **monotonic success counter, separate from `auto_nonce`**; (b) GUID; (c) timestamp | **(a)** | audit identity; no computational effect |
| **D6-16** | stream assignment rule | (a) **canonical sorted order → sequential streams**; (b) direct ID-derived index | **(a)**, revisit (b) if the ID numeric bound can be contracted | whether an unrelated driver's samples move when a driver is added |
| **D6-17** | `result_digest` canonicalisation | (a) **accepted encoder over iteration-ordered totals, with `n` and index in the stream**; (b) a separate scheme | **(a)** | what G2/G3 actually compare |
| **D6-19** | **who owns the Random Seed admissible domain** (§10.4b) | (a) **input-contract ownership, `sim_contract` references it**; (b) split ownership with a written justification for keeping `validation: null` | **(a)** | whether one semantic rule lives in two files |

### Class 3 — Step-1 constants derived from layout

| # | Decision | Recommended | Close before |
|---|---|---|---|
| **D6-08** | technical storage ceiling constant | `1048576 − H`, once the `_SimData` layout fixes `H` | **Step 1** |

**The end state before `sim_contract.yaml` exists is ZERO unresolved semantics
that the contract must encode.** Class 1 and Class 2 close in Step 0; Class 3 is
a constant the contract itself determines in Step 1, not a semantic question.

**Step-0 acceptance therefore reads:** every Class-1 and Class-2 decision closed
and independently accepted · the retained feasibility evidence accepted · no
unresolved semantic choice remaining for the contract. **D6-08 is not a Step-0
blocker**; it is derived in Step 1 from the accepted `_SimData` layout, before the
contract and its validator are considered complete.

### 20.2 The retained D6-04 Step-0 evidence package

D6-04 closes on **evidence, not prose**. Required deliverables:

1. **Cheng dispatch** implemented and demonstrated: BB when `min(α,β) > 1`, BC
   when `min(α,β) ≤ 1`.
2. **Parameter grid spanning the whole PERT family**, `α, β ∈ [1,5]`, including
   `1/5`, `5/1`, both near-1 boundaries, symmetric `3/3`, and interior shapes.
3. **Per shape, measured directly and not inferred from any density proxy:**
   **acceptance rate** · **uniforms consumed per accepted sample** · the **mean**
   of that consumption · its **high percentile and tail** · numerical output
   checks · independent theoretical mean/variance checks.
4. **The component-stream architecture** demonstrated: Cost → one stream;
   Risk → occurrence + severity streams.
5. **The exact stream-assignment rule** (D6-16), demonstrated row-order invariant.
6. **The exact jump matrices and modular arithmetic design** (§5.9), with
   cross-language vectors.
7. **Memory for all stream states** at the design target.
8. **Operation models under both D6-18 options**, at **two probability points**:
   the `p = 1` worst case *and* at least one representative lower probability
   (`p = 0.1` or similar). At `p = 1` the two options coincide, so that point
   alone cannot compare them; the lower-probability point is where A's saving and
   B's cost actually appear.
9. **Reproduction of the §4.3 prior indications** — the ~91 continued-fraction
   iterations, the `~5.9 × 10^-3` table error, and the `~≤ 5.0` peak density.
   Options A and C may not be finally rejected on unretained numbers.

**What "retained" requires.** The package is committed to a non-production
evidence area and must contain:

| | Item |
|---|---|
| 1 | the source scripts themselves |
| 2 | the exact inputs and parameter grids |
| 3 | the exact algorithm variants compared |
| 4 | Python and runtime versions where the result could depend on them |
| 5 | raw outputs, not only summaries |
| 6 | the summarised measurements |
| 7 | **the counting methodology** — what was counted as an operation, and why |
| 8 | expected theoretical values, where theory gives one |
| 9 | deterministic provenance — seeds, and hashes of inputs and outputs where appropriate |

A reviewer must be able to re-run it and obtain the same numbers. **It must not
become the production implementation.**

---
## 21. Stale statements corrected

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

### 21.1 Revision 3 → revision 4

| # | Where | Corrected |
|---|---|---|
| 26 | header | Three status labels introduced — **INHERITED** / **PROPOSED** / **SETTLED-IN-PLAN**. Revision 3 used "LOCKED" for both inherited authority and its own inventions; eight Phase-6 proposals relabelled |
| 27 | **§3.5 new** | **Row-order invariance did not cover accumulation.** Floating-point addition is not associative, so reordering drivers changes rounding even with identical streams. Canonical Permanent-ID evaluation and accumulation order locked, with a kind precedence and two independent accumulators |
| 28 | **§4.0 new** | **The degenerate case was inconsistent and would have divided by zero.** Triangular said one uniform, Uniform said nothing, Beta-PERT would have formed `0/0`. One rule for all three: detected before dispatch, returns `a`, **consumes zero uniforms**. Revision 3's justification for consuming one is withdrawn — it was already false under variable draws |
| 29 | §4.1, §4.2, §4.3 | degenerate rows aligned to §4.0 |
| 30 | §4.3 | The A/B/C figures reclassified as **PROVISIONAL / PRIOR MEASUREMENTS** from unretained scripts; no option finally accepted or rejected on them |
| 31 | **§4.5 D6-18 new** | Severity-stream advancement made an explicit decision. Conditional advancement insulates the occurrence path but **not** the severity path from occurrence history, which confounds a probability-only comparison |
| 32 | §8.4 | "Phase 6 reports the analytical annual series" **withdrawn** — reporting is presentation, which the same document defers twice. Phase 6 adds no annual stochastic output and no Annual Cash Flow presentation |
| 33 | §11.2 | Counter **exhaustion** defined: no silent wrap, no negative rollover, no reuse presented as new, no reset on Save/Reopen, explicit auditable refusal |
| 34 | §15.2 family 17, **§15.3 new** | "changed seed → different `result_digest`" **scoped**. Universal for the RNG stream; for the digest, only on a deliberately non-degenerate fixture. Equal digests across seeds are legal for degenerate models |
| 35 | §15.2 | families 13a–c (degenerate consumption), 21 (accumulation-order control), 22 (exhaustion) added |
| 36 | §19, **§20.2** | "throwaway measurement code" **withdrawn**. The Step-0 package is retained, committed, reproducible evidence with nine stated contents, and must not become the implementation |
| 37 | §20 | Decisions split into **three classes** — feasibility blockers, ordinary Step-0 authority, Step-1 layout constants — resolving revision 3's "eight open" versus "one blocker" |

---

### 21.2 Revision 4 → revision 5

| # | Where | Corrected |
|---|---|---|
| 38 | §2.1 | "inverse-CDF sampling for Triangular, Uniform and **Beta-PERT**" — a scope statement contradicting §4. Split by family: Uniform affine, Triangular inverse-CDF, **Beta-PERT Cheng BB/BC acceptance/rejection** |
| 39 | **§15.0 new**, §15 table, **§15.1a new**, family 15 | **Layer F claimed cross-language exactness that §15.1 denies.** One ULP in a transformed sample changes the digest. Split into **F1** (general, tolerance-bounded) and **F2** (one constructed exact-friendly fixture whose exactness is demonstrated, not assumed). G2 and G3 stay exact — they compare one implementation against itself |
| 40 | §5.4, §11 | The AUTO re-run no longer asserts `result_digest` **changes**; it **may change and normally does**, and equality is legal on a degenerate or zero-probability model. Run identity survives through `run_id` + `effective_seed` regardless |
| 41 | §3.5, family 21 | The accumulation-order control now requires a **constructed non-associative fixture**, with the oracle proving the two orders differ in binary64 first. Row-order invariance itself stays universal |
| 42 | **§10.4a new** | **D6-11 specified.** Revision 4 called it open in §1.2 and settled in §20 and defined nothing. The existing schema is a **flat global list with no per-module scoping**, so narrowing requires a schema extension to an accepted Phase-4 authority — stated as such. `Percentile` is resolved by naming discipline with no contract change; `MRG32k3a` and `RunSimulation` need scoping; `Rnd(`, `Randomize` and `NPV` stay globally forbidden |
| 43 | **§10.4b new**, D6-19 | **Seed-domain authority ownership** decided as a question. Input-contract ownership recommended, because that contract's own note is waiting on this decision. No semantic rule in two files |
| 44 | §3.4, §4.5 | **D6-18 was hard-coded to option A** in the execution path while the table recommended B. Both options are now written out, with the shared invariant stated: the occurrence stream advances exactly once per Risk per iteration |
| 45 | §12.1 | **Phase-7 replay claim corrected from component to driver granularity.** A Risk needs both its streams paired by iteration; under option A they must be replayed in lockstep from iteration 1 |
| 46 | §19, §20 | **Step-0 acceptance no longer requires "every open §20 item"**, which contradicted D6-08's Class-3 status. It requires Class-1 and Class-2 closed, the evidence accepted, and no unresolved semantics for the contract |
| 47 | header, §1, §5 | Stale governance text: "Revision 3 describes ONE candidate architecture"; the §1 legend now states that every LOCKED there is prior accepted authority; "RNG — the locked contract" retitled "the proposed contract" |

---

### 21.3 Revision 5 → revision 6

| # | Where | Corrected |
|---|---|---|
| 48 | §4.5, §5.6, §14, §20.2 | **Residual D6-18 Option-A hard-coding removed.** §4.5 still said the severity stream "advances only when the risk occurred"; §5.6 still called Phase-7 replay per-component; §14 still counted severity draws as conditional only. All three now branch explicitly, with the shared invariant stated: the occurrence stream advances exactly once per Risk per iteration under **both** options |
| 49 | §14, §20.2 | Severity invocations given per option — Option A `0 … 1.0 × 10^7` depending on the occurrence path, Option B exactly `1.0 × 10^7`. **They coincide at `p = 1`, so that point cannot compare them**; a lower-probability comparison is now required in the evidence package |
| 50 | **§5.3 → D6-20 new**, §20 | **The seed admissible domain had no authority status at all** — proposed in prose, absent from every table. Now a classified decision, SETTLED-IN-PLAN, with the alternatives examined and rejected on stated grounds. D6-19 (ownership) and D6-20 (the rule) are separate |
| 51 | §19 step 1 | **Step 1 no longer pre-empts D6-19** by listing "seed domain" as `sim_contract.yaml` content. Ownership-neutral: the D6-19 winner owns the range, the other references it, `sim_contract` always owns execution semantics |
| 52 | §18 | **Gate-B "different-seed divergence" scoped** into a universal RNG-stream claim and a fixture-scoped digest claim, consistent with §15.3 |
| 53 | §5.1 | **Corrected arithmetic error.** `Rnd`'s `2^24` period is `≈ 2.4×` a single run's consumption — `0.38` orders of magnitude, not "more than four". The conclusion stands without the exaggeration |
| 54 | §4.3, §20.2 | **The density-proxy inference withdrawn.** Peak density of the *target* says nothing about a rejection algorithm's acceptance rate, which depends on the proposal. Step 0 must measure acceptance rate and consumption directly, by shape |
| 55 | §7, §8.1, §9, §11.1, §16 | Five `D6-xx CLOSED` labels changed to **SETTLED-IN-PLAN** — the plan is not accepted, so nothing it decides is closed in the sense an accepted authority is |

---

## 22. Remaining unresolved blockers

**Two, both in Class 1, and both need measurement rather than argument.**

**D6-04 — the sampler, streams and jump arithmetic.** The architecture is
coherent but rests on a sampler whose cost and floating-point stability are
unmeasured, and on prior indications from scripts that were not retained. If the
retained package fails to reproduce them, the verdicts on options A and C move
with the numbers, and the stream architecture moves with the sampler. If the
floating acceptance path proves *stable* across implementations, layer G of §15
could be strengthened from statistical to exact — which would materially improve
the evidence model rather than weaken it. The jump arithmetic must also be shown
exact against canonical vectors before any of it can be locked.

**D6-18 — severity-stream advancement.** Option B is the semantically correct
answer and is recommended, but it invokes the severity sampler on every iteration
regardless of `Probability`, and whether that is affordable depends entirely on
Cheng's measured throughput. It cannot be settled before D6-04's measurements
exist, and it changes the worst-case operation model, which is why §20.2 requires
the operation model under **both** options.

Everything else is either settled in this plan or closes on a decision that does
not depend on measurement.

---

## 23. What this plan does NOT claim

1. **No Phase-6 code exists.** No VBA, no builder change, no contract, no test,
   no simulation or oracle code in any language.
2. **No Windows or Excel runtime has been executed for Phase 6.**
3. **`spec/sim_contract.yaml` does not exist** and may not be created until
   Step 0 closes every decision it must encode.
4. **No runtime duration is claimed.** No benchmark has been run.
5. **No flop total is asserted** without its measurement artefact, and the
   §4.3 figures are prior indications from unretained scripts, not evidence.
6. **Step 0 has not been executed.** No feasibility measurement has been run
   under this plan, and none is authorised.
7. **Nothing here is accepted.** Ten decisions remain open across three classes,
   and **two of them — D6-04 and D6-18 — could still change the architecture this
   revision describes**. Every `SETTLED-IN-PLAN` item is a proposal, not
   authority. D6-04 and D6-18 are *expected* to remain open pending Step-0
   measurement; §20.2 states exactly what Step 0 must measure to close them.
8. **The Phase-5 baseline is untouched.** `f571154` remains the accepted
   executable baseline; this document changes no file under `src/`, `spec/`,
   `builder/`, `bootstrap/` or `tests/`.
