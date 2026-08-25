# PCCM Phase 6 — Step 4 authority record

Step 4 builds the **pure Python Monte Carlo oracle**: preparation from the
accepted Phase-5 resolved model, canonical component streams, Cost Line and Risk
contributions, per-iteration nominal and PV totals, the result digest, the
statistics, the percentile ladder and the contingency lookup.

**No workbook.** No `_SimData`, no Results publication, no simulation status, no
`run_id`, no `next_auto_nonce`, no attempt metadata, no `PCCM_RunSimulation`, no
VBA, no Stage-A Phase-6 emission, no sensitivity, no annual stochastic
distributions. Step 4 produces pure Python values.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Phase-6 Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Phase-6 Step-1 contract authority | `7d8e73e59347b5d3acb3f80d848be6aaf1d3eb83` |
| Accepted Phase-6 Step-2 RNG reference | `9bdf5a117784e99397dd80efee07d20df508ac9c` |
| Accepted Phase-6 Step-3 sampler reference | `e939d47d8129b6617797fa52fd79b542656ce009` |
| Step-4 simulation oracle | this commit — reported by the delivery message and `PROVENANCE.txt` |

**No contract defect and no Phase-5 authority defect was found**, so nothing in
`spec/` changed and `calc_oracle.py`, `calc_numeric.py` and `calc_fingerprint.py`
are untouched.

---

## 2. Files

| File | Change |
|---|---|
| `builder/pccm_builder/sim_oracle.py` | **NEW** — preparation, engine, digest, contingency |
| `builder/pccm_builder/sim_stats.py` | **NEW** — scale-safe mean, SD, Type-7 percentiles |
| `builder/pccm_builder/sim_sample.py` | the accepted §0 invalid-state cleanup **only** |
| `builder/pccm_builder/__init__.py` | Step-4 public exports |
| `tests/test_phase6_sim_oracle.py` | **NEW** — 74 conformance tests |
| `tests/test_phase6_sim_oracle_validation.py` | **NEW** — 37 mutation controls and scope guards |
| `tests/test_phase6_sim_sample_validation.py` | **+4** — the Step-3 invalid-state degenerate regressions |
| `docs/phase6_step4.md` | **NEW** — this record |

### The accepted Step-3 cleanup (§0)

`sample_uniform`, `sample_triangular` and `sample_prepared_beta` now call
`RngReference.validate_state` **before any path can return**, the
zero-consumption degenerate path included. Previously an INVALID state paired
with a degenerate distribution returned a constant and the sampler boundary never
noticed. Consuming nothing is a property of the DISTRIBUTION, not an exemption
from the state contract.

Three regressions were added — invalid state with a degenerate Uniform,
Triangular and Beta-PERT — plus a fourth asserting the cleanup **moved no
accepted number**. Consumption for a valid-state degenerate distribution remains
**zero** with the state unchanged, and every Step-3 vector is unchanged.

---

## 3. Public Step-4 API

```python
# preparation
prepare_simulation(reference, sim, inputs, model, tolerances,
                   *, effective_seed, iterations) -> (PreparedSimulationModel, CalculationResult)
validate_iterations(sim, inputs, iterations)      -> int
business_minimum_iterations(inputs)               -> int
resolve_percentile_ladder(sim, inputs)            -> PercentileLadder
effective_seed_from_nonce(reference, nonce)       -> int

# the engine
run_simulation(reference, prepared)               -> SimulationResult

# digest, standalone
result_digest(version, total_nominal, total_pv)        -> str
result_digest_stream(version, total_nominal, total_pv) -> str
validate_result_digest_contract(sim)                   -> None

# reporting
contingency_at(summary, selected_confidence_level, deterministic_base) -> Contingency
deterministic_base_of(result)                                          -> DeterministicBase

# statistics, pure
sample_mean(values) · sample_standard_deviation(values)
percentile_type7(values, p) · describe(values, points)
```

---

## 4. Preparation — one Phase-5 call, then nothing recomputed

`prepare_simulation` runs in this order and draws nothing:

1. **refuse an inadmissible iteration count** — before any allocation, any stream
   construction and any random draw;
2. resolve the percentile ladder from its owners;
3. call the accepted Phase-5 `calculate(model, tolerances)` **once**;
4. take `Quantity`, `Probability`, `Knom`, `Kpv` and the driver identities from
   its `DriverFactors`;
5. resolve Min / Most Likely / Max through the accepted Phase-5 resolvers
   themselves;
6. prepare each Beta-PERT shape **once per driver**;
7. construct every component's initial stream state **once**;
8. return — the loop is entered only after all of this succeeded.

**There is no second calculation model.** FX, inflation, profiling, discounting,
Quantity validation, Probability validation and distribution business validation
are performed by Phase 5, through its own public entry point, and none of them is
recomputed inside an iteration.

Min / ML / Max come from `_resolve_distribution` and `_resolve_three_point` —
private Phase-5 names, imported deliberately. The alternative is a duplicate
implementation of ordering, numeric coercion and the D1 "Uniform ignores Most
Likely" rule, which is exactly the second calculation model Step 4 is forbidden
to build. A duplicate is free to drift; a call is not.

### What a prepared driver retains

`permanent_id`, `driver_kind`, `distribution`, `minimum`, `most_likely`,
`maximum`, `quantity` (Cost Lines), `probability` (Risks), `knom`, `kpv`, the
prepared Beta shape where applicable, and its component stream indices and
initial states. **No worksheet object, no `ListObject`, no `Range`, no cell
address, no workbook handle** — and no FX rate, weight vector or inflation
series, because those have already collapsed into `Knom` and `Kpv`.

A Uniform's `most_likely` is **`None`** in the prepared model. Carrying the
entered value forward would let an ignored input travel through preparation
looking authoritative.

Preparation is canonical: all Cost Lines ascending Permanent ID, then all Risks
ascending Permanent ID, ordinal UTF-16.

---

## 5. Effective seed, not nonce lifecycle

The engine receives an **effective seed**. `effective_seed_from_nonce` wraps the
accepted Step-2 pure mapping and **persists nothing**: no `auto_nonce`, no
counter increment, no `run_id`, no attempt metadata, and no decision about
whether a failed attempt consumed a nonce. Neither `prepare_simulation` nor
`run_simulation` takes a nonce argument — there is nowhere for the transactional
lifecycle to enter.

---

## 6. Iteration-count pre-flight

| Owner | Value |
|---|---|
| Business minimum — `input_contract.yaml` `monte_carlo_iterations` | **1000** |
| Technical ceiling — `sim_contract.yaml` | **1048543** (`1048576 − 33`) |

Refused **before** allocation, stream construction and any draw. Rejected:
non-integer, `bool`, below the business minimum, above the technical ceiling. The
ceiling refusal says it is **technical** and is not presented as business
validation. **No smaller performance cap is invented.** The boundary is proved by
`validate_iterations` on its own; no test runs a million-row simulation.

Ordering is structural, not incidental: a model that would itself be refused
still reports the ITERATION problem first.

---

## 7. Component streams

For one effective seed, `base_state = fixed_seed_to_state(effective_seed)`, and
the accepted Step-2 assignment is used unchanged:

```
0  COST CL-001 value
1  COST CL-002 value
2  RISK R-001  occurrence
3  RISK R-001  severity
4  RISK R-002  occurrence
5  RISK R-002  severity
```

All Cost Lines first, then **each Risk interleaved** — occurrence then severity —
not three global blocks. Exactly one current state per component inside the local
run context; no two components share a state variable, and nothing survives the
call.

---

## 8. One Cost Line in one iteration

```
sampled_unit_cost = accepted Step-3 sampler
contrib_nominal   = safe_product(sampled_unit_cost, Quantity, Knom)
contrib_pv        = safe_product(sampled_unit_cost, Quantity, Kpv)
```

The sample is **unit cost** uncertainty. Quantity is deterministic, sits outside
the distribution and is applied **exactly once**. Probability never appears. PV is
an independent accumulator over the same driver order and is never derived by
discounting the sampled nominal contribution.

**Quantity, proved linear** (degenerate unit cost `250`, `Knom = 1`):

| `Quantity` | Retained total | Applied twice would give |
|---|---|---|
| 1.0 | **250.0** | 250.0 |
| 3.0 | **750.0** | 2250.0 |
| 7.5 | **1875.0** | 14062.5 |

**And the sample is on the unit scale, not the total scale.** With
`Quantity = 4` a unit sample lives in `[80, 150]` and a total sample would live in
`[320, 600]`; the two supports do not overlap, and every recovered sample lands
in the first.

---

## 9. One Risk in one iteration — D6-18b

Occurrence first, from the Risk's own stream: `occurred = u < Probability`,
strict, **exactly one uniform per Risk per iteration**. Then the severity sampler
is invoked on the Risk's independent severity stream **unconditionally**.

| `p` | Occurred / 5000 | Occurrence uniforms | Severity uniforms |
|---|---|---|---|
| 0.0 | **0** | 5000 | **5000** |
| 0.3 | 1487 (0.2974) | 5000 | 5000 |
| 1.0 | **5000** | 5000 | 5000 |

A rare risk makes the point sharply: at `p = 0.05` over 3000 iterations the risk
occurred **130** times and the severity stream consumed **3000** uniforms.
Sampling only on occurrence would have consumed about 130.

**Degenerate severity is still invoked**: `Triangular(90, 90, 90)` at `p = 0.4`
consumed **0** severity uniforms, left the severity stream state **unchanged**,
and produced totals `{0.0, 90.0}` — the occurrence stream still consumed 2000.

---

## 10. Probability-only comparability — the reason D6-18b exists

Same model, same effective seed, same severity distribution; only `Probability`
differs.

| `p` | Occurrences | Severity uniforms | Severity final state | Digest |
|---|---|---|---|---|
| 0.2 | 409 | 2000 | `2069675377 250440616 4154188787 912224118 624009514 3823340866` | `1C1A6DB04C193F69` |
| 0.8 | 1588 | 2000 | **identical** | `4108E7DA20AF408F` |

Occurrence decisions differ, contributions differ, digests differ — and the
iteration-indexed severity sequence and final severity-stream state are
**identical**. The sequence is also reconstructed independently, straight from
the stream's own initial state, and matches both runs.

The diagnostics that prove this are **oracle diagnostics**: one record per
component, carrying stream index, initial state, final state and uniforms
consumed. Not `_SimData`, not a persisted schema, and **not a retained sample
matrix** — seven components, not seven arrays of 100,000.

---

## 11. Canonical accumulation order

Every Cost Line ascending Permanent ID, then every Risk ascending Permanent ID,
ordinal UTF-16. Nominal and PV are **independent accumulators over the same
driver order**, both through the accepted `safe_signed_sum`. Physical row order
reaches nothing.

### The constructed non-associative fixtures

Binary64 addition is not associative, but it is not universally sensitive either:
`1e16 + 1` ties straight back to `1e16`, so a small term is absorbed or survives
depending on **when** the two large terms cancel. Each order mutation therefore
gets a construction built for it, and each states both sums.

| Construction | Canonical | Mutated |
|---|---|---|
| `[1e16, 1, −1e16]` | **0.0** | risks first `[−1e16, 1e16, 1]` = **1.0** |
| `[1e16, 1, −1e16, 1]` | **1.0** | reversed `[1, −1e16, 1, 1e16]` = **0.0** |
| `[1e16, 1, −1e16, 1]` | **1.0** | register order `[1, 1, 1e16, −1e16]` = **2.0** |

The engine produces the canonical value in every case.

**The difference is required ONLY on these fixtures.** No test claims that
reversing arbitrary contributions always changes a Double sum — a separate
control records that on an ordinary model reversal changes nothing on most
iterations. Row-order invariance is universal precisely because canonical
accumulation order does not depend on the register.

---

## 12. Row-order invariance

Three physically reordered variants of the same five-driver model, same seed:

| Register order | Prepared canonical order | Digest |
|---|---|---|
| `CL-001 CL-002 CL-003 / R-001 R-002` | `CL-001 CL-002 CL-003 R-001 R-002` | `3AFC9494722696D4` |
| `CL-003 CL-001 CL-002 / R-002 R-001` | **identical** | **identical** |
| `CL-002 CL-003 CL-001 / R-002 R-001` | **identical** | **identical** |

Component assignments, iteration tuples and the digest are all identical. This
proves stream-assignment invariance **and** accumulation-order invariance
together — the comparison is on the retained totals and the digest, not on stream
numbers alone.

---

## 13. Retained output

Exactly `total_nominal[iteration]` and `total_pv[iteration]`, in **original
iteration order**, as immutable tuples. The originals are never sorted; statistics
sort copies.

A guard walks the whole result tree: with five drivers and seven components over
1000 iterations, **exactly two** arrays of length 1000 exist, the longest sequence
anywhere is 1000, and no sequence of length 5000, 7000 or 10000 exists at any
depth. A 300-driver 100,000-iteration run retains 200,000 Doubles, not 30,000,000.

---

## 14. Result digest — D6-17

```
stream  ::= F_S("PCCM-RD") F_I(SIM_METHOD_VERSION) section
section ::= F_S("RESULT") F_I(record_count) record*
record  ::= F_I(3) F_I(iteration_index) F_N(total_nominal) F_N(total_pv)
```

1-based index, original iteration order, samples never sorted, equality exact.
The encoders are the accepted Phase-5 `calc_fingerprint` primitives — **there is
no competing hash**; this module builds a stream and hands it over.
`validate_result_digest_contract` asserts the contract still describes the
grammar this module implements, so a changed tag, section name, field count or
index origin cannot be silently ignored.

### All seven retained Step-0 vectors — exact on stream AND digest

| Case | `n` | Version | Digest |
|---|---|---|---|
| base | 5 | 1 | `3181AF89642DE500` |
| reversed_iteration_order | 5 | 1 | `4E0FEE211853E8F6` |
| nominal_and_pv_swapped | 5 | 1 | `63A0E93074F0C2EA` |
| one_iteration_dropped | 4 | 1 | `0CAC531732B88B2A` |
| one_ulp_perturbation | 5 | 1 | `5DC1A76B56D75EF4` |
| version_2 | 5 | 2 | `7E8D58C46CCDD798` |
| **empty** | 0 | 1 | `12ED977808313D71` |

The helper is standalone — it takes arrays, not a run — so the empty framing
vector is reproducible without an engine. A real run can never be empty: the
business minimum is at least 1000 iterations.

---

## 15. Replay and seed scope

**Same-runtime replay is exact.** Same prepared model, same iterations, same
effective seed, same `RNG_VERSION`, same `SIM_METHOD_VERSION` gives the same
nominal tuple, the same PV tuple and the same digest — re-running one prepared
object and preparing a fresh one both reproduce it. **No tolerance.**

**The withdrawn universal claim is not restored.**

*A — universal:* different accepted seeds give different initial RNG states.
Proved for `1, 2, 999, 12345, 2147483646`.

*B — fixture-scoped:* on a non-degenerate fixture whose uncertainty reaches the
retained total, different seeds give different digests.

| Seed | Digest |
|---|---|
| 1 | `700F2069316E208A` |
| 2 | `1B7A91177924F43B` |
| 999 | `3D6D7E4960275CD7` |
| 12345 | `41E528453ED7C07D` |

*And a fully degenerate fixture gives the SAME digest for every seed* — which is
**accepted behaviour**, not a defect. Seeds `1, 2, 999, 12345, 2147483646` all
produce total `507.0` and digest `2FEA5AA773992F12`. A model with no uncertainty
has nothing to vary.

---

## 16. Statistics

Scale-normalised, `n − 1` for the sample deviation, Type-7 percentiles by the
convex form, sorting on copies only. No NumPy, no `statistics` module, no
worksheet function.

The scale is **the largest power of two not exceeding `max(|x|)`**, so `x / scale`
is exact rather than costing an ulp per value before any statistic is computed.
Both forbidden paths are avoided: a naive `SUM x**2` overflows beyond about
`1.3e154`, and an unguarded Welford `delta = x - mean` overflows too.

| Sample | Mean | Sample SD (`n−1`) | Population SD (`n`) |
|---|---|---|---|
| `2,4,4,4,5,5,7,9` | **5.0** | `√(32/7) = 2.13809…` | `2.0` |
| `1,2,3,4` | **2.5** | `√(5/3) = 1.29099…` | `√1.25` |
| `10,10,10` | 10.0 | **0.0** | 0.0 |
| `−3, 3` | 0.0 | `√18 = 4.24264…` | 3.0 |

**The extreme case.** For `[−1.7e308, 1.7e308, 1.7e308, 1.7e308]`: the naive sum
is `inf`, the naive sum of squares is `inf`, and `x[0] − mean` is `−inf`. The
accepted helpers return mean `8.5e307` and SD `1.6999999999999997e308`.

`n < 2` **refuses explicitly** rather than inventing a deviation. A real run has
`N ≥ 1000`, so this only arises for a helper called directly.

**What scale safety is not.** Left-to-right accumulation of `n` terms carries the
usual `O(n·eps)` relative drift — about `1e-13` at `n = 1000` — and nothing here
re-associates or compensates, because the accepted accumulation primitive is the
one the contract names and a private summation algorithm would be a second
numerical authority. Scale safety is a statement about RANGE: a statistic whose
true value is representable is produced rather than refused.

### Type-7 hand vectors

`h = (n−1)p`, `lo = ⌊h⌋`, `hi = min(lo+1, n−1)`, `f = h − lo`,
`Px = (1−f)x[lo] + f·x[hi]`.

| `n` | `p` | `h` | Result |
|---|---|---|---|
| 1 | 0, 0.5, 1 | 0 | 5.0 |
| 2 | 0.25 | 0.25 | 12.5 |
| 2 | 0.5 | 0.5 | 15.0 |
| 3 | **0.5** | **1 (integral)** | 20.0 |
| 3 | 0.75 | 1.5 | 40.0 |
| 4 | **1/3** | **1 (integral)** | 2.0 |
| 4 | 0.9 | 2.7 | **3.7** |
| 4 | 0, 1 | 0, 3 | 1.0, 4.0 |
| 10 | 0.5 | 4.5 | 5.5 |
| 10 | 0.1 | 0.9 | 1.9 |
| 10 | 0.9 | 8.1 | 9.1 |

The convex form is the point: between `−1.7e308` and `1.7e308` the difference
form `x_lo + f(x_hi − x_lo)` overflows at every `f`, while the convex result is
bracketed by its endpoints and always exists.

---

## 17. The reported percentile ladder

Resolved from its owners, never restated: the selectable levels from
`input_contract.yaml` `config_tables.confidence_levels`, the fixed ones from
`sim_contract.yaml` `statistics.fixed_nonselectable_percentiles`. The contract's
`p10_selectable` flag must agree with the resolved sets or the ladder is refused.

**11 distinct percentiles:** P10 P50 P55 P60 P65 P70 P75 P80 P85 P90 P95.
Headline: **P10 P50 P70 P90**. P10 is reported and **non-selectable**.

Every stored percentile is available by label for **both** measures. Mixed
fixture, 2000 iterations, seed 12345:

| | Nominal | PV |
|---|---|---|
| P10 | 253.008137 | 253.008137 |
| P50 | 335.150948 | 335.150948 |
| P55 | 343.786527 | 343.786527 |
| P60 | 355.873234 | 355.873234 |
| P65 | 367.319481 | 367.319481 |
| P70 | 390.597558 | 390.597558 |
| P75 | 451.591189 | 451.591189 |
| P80 | 497.997573 | 497.997573 |
| P85 | 537.274495 | 537.274495 |
| P90 | 574.842721 | 574.842721 |
| P95 | 612.809254 | 612.809254 |

mean **376.327100**, sample SD **119.533383**, min 202.472477, max 792.868055,
digest `7F58EA884DAA8D65`.

*(PV equals nominal on this fixture because it has one applied project year at
the base year, so `Kpv = Knom = 1`. A separate two-year fixture separates the two
factors and confirms nominal uses `Knom` and PV uses `Kpv`.)*

---

## 18. Selected Confidence Level is reporting only

`run_simulation` never sees it. Changing the selection reruns no RNG, alters no
retained sample, no `result_digest`, no mean and no stored percentile — it chooses
among percentiles already computed. Every selectable level was applied to one
finished run and the digest, mean and full ladder were identical each time, while
the contingency differed.

Validated against `input_contract.yaml`. **P10 is refused as a selector** with a
message distinguishing "reported but not selectable" from "not a confidence
level", as are `P42`, `p50`, `P100`, `""` and `median`.

---

## 19. Contingency

```
contingency = selected Px total − deterministic base estimate A
```

per measure, through the accepted `safe_subtract`. The baseline is the Phase-5
`result.totals.a_nom` / `a_pv` — **not** the simulation mean, **not** the
analytical expected total, **not** `A + EMV`; all three are different numbers on
the mixed fixture and each is shown to give a different answer.

| Selected | Nominal | PV | Baseline A |
|---|---|---|---|
| P50 | 99.150948 | 99.150948 | 236.0 |
| P80 | 261.997573 | 261.997573 | 236.0 |
| P95 | 376.809254 | 376.809254 | 236.0 |

A percentile below A gives a **negative** contingency and it is **not clamped**.

---

## 20. Analytical expectation cross-check

Statistical evidence that the simulation targets the accepted Phase-5 analytical
expectation `E`. The allowance is **owned by the test** — four standard errors of
the mean — and is not a runtime tolerance, not a digest tolerance and not a field
of `sim_contract.yaml`. No exact equality between a finite Monte Carlo mean and
an expectation is asserted anywhere.

`N = 20000`, seed `20260825`:

| Fixture | Simulation mean | Analytical `E` | Difference | 4·SE allowance |
|---|---|---|---|---|
| Stochastic Cost Line (Triangular) | 219.923118 | 220.000000 | 0.076882 | 0.832785 |
| Stochastic Risk occurrence | 99.596030 | 100.000000 | 0.403970 | 3.773287 |
| Beta-PERT severity | 108.955058 | 108.333333 | 0.621724 | 3.261802 |
| Mixed costs + risks | 380.405893 | 380.500000 | 0.094107 | 3.437223 |

---

## 21. Numerical-domain pipeline

Sampler → contribution → accumulation → statistics → contingency, over the
accepted plan's families. Every result below is finite; the domain is **not
narrowed**.

| Family | Min | Max | Mean | SD | P90 |
|---|---|---|---|---|---|
| 1 large positive | 1.00024e307 | 1.49992e307 | 1.24546e307 | 1.43057e306 | 1.44037e307 |
| 2 large negative | −1.49976e307 | −1.00008e307 | −1.25454e307 | 1.43057e306 | −1.05963e307 |
| 3 crossing zero | −9.69082e306 | 9.81766e306 | −1.40843e305 | 4.0014e306 | 5.1163e306 |
| 4 opposite-sign accumulation | 1.5e308 | 1.5e308 | 1.5e308 | 3.23e294 | 1.5e308 |
| 5 subnormal scale | 9.88131e−324 | 9.99989e−321 | 4.91101e−321 | 2.86064e−321 | 8.80919e−321 |
| 6 degenerate | 42 | 42 | 42 | **0** | 42 |
| 7 Beta BC `m = a` | 0.036335 | 69.8958 | 16.5126 | 14.0908 | 37.7141 |
| 7 Beta BC `m = b` | 30.1042 | 99.9637 | 83.4874 | 14.0908 | 97.7423 |

**Family 4** is the one that matters most: three contributions of `+1.5e308`,
`+1.5e308`, `−1.5e308`. A naive left-to-right accumulation reaches `inf` on the
second term, and the accepted `safe_signed_sum` returns the representable total
`1.5e308` the model actually has.

**Families 8 and 9** are exercised on the statistics helpers directly — near-max
opposite-sign retained totals, and Type-7 interpolation between `−1.7e308` and
`1.7e308`.

**Family 10:** a selected `P90` of `1.0e308` against a base of `−5.0e307` gives
`1.5e308`; against a base of `−1.0e308` the difference does not exist and is
**refused, naming the contingency stage**, rather than returned as infinity.

No `inf`, `-inf` or `NaN` reaches a retained array, a statistic or a percentile.

---

## 22. Error semantics and no partial result

A failure at any iteration **raises**; no half-filled tuple is returned as
success. The refusal names the iteration index, the Permanent ID, the driver kind
and the numerical stage:

```
iteration 2: Cost Line 'CL-002': nominal contribution (faithful rescue):
the exact result is outside finite Double range
```

That fixture is deliberately one Phase 5 accepts — the stochastic line is
symmetric about zero, so `A`, `B`, `C` and `E` are all exactly zero and the model
is perfectly calculable — while roughly one sampled value in ten is
unrepresentable once multiplied by `Quantity`. The engine therefore completes
iteration 1 and fails at iteration 2: a design that returned what it had would
have a one-element result to offer. It returns nothing, deterministically, on
every attempt.

**The Phase-5 refusal hierarchy is preserved, not replaced.** An inverted
three-point triple and a Probability outside `[0,1]` are refused as
`ModelInputRefusal`; an unrepresentable contribution as `NumericalRangeRefusal`.
`SimOracleError` is used only for what the oracle itself owns — the iteration
count, the contract wiring, the confidence-level selector — and no implementation
invariant is converted into a user-input refusal.

---

## 23. Scope and static proof

Parsed from the source of `sim_oracle.py` and `sim_stats.py`, over **identifiers
actually used** — imports, attribute accesses, bound names, arguments and
definitions — with docstrings and message prose excluded. A module that documents
what it does not do has to be allowed to say the words; what matters is whether it
CALLS or TOUCHES anything so named.

| Guard | Absent |
|---|---|
| Workbook / COM | `Range` `ListObject(s)` `Application` `ThisWorkbook` `ActiveWorkbook` `Workbook(s)` `Worksheet(s)` `Cells` `Dispatch` `EnsureDispatch` `load_workbook` `openpyxl` `win32com` `comtypes` `pythoncom` `xlwings` |
| Evidence / emission | `evidence` `vba_source` `stage_b_emit` `emit_stage_b` `workbook_builder` `build_workbook` `calc_emit` `gate_b_inspection` `numpy` `random` `statistics` |
| Publication | `_SimData` `SimData` `write_sim_data` `publish` `persist` `next_auto_nonce` `run_id` `allocate_run_id` `simulation_status` `attempt_metadata` `PCCM_RunSimulation` `write_text` `write_bytes` `save` |
| Out of scope | `sensitivity` `tornado` `annual` `annual_samples` `annual_matrix` `sample_matrix` `per_driver_samples` `driver_samples` |

Static identifier scanning cannot see a name assembled at run time, so it is
paired with a **runtime** guard: a complete run executes with `open` replaced by a
function that raises, and opens nothing at all.

**No module-level mutable simulation state.** Every module-level binding in
`sim_oracle.py` is a constant or a compiled pattern, and neither module exposes a
module-level `list`, `dict` or `set`.

The prepared model is checked at run time too: every field of every prepared
driver is a `builtins`, `sim_rng` or `sim_sample` type, and none carries a
`Range` or `ListObjects` attribute.

---

## 24. Test inventory

`tests/test_phase6_sim_oracle.py` — **74 conformance tests**

| Plan letter | Coverage |
|---|---|
| A | degenerate Cost Line; Uniform ML ignored end to end |
| B | stochastic Cost Line spans its support, one uniform per iteration |
| C | Quantity linear in `Q`; Quantity outside the distribution (sample-for-sample scaling) |
| D, E, F | `p = 0`, `p = 1`, intermediate rate |
| G | severity invoked every Risk iteration (rare-risk fixture) |
| H | degenerate severity: zero consumption, unchanged stream, still invoked |
| I | probability-only severity-sequence invariance; streams never shared |
| J | mixed 3 Cost + 2 Risk end to end |
| K | no-Beta model: one uniform per component per iteration |
| L | exact-friendly fixture reproduced by hand arithmetic |
| M | Beta-containing model: even consumption, at least one attempt per iteration |
| N | same-seed exact replay, object and fresh preparation |
| O | universal stream divergence; fixture-scoped digest divergence |
| P | degenerate cross-seed equal digest is accepted |
| Q | three physical reorderings, exact equality |
| R | non-associative fixtures, canonical order followed |
| S | all seven Step-0 digest vectors, stream and digest |
| T | Type-7 hand vectors for `n = 1, 2, 3, 4, 10` |
| U | mean and sample SD hand vectors; scale safety; `n < 2` refusal |
| V | full 11-value ladder, monotone, both measures |
| W | Selected CL reporting-only; P10 non-selectable |
| X | contingency baseline `A`; negative contingency unclamped |
| Y | ten extreme-domain families |
| Z | no file access at run time; no evidence or workbook import |
| — | iteration pre-flight; stream initialisation; retained shape; error semantics; analytical cross-check; package exports |

`tests/test_phase6_sim_oracle_validation.py` — **37 mutation controls and guards**

The instrument is a local re-implementation of one iteration, parameterised by
exactly one mutation. **Run unmutated it reproduces the accepted engine bit for
bit**, and that equality is asserted first — it is what makes every other control
non-vacuous.

| # | Mutation | Detected by |
|---|---|---|
| 2 | Cost sample read as TOTAL, not UNIT | recovered samples lie in `[80,150]`, not `[320,600]` |
| 3 | Quantity omitted | every iteration differs |
| 4 | Quantity applied twice | every iteration differs |
| 5 | Quantity introduced on a Risk | every occurrence scales by 3 |
| 6 | Probability folded into `Knom`/`Kpv` | the run stops containing non-occurrences |
| 7 | PV derived from nominal | two drivers with different `Kpv/Knom` ratios |
| 8 | severity sampled only on occurrence | totals differ; severity consumption is 1500, not ~450 |
| 9 | occurrence and severity streams merged | totals differ; indices and states differ |
| 10 | the comparability D6-18b buys | accepted states equal, mutated states differ |
| 11 | physical row order | 2.0 versus canonical 1.0 |
| 12 | Risks accumulated before Costs | 1.0 versus canonical 0.0 |
| 13 | accumulation reversed | 0.0 versus canonical 1.0 |
| 14 | *(the honest half)* reversal on an ordinary model | changes nothing — recorded, not hidden |
| 15 | built-in `sum` for the accepted accumulation | `inf` versus `1.5e308`, **only** on the constructed fixture |
| 17 | digest iteration index removed | different digest and stream |
| 18 | digest nominal/PV swapped | the retained swapped vector |
| 19 | digest version omitted / changed | different digest; version is load-bearing |
| 20, 21 | samples sorted before the digest | different digest, on the vector and on a real run |
| 22 | nearest-rank percentile | differs from Type-7 at `p = 0.25, 0.4, 0.9` |
| 23 | Type-7 unsafe difference interpolation | `inf` versus a finite convex result |
| 24 | population SD for sample SD | `2.0` versus `√(32/7)` |
| 25 | naive sum-of-squares SD | `inf` versus `1.7e308` |
| 26 | naive mean accumulation | `inf` versus `8.5e307` |
| 27 | unguarded Welford deviation | `−inf` versus a finite SD |
| 28 | contingency baseline changed from `A` | mean, `E` and `A + EMV` each give a different answer |
| 29 | contingency clamped | a negative contingency is reported |
| 30 | Selected CL inserted into execution | one digest accepted, ten under the leak |
| 31–37 | scope, publication, sensitivity, matrix, prepared-model and module-state guards | see §23 |

---

## 25. Regression

| Check | Result |
|---|---|
| Full Python suite | **2209 passed, 0 failed** |
| Before Step 4 | 2094 passed |
| New Step-4 tests | **+111** (74 conformance + 37 controls) |
| New Step-3 regressions | **+4** (the §0 invalid-state cleanup) |
| | 2094 + 111 + 4 = **2209** |
| Stage-A verifier / build | **351 passed, 0 failed** — unchanged, Step 4 emits nothing |
| Step-1 contract tests | green |
| Step-2 RNG and jump tests | green |
| Step-3 sampler tests | green, every vector unchanged |

No test was deleted, skipped or weakened.

### Baseline preservation

Byte-identical to the accepted Step-3 commit `e939d47`: `spec/` (whole
directory), `src/`, `bootstrap/`, `evidence/`, `docs/phase6_plan.md`,
`docs/phase6_step0.md`, `docs/phase6_step1.md`, `docs/phase6_step2.md`,
`docs/phase6_step3.md`, `builder/build_stage_a.py`, `calc_oracle.py`,
`calc_numeric.py`, `calc_fingerprint.py`.

### Phase-5 fingerprint vectors

```
fingerprint("PCCM-FP")      6551C6F365DA7F3F
fingerprint_probe(A|B)      42E49DC715F06970
fingerprint_probe(AB|)      7558FD9248656EAD
canonical_number(1/3)       3.3333333333333331E-01
```

Unchanged.

---

## 26. Runtime evidence

**No Windows or Excel execution was performed.** Everything in this record is
Linux Python 3.11. Step 4 emits nothing into the workbook and creates no VBA, so
there is no Stage-B surface to exercise at this step.

---

## 27. GATE-B TEMP-DIR CLEANUP DEBT — OPEN

**Status: OPEN. It MUST be resolved before the Phase-6 Gate-B harness extension
/ Windows execution step.**

Repeated test execution leaves `pccm-gateb-*` temporary directories under `/tmp`.
They are never removed, so they accumulate across runs. In the incident observed
during Step-3 execution there were approximately **56,986 directories occupying
about 30 GB**, which exhausted the session's writable allowance and truncated a
gitignored build artifact (`build/phase5_cases.json`) to zero bytes mid-write.
Two Phase-5 tests then failed on the truncated JSON.

The leak rate was measured again during this step: **one full suite run leaves
385 directories**, roughly 220 MB. That is unremarkable once and cumulative
forever — about 150 full runs reach the same exhaustion.

Independent review accepted that this was an **environment / harness** issue and
**not a Step-3 sampler defect**, and the transient `ENOSPC` failures are not
product-code failures. Regenerating the artifact restored a fully green suite
with no source change.

The helper that creates those directories is **outside the authorised Step-4
oracle scope and was not changed here.** It is recorded so it is not forgotten:
the Gate-B harness must clean up its own temporary directories before it is
extended for Phase 6, or the same exhaustion will recur on the Windows machine
where the failure mode is less obvious and the disk is not disposable.

---

**STEP 4 — ACCEPTANCE REQUESTED**
