# PCCM Phase 6 — Step 3 authority record

Step 3 implements the pure Python **stochastic transforms**: Uniform,
Triangular, Beta-PERT through the exact locked Cheng BB/BC formulation, and
Bernoulli occurrence.

**No Monte Carlo.** No iteration engine, no Cost Line or Risk contribution, no
`Quantity`, `Knom`, `Kpv`, totals, `_SimData`, `result_digest`, statistics or
contingency. Step 4 pairs occurrence with unconditional severity; Step 3 provides
the pieces.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Phase-6 Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Phase-6 Step-1 contract authority | `7d8e73e59347b5d3acb3f80d848be6aaf1d3eb83` |
| Accepted Phase-6 Step-2 RNG reference | `9bdf5a117784e99397dd80efee07d20df508ac9c` |
| Step-3 samplers | this commit — reported by the delivery message and `PROVENANCE.txt` |

**No contract defect was found**, so nothing in `spec/` changed.

---

## 2. Files

| File | Change |
|---|---|
| `builder/pccm_builder/sim_sample.py` | **NEW** — the samplers, 482 lines |
| `builder/pccm_builder/sim_rng.py` | the authorised `uniforms(state, 0)` cleanup **only** |
| `builder/pccm_builder/__init__.py` | exports the Step-3 surface |
| `tests/test_phase6_sim_sample.py` | **NEW** — 43 conformance tests |
| `tests/test_phase6_sim_sample_validation.py` | **NEW** — 31 mutation controls |
| `tests/test_phase6_sim_rng.py` | **+1** — the zero-count regression |
| `docs/phase6_step3.md` | **NEW** — this record |

### The accepted Step-2 cleanup

`RngReference.uniforms(state, count)` now validates the incoming state **once,
up front, including at `count == 0`**. Previously the zero-draw case returned
whatever it was handed, because the loop that would have validated never ran — a
defensive hole that only widens once samplers pass states through this API. No
other Step-2 behaviour changed, and every retained RNG and jump vector still
passes.

---

## 3. Public reference surface

```python
SampleResult(value, state, uniforms_consumed, proposal_attempts)
BernoulliResult(occurred, state, uniform, uniforms_consumed)
PreparedBetaPert(...)                     # per-driver shape constants

sample_distribution(reference, state, family, a, m, b) -> SampleResult
sample_uniform(reference, state, a, b, m=None)         -> SampleResult
sample_triangular(reference, state, a, m, b)           -> SampleResult
sample_beta_pert(reference, state, a, m, b)            -> SampleResult
prepare_beta_pert(a, m, b)                             -> PreparedBetaPert
sample_prepared_beta(reference, state, shape)          -> SampleResult
bernoulli_occurs(reference, state, probability)        -> BernoulliResult
is_degenerate(family, a, m, b)                         -> bool
```

**Consumption is observable, not hidden.** Every sampler returns the state it
produced plus how many uniforms it consumed and, for the rejection sampler, how
many proposal attempts it made — because under an acceptance/rejection sampler
the consumption is not a fixed property of the call, and a later implementation
must reproduce it, not merely reproduce the value.

**`prepare_beta_pert` draws nothing** and holds no RNG state. `α` and `β` are
fixed for the life of a driver, so a simulation sampling one driver 100,000 times
computes the square root and the setup logarithms **once**.

---

## 4. Uniform

| `a` | `m` | `b` | Value | Uniforms |
|---|---|---|---|---|
| `0.0` | — | `100.0` | `12.701112204657713` | 1 |
| `-50.0` | — | `50.0` | `-37.29888779534228` | 1 |
| `100.0` | **`500.0`** | `100.0` | `100.0` | **0** |
| `-1e308` | — | `1e308` | `-7.459777559068457e+307` | 1 |

**Most Likely is ignored, whatever it holds** — value, state and consumption are
byte-identical across `None`, `0`, `50`, `±1e300`, and both endpoints, through
both the direct and the dispatched call.

The transform is the **convex form** `x = (1−u)a + ub`. `a + u(b−a)` is not
implemented: at `a = −MAX, b = +MAX` the difference **overflows** while every
convex result is finite, and `test_01` in the mutation suite asserts exactly that
contrast.

---

## 5. Triangular

Conditioned on `s = max(|a|,|m|,|b|)`, evaluated in the conditioned space and
rescaled after. One uniform, always.

| Case | Result |
|---|---|
| `a=0, m=30, b=100`, `u = 0.001` | `1.7320508075688772` |
| **`u = c = 0.3`** | **`30.0`** — the branch point lands on the mode |
| `u = 0.5` | `40.83920216900384` |
| `u = 0.999` | `97.35424868893541` |
| **`m = a`**, `u = 0.5` | `29.28932188134524` — `c = 0`, upper branch |
| **`m = b`**, `u = 0.5` | `70.71067811865476` — `c = 1`, lower branch |
| `a=−1e308, m=0, b=1e308`, `u = 0.25` | `−2.9289321881345245e+307` — finite |

`u ≤ c` takes the lower branch; `test_05` in the mutation suite shows `<`
produces a different value at `u = c` exactly.

---

## 6. Beta-PERT parameterisation and dispatch

| `r` | `α` | `β` | `α+β` | Dispatch |
|---|---|---|---|---|
| 0.00 | 1.0 | 5.0 | 6.0 | **BC** |
| 0.25 | 2.0 | 4.0 | 6.0 | BB |
| 0.50 | 3.0 | 3.0 | 6.0 | BB |
| 0.75 | 4.0 | 2.0 | 6.0 | BB |
| 1.00 | 5.0 | 1.0 | 6.0 | **BC** |

`λ = 4`, `r` computed in the conditioned space. **Equality belongs to BC**, so
`m = a` and `m = b` reach BC *by the rule* — not as endpoints special-cased with
another sampler. BB orients `(min, max)`; BC orients `(max, min)` — **opposite**,
and inverting one returns a valid Beta variate of the **mirrored** distribution.

---

## 7. Cheng vector conformance — all exact

Every one of the five retained cases, all 24 samples each: **value, proposal
attempts, uniforms consumed, cumulative uniforms and the post-sample six-word
state**.

| Case | Dispatch | Samples exact | Attempts | Uniforms | Final state |
|---|---|---|---|---|---|
| BB interior 2/4 | BB | **24 / 24** | 28 | 56 | ✔ |
| BB symmetric 3/3 | BB | **24 / 24** | 27 | 54 | ✔ |
| BB near-boundary 1.04/4.96 | BB | **24 / 24** | 35 | 70 | ✔ |
| BC α=1 β=5 | BC | **24 / 24** | 31 | 62 | ✔ |
| BC α=5 β=1 | BC | **24 / 24** | 30 | 60 | ✔ |

Every literal is a **literal**: `1.3862944` is those eight digits and is not
evaluated as `log(4)`; `2.609438` is not `1 + log(5)`; `0.0138889`, `0.0416667`
and `0.777778` are not `1/72`, `3/72` and `7/9`.

**A rejected proposal consumes both uniforms and the retry continues from the
resulting state. There is no rewind** — `test_20` proves the post-state equals
`2 × attempts` draws forward, not two.

---

## 8. Bernoulli

| `p` | Rate over 20,000 draws |
|---|---|
| 0.0 | **0.0000** |
| 0.1 | 0.0993 |
| 0.5 | 0.4972 |
| 0.9 | 0.8990 |
| 1.0 | **1.0000** |

`occurred = u < Probability`, **strict**. `p = 0` never occurs and `p = 1` always
occurs **exactly**, with no special case anywhere — because raw MRG output is
strictly inside `(0,1)`, strictness alone carries both. `u == p` does **not**
occur; the neighbours either side are tested.

A probability outside `[0,1]` is **refused, not clamped**.

---

## 9. Consumption — the contract, verified

| | Uniforms |
|---|---|
| Uniform, non-degenerate | **1** |
| Triangular, non-degenerate | **1** |
| Beta-PERT, non-degenerate | **2 × proposal_attempts** |
| Bernoulli | **1** |
| **Any degenerate distribution** | **0**, state unchanged |

Degeneracy is **family-specific**, per the accepted post-Step-1 correction:
Uniform is degenerate iff `a == b` (Most Likely unread); Triangular and Beta-PERT
iff `a == m == b`. A degenerate Beta forms **no `r`**, so `0/0` cannot arise.

---

## 10. Numerical domain

Negative supports, zero-crossing supports, near-`Double`-maximum and
subnormal-scale endpoints all sample finitely. **No positivity rule and no
magnitude cap.** A representable result is never refused because a naive
intermediate would have overflowed — that is what the convex and conditioned
forms are for. A result that genuinely cannot be represented raises
`SimSampleError` **naming the family and the numerical stage**; nothing silently
returns `inf` or `NaN`, and nothing clips.

Refused, never repaired: unknown family · misordered endpoints · missing Most
Likely where required · non-finite parameters · probability outside `[0,1]` ·
invalid RNG state. **A populated Uniform Most Likely is ignored, not
order-validated** — it is not a mode.

---

## 11. Test inventory

| Suite | Tests | Kind |
|---|---|---|
| `test_phase6_sim_sample.py` | **43** | conformance, transforms, vectors, consumption, scope |
| `test_phase6_sim_sample_validation.py` | **31** | **mutation controls** |
| `test_phase6_sim_rng.py` | **+1** | the Step-2 zero-count regression |
| **Total new** | **75** | |

Coverage against the required inventory: **A** three families dispatch, unknown
refused · **B** family-specific degeneracy in all three · **C** ignored-ML
invariance across seven values and both call paths · **D** one-draw consumption ·
**E** branch point and both neighbours · **F** `m=a` / `m=b` · **G** `α+β = 6`
across nine shapes · **H**/**I** BB and BC dispatch including the boundary ·
**J** all five Cheng cases, 120 samples · **K** exact attempt and draw counts ·
**L** post-sample states · **M** negative and zero-crossing supports · **N**
extreme-domain and subnormal rescaling · **O** `p=0`/`p=1` and strictness ·
**P** parameter refusal · **Q** no evidence dependency · **R** no
`random`/`secrets`/NumPy · **S** no Monte Carlo.

### Mutation controls

Every required control, plus seven more: unsafe Uniform difference form · Uniform
ML used · Uniform degeneracy reverted to `a==m==b` · misordered Uniform ·
Triangular `<` for `<=` · swapped branches · conditioning removed · misordered
Triangular · changed `λ` · relaxed dispatch boundary · **`log(4)` for the
literal** · **`1+log(5)` for the literal** · changed squeeze operator · changed
squeeze coefficient · locked logit replaced · changed BC literals · inverted
orientation · flipped return orientation · rejected proposal failing to advance ·
one uniform per proposal · degenerate driver that consumes · unsafe Beta rescale ·
positivity restriction · non-strict Bernoulli · broken `p=0`/`p=1` · clamped
probability · fourth family · non-finite result returned.

`test_00` asserts the **unmutated** sampler matches every retained case, and
`test_13` asserts the mutation harness itself reproduces the locked corpus —
without both, every rejection would prove nothing.

### One control that had to be constructed, and why

**The squeeze literals are invisible in the retained corpus.** Replacing
`1.3862944` with `log(4)` changes **nothing** across all 72 BB samples — and
that is exactly what Step 0 measured: over 805,837 predicate evaluations the
closest relative margin to a boundary was `7.6e-07`, while the literal gap is
`3.9e-08`. A short corpus cannot flip a squeeze decision by luck.

So those controls use a **constructed witness** — a deterministic `(u1, u2)`
placing `5z` between the two thresholds — exactly as Step-0 control 18b does.
`test_14` asserts both halves: the witness flips the decision, **and** the corpus
does not, so a control that only re-ran 24 samples would have passed the mutant
and reported success. `test_17b` shows the contrast: reversing the squeeze
direction *is* visible in the corpus.

---

## 12. Scope discipline

| Claim | How it is proven |
|---|---|
| no evidence read at run time | AST scan: no import names it and **no file-access call exists at all** (`test_39`) |
| no `random`, `secrets`, NumPy, SciPy | AST import scan plus a source scan for `default_rng`, `Randomize`, `Rnd(` (`test_40`) |
| no Monte Carlo or contribution | no function name contains `iterate`, `run_simulation`, `simulate`, `contribution`, `accumulate`, `percentile`, `result_digest`, `contingency`; **no identifier binds** `Quantity`, `Knom`, `Kpv`, `total_nominal` or `total_pv`; no `for iteration` (`test_41`) |
| no global mutable state | no `global`/`nonlocal`, no lowercase module-level binding (`test_42`) |
| no Phase-6 VBA | no `modSimRng.bas` / `modSimSample.bas` / `modSimEngine.bas`, and no `.bas` mentions `Cheng` or `MRG32k3a` (`test_43`) |

Every uniform comes from the accepted `RngReference`. The injected-`u` helpers
(`_uniform_from_u`, `_triangular_from_u`, `_bernoulli_from_u`) are private, pure
and **consume nothing** — the caller supplies `u`, so they are not a second RNG.
Cheng conformance uses **actual deterministic stream draws**, because
rejection and consumption behaviour is part of the authority.

---

## 13. Regression

| Check | Result |
|---|---|
| Full Python suite | **2094 passed, 0 failed** |
| Before Step 3 | 2019 passed |
| New Step-3 tests | **+75** (43 + 31 + 1) |
| Stage-A verifier / build | **351 passed, 0 failed** — unchanged, Step 3 emits nothing |
| Step-1 contract tests | green |
| Step-2 RNG and jump tests | green, all vectors unchanged |

No test was deleted, skipped or weakened.

### Baseline preservation

Byte-identical to the accepted Step-2 reference `9bdf5a1`:

| Path | `git diff 9bdf5a1` |
|---|---|
| `pccm/spec` | **EMPTY** — the whole directory |
| `pccm/src` | **EMPTY** |
| `pccm/bootstrap` | **EMPTY** |
| `pccm/evidence` | **EMPTY** |
| `pccm/docs/phase6_plan.md`, `_step0`, `_step1`, `_step2` | **EMPTY** |
| `pccm/builder/build_stage_a.py` | **EMPTY** |

Phase-5 fingerprint vectors, recomputed live:

```
fingerprint("PCCM-FP")     6551C6F365DA7F3F
fingerprint_probe(A|B)     42E49DC715F06970
fingerprint_probe(AB|)     7558FD9248656EAD
canonical_number(1/3)      3.3333333333333331E-01
```

`test_phase5_fingerprint.py` — **52 passed**.

---

## 14. Step-3 acceptance gate

| Requirement | Status |
|---|---|
| Uniform correct; ignored ML cannot affect it | yes — §4 |
| Triangular correct at all boundaries | yes — §5 |
| Beta-PERT parameterisation correct | yes — §6 |
| exact locked Cheng BB/BC vectors pass | yes — 120/120 samples, §7 |
| RNG consumption counts match exactly | yes — §9 |
| degenerate distributions consume zero | yes — §9 |
| Bernoulli strict comparison correct | yes — §8 |
| extreme-domain cases finite where representable | yes — §10 |
| no sampler reads evidence at run time | yes — §12 |
| no Monte Carlo engine | yes — §12 |
| no spec change | yes — `pccm/spec` diff empty |
| no VBA | yes — §12 |
| no Windows/Excel runtime | yes — Linux only |
| full regression green | yes — §13 |

---

**STEP 3 — ACCEPTANCE REQUESTED**
