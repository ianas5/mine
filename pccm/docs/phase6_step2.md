# PCCM Phase 6 — Step 2 authority record

Step 2 is the **first Phase-6 executable step**. It implements the pure Python
reference for the deterministic random-number backbone: seeding, one exact
MRG32k3a step, the canonical uniform, the `2^127` stream jump, and canonical
component-stream assignment.

**Nothing samples.** There is no Uniform, Triangular or Beta-PERT sampler, no
Bernoulli trial, no iteration, no statistic, no digest — and not even the Uniform
*distribution* transform `x = (1−u)a + ub`, which belongs to Step 3. Step 2
produces raw MRG uniforms and stream identities.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Phase-6 Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Phase-6 Step-1 contract authority | `7d8e73e59347b5d3acb3f80d848be6aaf1d3eb83` |
| Step-2 reference | this commit — reported by the delivery message and `PROVENANCE.txt` |

**Implementation authority is `spec/sim_contract.yaml`**, plus the seed
admissibility owned by `spec/input_contract.yaml`. The retained Step-0 vectors
are a **test and provenance authority only**.

**No authority conflict was found.** §12 of the authorisation warned that the
contract's `(component_kind, permanent_id, role)` must not be read as three
global blocks; it is not, and §6 below sets out why the reading is consistent
rather than chosen.

---

## 2. Files

| File | Change |
|---|---|
| `builder/pccm_builder/sim_rng.py` | **NEW** — the reference, 487 lines |
| `builder/pccm_builder/__init__.py` | exports the Step-2 surface |
| `tests/test_phase6_sim_rng.py` | **NEW** — 37 conformance tests |
| `tests/test_phase6_sim_rng_validation.py` | **NEW** — 27 mutation controls |
| `docs/phase6_step2.md` | **NEW** — this record |

Nothing under `spec/`, `src/`, `bootstrap/`, `evidence/` or the earlier phase
records changed, and `build_stage_a.py` is untouched.

---

## 3. Public reference surface

```python
RngState(words)                 # six exact ints, oldest-first, immutable
Draw(state, uniform)            # the state AFTER the advance, and its uniform
Component(kind, permanent_id, role)

RngReference.from_contracts(sim, inputs) -> RngReference
    .validate_state(state)              -> RngState
    .fixed_seed_to_state(seed)          -> RngState
    .auto_seed_from_nonce(nonce)        -> int
    .next_uniform(state)                -> Draw
    .uniforms(state, count)             -> (tuple[float], RngState)
    .jump_to_next_stream(state)         -> RngState
    .stream_initial_state(base, k)      -> RngState
    .canonical_sort_key(component)      -> sort key
    .components_for(cost_ids, risk_ids) -> tuple[Component, ...]
    .assign_component_streams(components)      -> ((Component, index), ...)
    .component_stream_states(base, components) -> ((Component, index, RngState), ...)
```

**No global mutable state, no singleton, no hidden seeding.** `RngState` and
`RngReference` are frozen; advancing returns a **new** state, so a caller cannot
accidentally share one generator between two components, and a test can hold a
state across a jump without it moving underneath. `test_37` asserts the module
declares no `global` or `nonlocal` and binds nothing mutable at module scope.

**Every operational constant is derived from the contract**, and the seed domain
is read from `input_contract.yaml`, its owner. There is no second copy that could
drift.

---

## 4. Why the oracle uses exact Python integers

The reference arithmetic is arbitrary-precision integer arithmetic **including
the jump**, and `MultModM` deliberately does not appear.

> Python exact integers are the **oracle**.
> The VBA-safe Double / `MultModM` path is a **later implementation** that must
> prove itself against this oracle.

Writing the oracle in the same restricted arithmetic the implementation will use
would mean the two agree because they share a technique — which proves nothing
about either.

---

## 5. Vector results

### Seeding and the recurrence

| Seed | State | First uniform |
|---|---|---|
| `1` | `[1] × 6` | `0.0003395772237870988` |
| `2` | `[2] × 6` | `0.0006738822302239724` |
| `12345` | `[12345] × 6` | **`0.12701112204657714`** |
| `2147483646` | `[2147483646] × 6` | `0.7564584266728147` |

```
seed 12345, first uniform  0.12701112204657714   float.hex 0x1.041e683b58b4bp-3
seed 12345, state after 20 [1251035728, 906640697, 235742957, 1104343287, 1050843907, 3914401992]
```

All 20 retained uniforms and all four per-seed vectors reproduce **exactly**.
Comparison is through `float.hex()` — bit pattern identity, not a tolerance. **No
approximate comparison belongs to an RNG backbone**, and none appears here.

### AUTO nonce → effective seed

| Nonce | Effective seed |
|---|---|
| `0` | `1` |
| `1` | `48271` |
| `2` | `182605794` |
| `3` | `1291394886` |
| `10` | `1596680831` |
| `1000` | `429183498` |
| `2147483645` | `1899818559` |
| `2147483646` | **refused — exhausted** |

Computed by exact modular exponentiation, `O(log nonce)`. `test_09` checks
agreement with stepped multiplication over the first 3,000 nonces **and**
evaluates nonce `900,000,000`, where stepping would need nine hundred million
multiplications. The function is **pure**: it persists nothing and increments
nothing; the transactional nonce lifecycle belongs to the later engine boundary.

### Jump vectors — all exact

| Stream | Initial state | Match |
|---|---|---|
| 0 | `12345, 12345, 12345, 12345, 12345, 12345` | ✔ |
| 1 | `3692455944, 1366884236, 2968912127, 335948734, 4161675175, 475798818` | ✔ |
| 7 | `3281794178, 2616230133, 1457051261, 2762791137, 2480527362, 2282316169` | ✔ |
| 399 | `2260181002, 1948664812, 612976419, 1919355493, 2890171896, 2701138777` | ✔ |
| 401 | `215541976, 1807926449, 2979430890, 2228004365, 3803991720, 370726289` | ✔ |

Stream 1 is the **published RngStreams second-stream state** for the `12345`
default, which is what settles the orientation. The first five uniforms and the
state after five draws are reproduced exactly for every retained stream.

**Orientation, stated because getting it wrong is silent.** State is stored and
exposed **oldest-first**; the matrices operate on **newest-first** triples, so
each triple is reversed on the way in and back on the way out. No transpose was
guessed and no remembered literature was used — the retained vectors are the
authority, and three mutation controls (`test_08`, `test_09`, `test_10`) prove a
transpose, a dropped reversal and a half-dropped reversal all fail.

### Stream assignment — all exact

```
total components 400          (200 Cost Lines + 2 × 100 Risks)
   0  COST  CL-001  value
   1  COST  CL-002  value
   …
 396  RISK  R-099   occurrence
 397  RISK  R-099   severity
 398  RISK  R-100   occurrence
 399  RISK  R-100   severity
```

---

## 6. `component_kind` is the driver-kind axis, and `role` is separate

The authorisation warned against reading `COST_SAMPLE`, `RISK_OCCURRENCE` and
`RISK_SEVERITY` as one ordered axis, which would produce three global blocks with
every occurrence stream before every severity stream. It does not, and the
contract supports the correct reading without straining:

- `components.kinds[]` gives each component key a **`driver_kind`** and a
  **`role`** as two separate fields — so the contract already distinguishes the
  axes;
- `stream_assignment.sort_keys` lists `component_kind` and `role` as **two of
  three** keys, with `permanent_id` between them, so they cannot be the same
  axis;
- the retained Step-0 vectors interleave per Risk.

The reference therefore ranks the **driver-kind** axis (`COST` before `RISK`),
then the Permanent ID, then the role. The kind axis is derived from the declared
order of `components.kinds`, cross-checked against
`accumulation.driver_kind_order`; **a contract stating those two orders
differently is refused** (`test_24`), because one of them would otherwise be
silently unused.

`test_21` proves the rejected reading differs: under global blocks
`R-100 occurrence` would be stream **299**; the retained vector says **398**.

---

## 7. Ordinal UTF-16, implemented as authority

The Permanent-ID key is the **accepted Phase-5 `utf16_sort_key`**, reused rather
than reimplemented, so no second collation authority exists. Python's own string
ordering compares **code points** and is not the definition — today's IDs are
ASCII so the two coincide, but `test_30` exercises an astral character where they
**disagree** and asserts the UTF-16 order wins.

No locale collation, no case folding, no numeric-suffix sorting. The accepted
non-numeric consequence is preserved: **`CL-1000` sorts before `CL-999`**.

**Row order cannot reach the assignment** — it is a sort, not a position.
`test_28` shuffles the design-target component set five times with a seeded
shuffle and asserts the mapping is identical.

**Adding a driver may shift later streams**, and `test_32` asserts it does rather
than trying to prevent it. That is D6-16 Family A's stated consequence.

---

## 8. Test inventory

| Suite | Tests | Kind |
|---|---|---|
| `test_phase6_sim_rng.py` | **37** | conformance against the retained vectors, exact-double comparison, scope discipline |
| `test_phase6_sim_rng_validation.py` | **27** | **mutation controls** |
| **Total new** | **64** | |

Coverage against the required inventory: **A** four fixed seeds · **B** ten
invalid seeds including `True`/`False` · **C** AUTO vectors and both boundaries ·
**D** the repeated-scalar mapping · **E** the first 20 uniforms · **F** exact
state after retained draw counts · **G** `(0,1)` over 2,000 draws · **H** streams
0/1/7/399/401 · **I** orientation and the matrix boundary · **J** exact component
assignments · **K** row-order invariance · **L** ordinal UTF-16 and non-numeric
order · **M** no evidence dependency · **N** no `random`/`secrets`/NumPy.

### Mutation controls

Every required control, plus five more:

| Control | Result |
|---|---|
| MRG modulus changed (`m1`, `m2`, swapped) | caught |
| each recurrence coefficient changed | caught |
| `norm` perturbed by one ULP | caught |
| `p1`/`p2` state shift swapped | caught |
| `<` substituted for the accepted `<=` boundary | caught |
| all-zero state accepted | refused |
| **each of the 18 jump-matrix elements** changed | caught |
| jump matrix transposed | caught |
| matrix-boundary reversal dropped | caught |
| reversal dropped on one side only | caught |
| the two jump matrices swapped | caught |
| AUTO multiplier changed | caught |
| AUTO modulus changed | caught |
| sequential stepping substituted for the power | caught — see below |
| seed domain widened | caught |
| a mixer substituted for the repeated scalar | caught |
| stream index origin changed from 0 | caught |
| physical row order used | caught |
| numeric-ID sorting used | caught |
| locale / case-folding order used | caught |
| **global occurrence/severity blocks** | caught |
| severity before occurrence within a Risk | caught |
| Risks before Cost Lines | caught |
| contract with two disagreeing kind orderings | refused |
| duplicate component | refused |
| unknown kind or role | refused |

**Stepping substituted for the power is the interesting one.** Unbounded stepping
*agrees* with the power — that is exactly why it is a trap. The way it fails in
practice is by being bounded, so `test_14` asserts unbounded stepping agrees at
nonce 5,000,000 and that a loop capped at 1,000, 100,000 or 4,999,999 silently
returns the **wrong** seed. The authority is a power precisely so that cannot
arise.

**Mutations live in the tests.** Every variant is built by replacing a field of
the frozen `RngReference` or by reimplementing one step locally as the defect
would have written it. `spec/sim_contract.yaml` was never edited to manufacture a
failure.

`test_00` asserts the **unmutated** reference matches every retained corpus —
without it, every rejection below would prove nothing.

---

## 9. Scope discipline — what does not exist

| Claim | How it is proven |
|---|---|
| the reference reads no evidence at run time | AST scan: no import names it, and **no file-access call of any kind exists** — it opens nothing (`test_33`) |
| no `random`, `secrets`, NumPy or SciPy | AST import scan plus a source scan for `default_rng`, `Randomize`, `Rnd(` (`test_34`) |
| no sampler and no simulation | no function name contains `sample`, `cheng`, `bernoulli`, `occurrence`, `severity`, `iterate`, `run_simulation`, `percentile`, `result_digest` or `contingency`; and not even `(1-u)*a`, `Quantity`, `Knom` or `Kpv` appears (`test_35`) |
| no Phase-6 VBA | no `modSimRng.bas` / `modSimEngine.bas` / `modSimReport.bas`, and no `.bas` mentions MRG32k3a (`test_36`) |
| no global mutable state | no `global`/`nonlocal`, no lowercase module-level binding (`test_37`) |
| no arithmetic seek to an iteration | a jump moves `2^127` draws, not `k`; `test_23` asserts a jump differs from a step and that no seek method exists |

No workbook read, no Excel, no `Range`, no `ListObject`, no COM, no Windows.

---

## 10. Regression

| Check | Result |
|---|---|
| Full Python suite | **2019 passed, 0 failed** (238.32 s) |
| Before Step 2 | 1955 passed |
| New Step-2 tests | **+64** (37 + 27) |
| Stage-A verifier / build | **351 passed, 0 failed** — unchanged, Step 2 emits nothing |
| Step-1 contract tests | green, unchanged |

No test was deleted, skipped or weakened.

### Baseline preservation

Byte-identical to the accepted Step-1 authority `7d8e73e`:

| Path | `git diff 7d8e73e` |
|---|---|
| `pccm/spec` | **EMPTY** — the whole directory |
| `pccm/src` | **EMPTY** |
| `pccm/bootstrap` | **EMPTY** |
| `pccm/evidence` | **EMPTY** |
| `pccm/docs/phase6_plan.md` | **EMPTY** |
| `pccm/docs/phase6_step0.md` | **EMPTY** |
| `pccm/docs/phase6_step1.md` | **EMPTY** |
| `pccm/builder/build_stage_a.py` | **EMPTY** |

Phase-5 fingerprint vectors, recomputed live at this commit:

```
fingerprint("PCCM-FP")     6551C6F365DA7F3F
fingerprint_probe(A|B)     42E49DC715F06970
fingerprint_probe(AB|)     7558FD9248656EAD
canonical_number(1/3)      3.3333333333333331E-01
```

`test_phase5_fingerprint.py` — **52 passed**.

---

## 11. Step-2 acceptance gate

| Requirement | Status |
|---|---|
| fixed seed → state exact | yes |
| AUTO nonce → seed exact | yes |
| all retained RNG vectors pass | yes |
| first 20 uniforms exact | yes — `float.hex()`, no tolerance |
| jump states exact | yes — 0, 1, 7, 399, 401 |
| stream 1 matches the published second-stream state | yes |
| a stream beyond 400 passes | yes — 401 |
| component assignments match retained vectors | yes |
| row reorder changes nothing | yes |
| ordinal order implemented explicitly | yes — accepted Phase-5 key, astral case tested |
| no sampler | yes — §9 |
| no Monte Carlo | yes — §9 |
| no VBA | yes — §9 |
| no spec change | yes — `pccm/spec` diff empty |
| no Windows/Excel runtime | yes — Linux only |
| full regression green | yes — §10 |

---

**STEP 2 — ACCEPTANCE REQUESTED**
