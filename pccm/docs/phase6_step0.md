# PCCM Phase 6 — Step 0 authority record

Step 0 closes the remaining Phase-6 semantic decisions and produces the retained
feasibility evidence the Class-1 decisions need. It writes no production code, no
contract and no VBA.

---

## 1. Accepted planning baseline

| | |
|---|---|
| **Accepted Phase-6 planning baseline** | `03aa5044cb535513976f0ec3840bc332747678c8` |
| **Accepted Phase-5 executable baseline** | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| **Authority for this round** | `pccm/docs/phase6_plan.md` at revision 6 |

One docs-only change was made to `phase6_plan.md` in this round: the §14 wording
correction authorised in §0 of the Step-0 authorisation. It replaces "for
identical statistical output" with a statement that both D6-18 options target the
same probability law while following different realised severity paths, and it
carries an inline note classifying itself as **non-semantic**. It changes no
decision, no number and no requirement. **No other planning change was made and
no new planning revision was created.**

---

## 2. Evidence package provenance

| | |
|---|---|
| **Location** | `pccm/evidence/phase6_step0/` |
| **Regenerate** | `cd pccm/evidence/phase6_step0 && python3 scripts/run_all.py` |
| **Interpreter** | CPython 3.11.15, `Linux-6.18.44-fc-v21-x86_64-with-glibc2.39` |
| **Third-party packages** | none |
| **Network** | none |
| **Manifest** | `pccm/evidence/phase6_step0/manifest.json` — SHA-256 and byte count of every script, input, raw result, vector file and summary |
| **Determinism check** | two consecutive runs produce a byte-identical `manifest.json`; because the manifest hashes every generated file, that single comparison covers all of them |

**The manifest cannot contain its own commit hash** — a commit cannot record the
hash it is about to be given. It records the accepted planning baseline, the
accepted Phase-5 baseline, and `git_head_at_generation`, which is the **parent**
of the Step-0 commit and equals the accepted planning baseline. The Step-0 commit
hash is reported in the delivery message; `git log -1` on the branch gives it
directly.

**External reference provenance.** Two published values are used as *checks*, not
as inputs: the RngStreams default initial state `[12345] × 6` and the first
uniform it produces. Everything else — including both jump matrices — is derived
inside the package from the recurrence. No remembered literature value is treated
as evidence anywhere in this package; see §5.3.

### Evidence code is not production code

`scripts/` is not imported by `pccm/builder`, `pccm/src` or `pccm/tests`, nothing
in the production trees references it, and no file in it is offered as a Phase-6
implementation. The dependency runs one way only: `scripts/digest.py` and
`scripts/stream_map.py` **import** the accepted Phase-5 encoder from
`pccm/builder/pccm_builder/calc_fingerprint.py`, read-only, so that D6-16 and
D6-17 demonstrably reuse accepted authority rather than reimplementing it.

`scripts/controls.py` holds negative controls for the **evidence**. They are not
Phase-6 tests, they are outside `pccm/tests`, and the project's pytest suite does
not run them.

---

## 3. Negative controls on the evidence itself

Each control plants a defect this package is supposed to detect and asserts the
detection fires. A control that cannot fail proves nothing.

| # | Control | Fired |
|---|---|---|
| 1 | wrong MRG modulus `m1` | yes |
| 2 | wrong jump-matrix element | yes |
| 3 | naive unsafe modular matrix-vector product in float | yes |
| 4 | BB/BC dispatch boundary moved | yes |
| 5 | occurrence and severity streams merged | yes |
| 6 | physical-row-order stream assignment | yes |
| 7 | D6-18 advancement rule swapped | yes |
| 8 | result-digest field/order mutation | yes |
| 9 | non-primitive auto-seed multiplier | yes |
| 10 | counter exhaustion wrap | yes |
| 11 | invalid seed produces the forbidden all-zero state | yes |
| 12 | result-stream framing tag mutation | yes |
| 13 | stream assignment under a different ordering rule | yes |
| 14 | Family B collision under the accepted unbounded ID pattern | yes |

**14 / 14 fired.** Recorded in `summaries/controls.json` and `summaries/controls.txt`.

---

## 4. Class-2 decisions — closed

### D6-03 — AUTO seed source, freshness and exhaustion — **CLOSED: option (b)**

A full-period multiplicative cycle over the seed domain. Evidence:
`raw/seed_map.json`.

| Element | Value |
|---|---|
| Modulus | `2147483647` (`2^31 − 1`, prime) |
| Multiplier | `48271` |
| Recurrence | `x_{k+1} = 48271 · x_k mod 2147483647` |
| Starting nonce | `0` |
| First effective seed | `1` (nonce 0), then `48271` (nonce 1), `182605794` (nonce 2), `1291394886` (nonce 3) |
| Persisted state | `auto_nonce`, the counter — **not** the seed |
| Collision-free range | nonces `0 … 2147483645`, all distinct |
| Period | `2147483646` |
| Exhaustion point | nonce `2147483646`, which returns to seed `1` |
| Refusal behaviour | **refuse to run**; never wrap silently |

**Full period is proved, not asserted.** `48271` is a primitive root mod
`2^31 − 1`: `2147483646 = 2 · 3 · 7 · 11 · 31 · 151 · 331` (complete
factorisation, computed in the package), and `48271^((p−1)/q) ≠ 1 mod p` for
every one of those seven primes. The order is therefore exactly `p − 1`. The
package additionally checks `48271^(p−1) ≡ 1` and that the first 200,000 nonces
produce no repeat.

**The period equals the seed domain exactly.** `2147483646` values, and the
accepted D6-20 FIXED domain is `1 … 2147483646` — also `2147483646` values. The
cycle visits every admissible seed exactly once before exhausting. That is a
property of the choice, and it is checked in the evidence.

**Scope of the freshness guarantee — stated plainly.** Uniqueness holds **only
within one workbook's persisted nonce history**. Two independent workbooks each
start at nonce `0` and will produce the *same* seed sequence. **This mechanism
provides no cross-workbook uniqueness and none is claimed.** A user who needs two
workbooks to differ must set FIXED seeds. No timestamp-derived uniqueness is used
anywhere.

### D6-05 — scalar seed → six-word MRG32k3a state — **CLOSED: option (a), repeated scalar**

`state = [s, s, s, s, s, s]`.

Validity over the whole accepted domain follows from one inequality chain,
checked in `raw/seed_map.json`:

```
seed_max  2147483646  <  m2  4294944443  <  m1  4294967087
```

Every admissible seed is therefore a valid residue in **both** components, and
since the minimum admissible seed is `1`, neither the `m1` triple nor the `m2`
triple can be all-zero. Endpoints `1` and `2147483646` are checked explicitly;
`0` is checked to be **invalid**, which is why the accepted domain excludes it.

**No mixer is introduced.** A mixer would be a new portability surface and a new
cross-language contract; nothing in the evidence demonstrates a requirement for
one. Seed `12345` is retained as a probe precisely because `[12345] × 6` is the
published RngStreams default state, so it compares directly against independent
MRG32k3a material (§5.1).

### D6-14 — Phase-5 analytical prerequisite — **CLOSED: option (a), require Phase-5 `CURRENT`**

Closed on authority. Option B (silently recalculate) would make a simulation run
mutate Phase-5 published outputs and the analytical fingerprint underneath the
user, which contradicts the accepted Phase-5 publication and attempt-state
semantics: a Phase-5 publication is a recorded event with its own stamp and
attempt fields, and Phase 6 is not authorised to manufacture one. Option C (one
combined transaction) would additionally have to define what a partial failure
leaves behind across two publication states, which is new semantics with no
demonstrated need.

Option A preserves both: Phase 6 reads a `CURRENT` Phase-5 result or refuses, and
a refusal leaves the prior Phase-5 state exactly as it was.

### D6-15 — `run_id` — **CLOSED: option (a), monotonic success counter**

| Element | Rule |
|---|---|
| Initial value | `0` before any successful run; the first successful commit allocates `1` |
| Allocation | on **successful commit only** |
| Persistence | stored in the workbook; survives save and reopen unchanged |
| Failure behaviour | a failed or refused run allocates nothing; the counter does not advance |
| Maximum | `2147483647` (VBA `Long` maximum) |
| Exhaustion | **refuse to run**; never wrap, never reuse |
| Relationship to `auto_nonce` | **none** — independent counters. A FIXED-seed run allocates a `run_id` and no nonce; an AUTO run allocates both |
| Relationship to `effective_seed` | **none** — `run_id` is audit identity with no computational effect; it enters no RNG input and no digest input |
| Save / reopen | the persisted value is authoritative; reopening never resets it |

No GUID, timestamp or COM identity is used. None is needed: the requirement is a
monotonic audit identity within one workbook, which a counter satisfies exactly.

### D6-16 — stream assignment — **CLOSED: Family A, on authority**

**The accepted Permanent-ID authority was inspected first.** The accepted
patterns are:

```
^CL-[0-9]{3,}$        ^R-[0-9]{3,}$
```

`{3,}` is **unbounded above**. The repository therefore does **not** guarantee a
bounded numeric Permanent-ID domain, and Family B requires one.

This is demonstrated, not argued. `raw/stream_architecture.json` contains
collision witnesses: for **any** proposed finite Risk-block offset `K`, the
Cost-Line ID `CL-` + `str(K).zfill(3)` is admitted by the accepted pattern and
maps to index `K` — the same index as `R-000`. Witnesses are recorded for
`K = 1000`, `K = 100000` and `K = 2147483647`, and control 14 asserts the
collision actually occurs.

**Family A is selected because Family B has no bounded numeric domain in any
accepted contract — not because revision 6 recommends it.** Family B remains the
better user experience and would become available if a future accepted contract
bounds the ID numeric part; that is a contract change, not a Step-0 choice.

Family A, demonstrated in `raw/stream_architecture.json`:

| Requirement | Evidence |
|---|---|
| Deterministic assignment | components ordered by `(ComponentKind, PermanentId, Role)`; stream index is the position in that order |
| Collision freedom | the assignment is a **bijection onto `0 … N−1`**, checked directly |
| Row-order invariance | the same map results from a shuffled input (seeded shuffle, recorded) |
| Locale independence | ordering uses the **accepted Phase-5 `utf16_sort_key`**, ordinal on UTF-16 code units — not Python string order, not locale collation, not case-insensitive comparison |
| Deterministic cross-language mapping | every admitted ID is ASCII (checked), so code-point and UTF-16 code-unit order coincide; no new collation authority is created |

Component architecture at the design target — Cost Line → **1** stream; Risk →
**2** streams (occurrence, severity):

```
200 cost components + 200 risk components (2 × 100) = 400 components
stream state memory                                 = 400 × 6 × 8 = 19,200 B
```

Streams `0 … 9` and `396 … 399` are recorded in
`vectors/stream_assignment_vectors.json`.

**Two consequences are preserved explicitly, not buried:**

1. **Adding or removing a driver may shift later drivers to different streams**,
   so an unrelated driver's samples change. Adding a driver changes the model, so
   the run is a new run — but a user may expect an unrelated driver's numbers to
   hold still, and they will not.
2. **Ordinal order is not numeric order.** Once IDs widen, `CL-1000` sorts
   *before* `CL-999`. This is deterministic and portable, and it is recorded in
   the evidence, but it is not the order a reader would guess.

**Stream assignment is kept separate from driver accumulation ordering.** Nothing
here changes accumulation order, which stays as the accepted plan defines it.

### D6-17 — `result_digest` canonicalisation — **CLOSED: option (a), the accepted encoder**

**No competing hash algorithm is created.** The base, both moduli, both initial
accumulators, the UTF-16 code-unit folding and the canonical Double text are all
**imported** from `pccm/builder/pccm_builder/calc_fingerprint.py`. D6-17 chooses
only the framing.

The locked stream:

```
stream ::= F_S("PCCM-RD") F_I(version) section("RESULT", record*)
record ::= F_I(field_count) F_I(iteration_index) F_N(nominal) F_N(pv)
section::= F_S("RESULT") F_I(record_count) record*
```

| Element | Value |
|---|---|
| Digest / version tag | `PCCM-RD`, then the algorithm version as a stream integer. Distinct from the input fingerprint's `PCCM-FP`, so an input stream and a result stream can never coincide |
| `n` | the section's **record count**, hashed by the accepted grammar |
| Iteration index | the record's **first field**, 1-based, hashed |
| Nominal value | `F_N`, canonical Double |
| PV value | `F_N`, canonical Double |
| Field ordering | index, nominal, PV — fixed, never sorted |
| Canonical Double encoding | the accepted `canonical_number`: 17 significant digits, decimal point always present, uppercase `E`, exponent sign always present, at least two exponent digits, negative zero normalised to positive zero |
| Type tags and separators | the accepted `F_S` / `F_I` / `F_N` tagging with UTF-16 lengths — inherited, not redefined |

Proved in `raw/digest.json`, with all seven digests distinct:

| Property | How |
|---|---|
| Iteration order matters | reversing the iteration order changes the digest |
| Truncation cannot masquerade as a shorter valid run | dropping one iteration changes the digest — the record count is hashed, so a truncated run is a *different* stream, not a valid shorter one |
| Nominal/PV swap changes the digest | swapping the two arrays changes the digest |
| Same-runtime replay is exact | the digest is a pure function of the totals; identical inputs give an identical 16-hex-character digest |
| One-ULP sensitivity | perturbing one total by one ULP (`math.nextafter`) changes the digest |
| Version participates | version `2` differs from version `1` on identical data |

### D6-19 — Random Seed admissible-domain ownership — **CLOSED: option (a), input-contract ownership**

**Decided from authority.** `spec/input_contract.yaml` already owns
`inpRandomSeed` and its note reads: *"The admissible domain is fixed when the RNG
is implemented."* That is an authority **deferral in place** — the input contract
is holding the rule open and waiting for this decision — not a delegation to
another file. Resolving it anywhere else would leave the note undischarged and
put one semantic rule in two files.

Therefore: Step 1 updates `input_contract.yaml` with the resolved admissibility
(whole integer, `1 … 2147483646`, blank = AUTO) and replaces `validation: null`
with the real rule. `sim_contract.yaml` **references** it and owns RNG and
seeding *execution* semantics only. **No duplicate range may exist in two
contracts.**

The same discipline applies to `monte_carlo_iterations`: its `≥ 1000` business
minimum stays in `input_contract.yaml`, and Phase 6 adds only the technical
storage ceiling (D6-08), which is a different category.

### D6-20 — not reopened

Nothing in the Step-0 evidence contradicts the accepted domain. D6-05's proof
*depends* on the upper bound `2147483646` being below `m2`, so the evidence
supports D6-20 rather than straining it. No authority conflict was found and the
domain is not widened.

---

## 5. D6-04 — Beta sampler, component streams and jump arithmetic

### 5.1 The RNG backbone

`raw/rng_backbone.json`, `vectors/rng_vectors.json`.

Two independent arithmetic paths over the same recurrence:

- **Path 1** — Python unbounded integers. The mathematical definition.
- **Path 2** — the arithmetic a VBA implementation would perform: every value
  held as a `Double` that happens to be an integer, reduced by `Fix(p/m)`.

**The two paths agree exactly** over the recorded draws. That is what
substantiates the plan's §5.2.1 claim that `Double` arithmetic is exact for this
recurrence — demonstrated, not asserted.

Checks recorded: every uniform strictly inside `(0,1)`; the first uniform from
the all-`12345` state is `0.12701112204657714`, the published RngStreams value;
first five uniforms recorded for seeds `1`, `2`, `12345`, `2147483646`; state
after 20 draws recorded.

### 5.2 Cheng BB/BC — candidate B, measured

`raw/cheng.json`. **17 shapes spanning the whole one-dimensional PERT family**
(`α = 1 + 4r`, `β = 5 − 4r`, `α + β = 6`): `r = 0` (`1/5`, BC), `0.01`, `0.05`,
`0.1`, `0.2`, `0.25`, `0.3`, `0.4`, `0.5` (`3/3`), `0.6`, `0.7`, `0.75`, `0.8`,
`0.9`, `0.95`, `0.99`, `r = 1` (`5/1`, BC). Nothing outside that line is
measured, and nothing outside it can arise from a PCCM driver.

20,000 samples per shape, fixed retained seed `987654321` (all six state words),
recorded in `inputs/seeds.json`.

**Counting definitions, stated exactly:**

- **one proposal attempt** — one entry into the rejection loop body, counted at
  the top before any uniform is drawn;
- **one uniform consumed** — one call to the underlying MRG32k3a `next_u`. Cheng
  draws exactly two per attempt, so consumption is `2 × attempts`;
- **one accepted sample** — one return from the sampler with a variate in `(0,1)`.

Acceptance is measured from the loop, **never inferred from the target density**.

| `r` | `α` | `β` | dispatch | acceptance | uniforms / sample | p50 | p90 | p95 | p99 | max | transcendentals / sample |
|---|---|---|---|---|---|---|---|---|---|---|---|
| 0.00 | 1.00 | 5.00 | **BC** | 0.7464 | 2.6795 | 2 | 4 | 6 | 8 | 14 | 4.007 |
| 0.01 | 1.04 | 4.96 | BB | 0.7630 | 2.6212 | 2 | 4 | 6 | 8 | 16 | 4.825 |
| 0.05 | 1.20 | 4.80 | BB | 0.8076 | 2.4766 | 2 | 4 | 4 | 6 | 14 | 4.559 |
| 0.10 | 1.40 | 4.60 | BB | 0.8455 | 2.3654 | 2 | 4 | 4 | 6 | 12 | 4.364 |
| 0.20 | 1.80 | 4.20 | BB | 0.8878 | 2.2527 | 2 | 4 | 4 | 6 | 12 | 4.191 |
| 0.25 | 2.00 | 4.00 | BB | 0.9000 | 2.2222 | 2 | 2 | 4 | 6 | 12 | 4.158 |
| 0.30 | 2.20 | 3.80 | BB | 0.9097 | 2.1985 | 2 | 2 | 4 | 4 | 10 | 4.142 |
| 0.40 | 2.60 | 3.40 | BB | 0.9198 | 2.1743 | 2 | 2 | 4 | 4 | 8 | 4.152 |
| 0.50 | 3.00 | 3.00 | BB | 0.9223 | 2.1684 | 2 | 2 | 4 | 4 | 8 | 4.202 |
| 0.60 | 3.40 | 2.60 | BB | 0.9198 | 2.1743 | 2 | 2 | 4 | 4 | 8 | 4.152 |
| 0.70 | 3.80 | 2.20 | BB | 0.9097 | 2.1985 | 2 | 2 | 4 | 4 | 10 | 4.142 |
| 0.75 | 4.00 | 2.00 | BB | 0.9000 | 2.2222 | 2 | 2 | 4 | 6 | 12 | 4.158 |
| 0.80 | 4.20 | 1.80 | BB | 0.8878 | 2.2527 | 2 | 4 | 4 | 6 | 12 | 4.191 |
| 0.90 | 4.60 | 1.40 | BB | 0.8455 | 2.3654 | 2 | 4 | 4 | 6 | 12 | 4.364 |
| 0.95 | 4.80 | 1.20 | BB | 0.8076 | 2.4766 | 2 | 4 | 4 | 6 | 14 | 4.559 |
| 0.99 | 4.96 | 1.04 | BB | 0.7630 | 2.6212 | 2 | 4 | 6 | 8 | 16 | 4.825 |
| 1.00 | 5.00 | 1.00 | **BC** | 0.7464 | 2.6795 | 2 | 4 | 6 | 8 | 14 | 4.007 |

**Family bounds:** acceptance `0.7464 … 0.9223` · uniforms per sample
`2.1684 … 2.6795` · p99 ≤ `8` · maximum observed `16` · transcendental calls per
sample `4.007 … 4.825`.

The table is symmetric about `r = 0.5`, as the shape family requires. That
symmetry was not enforced anywhere in the code; it emerges from the measurement
and is a useful independent sanity signal.

**Independent distribution checks**, against theory computed separately
(`mean = α/(α+β)`, `var = αβ/((α+β)²(α+β+1))`):

| | worst over the family |
|---|---|
| mean absolute error | `2.433e-03` |
| variance absolute error | `6.033e-04` |

At 20,000 samples the standard error of the mean is ≈ `1.3e-03`, so the observed
errors are Monte Carlo noise, not bias. **This check earned its place**: it
caught a real defect in the BC branch of the reference implementation, where the
parameter swap was inverted and the sampler returned a valid Beta variate of the
*mirrored* distribution — mean error `6.65e-01`. A sampler that returns
plausible-looking variates from the wrong distribution is exactly the failure a
theoretical-moment check exists to catch.

**One measurement note, stated rather than silently optimised away.** The
recorded per-sample transcendental count includes per-call setup (`sqrt` in BB;
`k1`, `k2` in BC) that a production implementation hoists to per-driver setup,
because `α` and `β` are per-driver constants. The `sqrt` breakdown shows exactly
one per sample for BB shapes, so hoisting would remove exactly `1.0` from those
rows. The figures below are the **un-hoisted**, i.e. pessimistic, ones.

### 5.3 Floating acceptance-path evidence

`raw/cheng.json` (per-shape margins) and `raw/acceptance_margin_model.json`.

Every accept/reject comparison in BB and BC is routed through one instrumented
`>=` that records how close the two sides came. **805,837 predicate evaluations**
were recorded across the family.

| | |
|---|---|
| Closest observed **relative** margin | `7.605e-07` |
| Closest observed **absolute** margin | `2.228e-06` |
| Closest margin ÷ one ULP (`2^-52`) | `3.4e+09` |
| Evaluations within `1e-9` relative of a boundary | **0** |
| Evaluations within `1e-12` relative of a boundary | **0** |
| Evaluations within `1e-15` relative of a boundary | **0** |

Extrapolating from the measured near-zero density of relative margins
(`3.57 … 5.13` per unit relative margin, from the counts below `1e-3`), the
expected number of one-ULP branch flips in a full design-target run is
**`≤ 9.4e-08`** — roughly one run in eleven million.

**The assumptions of that extrapolation are recorded with it** and are not
proved: that the margin density is roughly constant near zero, that it transfers
to the design-target workload, that one ULP is the relevant perturbation, and
that this concerns **Python** `Double` arithmetic.

**This is diagnostic evidence about floating-path fragility and nothing more.**
Step 0 has no VBA runtime and no Phase-6 VBA implementation. It therefore does
**not** conclude that Python and VBA Cheng paths are cross-language exact, and
**revision 6's Layer-G rule stands unchanged**: a full seeded Beta simulation is
not required to match sample-for-sample across Python and VBA.

### 5.4 Candidate A — inverse CDF

**The prior "~91 continued-fraction iterations" claim did not reproduce.** It is
withdrawn and replaced by measurement.

| Element | A1 (naive) | A2 (practical) |
|---|---|---|
| Incomplete-beta method | Lentz-style continued fraction with the standard `x < (α+1)/(α+β+2)` reflection | same |
| Convergence criterion (CF) | `|δ − 1| < 3e-16`, cap 400 | same |
| Inverse / root method | 60 unconditional bisection halvings | **safeguarded Newton**: analytic derivative (the Beta density), bracket maintained, a step leaving the bracket replaced by a bisection step |
| `log B(α,β)` | recomputed inside every evaluation | **hoisted to per-driver setup** — `α`, `β` are per-driver constants |
| Worst CF iterations | **14** | **14** |
| Worst root iterations | 60 (fixed) | **11** |
| Transcendental calls / sample | **360** | **66** |
| Per-driver setup | none | 3 `lgamma`, once |
| Uniforms / sample | **1** | **1** |
| Accuracy | reference | **`4.108e-15`** max absolute difference from A1 |

Worst shapes for the continued fraction: `r ∈ {0.01, 0.05, 0.95, 0.99}` under A1
and `r ∈ {0.05, 0.95}` under A2, 14 iterations in both cases. Worst root
iterations (11) occur at the near-boundary shapes `r ≤ 0.25` and `r ≥ 0.75`; the
interior shapes `0.3 … 0.7` need 10.

**Operation-counting method:** transcendental calls (`log`, `log1p`, `exp`,
`sqrt`, `lgamma`) counted by instrumented wrappers, the **same** wrappers in
every candidate. **No flops-per-iteration multiplier is applied anywhere.** The
earlier `×20` figure is withdrawn along with the iteration claim.

**The inverse-CDF family is not rejected on A1's implementation.** A1 is retained
because it is the implementation the earlier estimate described and because it is
an independent accuracy reference; A2 is the same family implemented competently
and is `5.5×` cheaper. A **practical VBA-suitable implementation does exist** —
A2 uses only `log`, `exp`, `lgamma` and arithmetic, all available, with no
recursion and no dynamic allocation.

A2 also surfaced a genuine defect worth recording, because it is the kind of bug
that yields plausible wrong answers: at convergence the just-updated bracket end
equals `x`, the strict test `lo < nx < hi` is false, and an unguarded
implementation throws the converged value away to bisect a bracket whose far end
was never tightened. That produced errors of `1.5e-05 … 3.9e-05` — small enough
to look like tolerance, large enough to be wrong. It was caught by comparing
against A1, and the `nx == x` branch that fixes it carries the explanation
inline.

### 5.5 Candidate C — precomputed table

**The prior "~5.9e-3 normalised error at 4096 nodes" claim did not reproduce.**
It is withdrawn and replaced by the measured error/memory curve.

- **Construction:** per driver, a table of `nodes + 1` inverse-CDF values on a
  uniform grid of `u ∈ [0,1]`, built with the A2 inverse CDF.
- **Interpolation:** linear between adjacent nodes.
- **Parameter grid:** `r ∈ {0, 0.05, 0.5, 0.95, 1}` — a **stated subset** of the
  17-shape grid, chosen because the tails are where a uniform node grid fails, so
  `r = 0` and `r = 1` must be present, with interior shapes to show the error is
  not an artefact of the extremes alone. The full 17-shape sweep at 65,536 nodes
  did not complete in reasonable time and the subset is declared rather than
  quietly substituted.
- **`u` grid:** the 11 probes `0.001 … 0.999`, including both tails.

| nodes | max abs normalised error | worst at | bytes / driver | bytes, 300 drivers | build evals / driver |
|---|---|---|---|---|---|
| 256 | `1.667e-01` | `r=0.0, u=0.999` | 2,056 | 616,800 | 257 |
| 1,024 | `2.965e-04` | `r=0.0, u=0.999` | 8,200 | 2,460,000 | 1,025 |
| 4,096 | `9.329e-05` | `r=0.0, u=0.999` | 32,776 | 9,832,800 | 4,097 |
| 16,384 | `1.757e-05` | `r=1.0, u=0.001` | 131,080 | 39,324,000 | 16,385 |

**Larger tables do materially improve the result**, so the curve is reported
rather than one node count being judged alone. The worst location is the tail in
every case, as expected: the quantile function is steepest there and a uniform
node grid resolves it worst.

### 5.6 Jump-ahead arithmetic

`raw/jump.json`, `vectors/jump_vectors.json`.

**Provenance.** `A1^(2^127)` and `A2^(2^127)` are **derived** in the package by
repeated squaring of the transition matrices `A1` and `A2`, which are written
directly from the MRG32k3a recurrence. They are **not** copied from any published
table, and no published table is treated as evidence. This matters: recalled
literature values for these matrices did **not** match the derived ones — they
were reversed in both indices — and the derived matrices are what the checks
below validate. Recollection is not evidence and is not used as any.

```
A1^(2^127) = 1230515664, 986791581,1988835001
             3580155704,1230515664, 226153695
              949770784,3580155704,2427906178
  sha256 e31a727398a2d4461cf708f77034b9bc5e60f88c54556c56c3f4b015a813b66a

A2^(2^127) = 2093834863,  32183930,2824425944
             1022607788,1464411153,  32183930
             1610723613, 277697599,1464411153
  sha256 0d20b47aa206b1231c22e20afaa84e71b81cff52a2395be9aeb0bbb97b1e8208
```

**State-vector convention, stated because getting it wrong is silent.** `A1` and
`A2` operate on **newest-first** vectors `[x_{n−1}, x_{n−2}, x_{n−3}]`; PCCM
stores state **oldest-first** as `[s10, s11, s12]`. Both directions are reversed
at the boundary. This was determined empirically by the one-step check below, not
assumed — the wrong convention produces a plausible but wrong stream.

**Why the naive form is wrong**, with the actual derived matrix rather than a
bound:

```
worst single naive product   15,376,650,911,435,158,544  = 1,707 × 2^53
worst naive three-term sum   29,883,662,290,365,631,276  = 3,318 × 2^53
```

(The plan's `2048×` is the bound for the largest *representable* entry; these are
the actual worst values for the actual derived `J1` against a maximal state
vector. The conclusion is unchanged and now rests on measured numbers.)

**Two independent arithmetic paths, compared exactly:**

1. Python arbitrary-precision integers;
2. the proposed VBA-safe decomposition — L'Ecuyer `MultModM` with `H = 2^17`,
   written in float arithmetic to mirror what VBA would do.

| Check | Result |
|---|---|
| One step by matrix == one step by recurrence | **pass** |
| An 8-step jump == 8 sequential steps | **pass** |
| Safe `MultModM` == exact-integer reference at the `2^127` jump | **pass** |

**Required jump vectors, all retained** in `vectors/jump_vectors.json` with the
first five uniforms and the state after five draws for each:

| stream | initial state |
|---|---|
| base | `12345, 12345, 12345, 12345, 12345, 12345` |
| 0 | `12345, 12345, 12345, 12345, 12345, 12345` |
| 1 | `3692455944, 1366884236, 2968912127, 335948734, 4161675175, 475798818` |
| 7 | `3281794178, 2616230133, 1457051261, 2762791137, 2480527362, 2282316169` |
| 399 | `2260181002, 1948664812, 612976419, 1919355493, 2890171896, 2701138777` |
| 401 | `215541976, 1807926449, 2979430890, 2228004365, 3803991720, 370726289` |

Stream 1 reproduces the published RngStreams second-stream state for the `12345`
default — an independent confirmation that the derived matrices and the vector
convention are both right.

**No substream matrices are computed.** None is needed: the design advances each
component stream sequentially and replays by resetting to the initial state
(plan §5.8). Computing `2^76` matrices for a facility the design does not use
would be unearned work and a second contract to maintain.

**This proves arithmetic feasibility. It does not claim a VBA runtime passed.**

### 5.7 The D6-04 comparison

`raw/d6_04_comparison.json`. One counting method, one workload, four
implementations. Workload = the accepted design target's worst case: 200 Cost
Lines + 100 Risks, all Beta-PERT, every risk at `Probability = 1` →
`3.0 × 10^7` Beta samples, `1.0 × 10^7` occurrence uniforms, 300 drivers.

| candidate | transcendentals / sample | transcendentals / run | uniforms / run | extra resident bytes | worst sampling error |
|---|---|---|---|---|---|
| **A1** naive inverse CDF | 360 | `1.08e+10` | `4.00e+07` | 0 | reference |
| **A2** Newton inverse CDF | 66 | `1.98e+09` | `4.00e+07` | 0 | `4.11e-15` vs A1 |
| **B** Cheng BB/BC | 4.825 | `1.45e+08` | `9.04e+07` | 0 | none — exact in distribution |
| **C** table, 4096 nodes | 0 | `8.11e+07` (build only) | `4.00e+07` | **9,832,800** | `9.33e-05` |

**Candidate C — REJECTED**, on measured grounds, none of which is the withdrawn
`5.9e-3` figure:

1. **The error is systematic, not random.** Interpolation bias is the same on
   every iteration, so it does not average away over 100,000 iterations — it
   biases every reported quantile, which is the deliverable. `9.33e-05`
   normalised is ~9,300 currency units on a driver with a 100M range, applied
   identically to every iteration.
2. **Memory.** 9.83 MB at 4096 nodes against a plan §14 peak-resident budget of
   ~8.2 MB — it more than doubles peak residency. 16,384 nodes costs 39.3 MB.
3. **It is strictly additive complexity.** The table must be *built* with an
   inverse CDF, so choosing C means implementing A2 in VBA **anyway**, plus the
   table and interpolation machinery on top.
4. **It adds a second replay contract.** The table becomes part of what G2/G3
   replay must reproduce bit-exactly across languages.

Its measured compute advantage over B (`8.11e+07` vs `1.45e+08` transcendental
calls) is real and is recorded, but it is not worth a systematic bias in the
output distribution plus a second numerical contract.

**Candidate A1 — REJECTED as an implementation**, `5.5×` more expensive than A2
for identical output. This rejects the implementation, **not** the family.

**A2 versus B is the live decision, and it is a genuine trade:**

- **B is `13.7×` cheaper** in transcendental calls (`1.45e+08` vs `1.98e+09`),
  and costs `2.3×` the uniforms (`9.04e+07` vs `4.00e+07`).
- **A2 consumes exactly one uniform per sample.** A one-ULP difference between
  two implementations perturbs one sample and **cannot cascade**, because the
  stream position after the sample is the same either way. Under B, a one-ULP
  difference in an acceptance test changes how many uniforms that stream
  consumed and **desynchronises every subsequent draw on it**.

**Selected: B — Cheng BB/BC with component streams and `2^127` jumps.**

Grounds, stated so the reviewer can weigh them:

1. The `13.7×` measured compute advantage, on one counting method applied to all
   four candidates.
2. A2's non-cascading property buys less than it appears to. Realising it as
   *exactness* would require the VBA continued fraction to match Python bit for
   bit, including a data-dependent termination test — a strong requirement Step 0
   cannot check. Without that, A2 also yields tolerance-bounded comparison, and
   its advantage reduces to "divergence stays local", which is real but does not
   deliver cross-language exactness either.
3. The measured acceptance-path margins (§5.3) show that under Python `Double`
   arithmetic the branch decisions are nowhere near their boundaries, which is
   consistent with B being stable in practice — while explicitly not proving
   anything about VBA.
4. The plan's Layer-G honesty already accounts for B's desynchronisation risk and
   is unchanged by this evidence.

**What would change this decision, stated plainly:** a Windows measurement
showing B's compute advantage is unnecessary at the design target would make A2
architecturally preferable, because non-cascading divergence is worth paying for
when the compute is affordable. **Step 0 cannot produce that measurement** — it
is Gate-B evidence and Step 0 is explicitly not a Windows/VBA performance
benchmark. The recommendation therefore rests on the measured ratio and the
architectural argument above, **not** on the withdrawn revision-6 figures.

---

## 6. D6-18 — when the Risk severity stream advances

`raw/d6_18_operation_model.json`. Shared in both options: the occurrence stream
advances **once per Risk per iteration**, `1.0 × 10^7` uniforms at the design
target. Consumption estimates use the **measured** Cheng behaviour — the worst
shape (`2.6795` uniforms, `4.825` transcendentals per sample) is quoted below,
and the mid shape is also retained in the raw file so the figures are not tied to
one arbitrary shape.

| `p` | A invocations | B invocations | A severity uniforms | B severity uniforms | A transcendentals | B transcendentals |
|---|---|---|---|---|---|---|
| 1.00 | 10,000,000 | 10,000,000 | 26,795,000 | 26,795,000 | 48,253,000 | 48,253,000 |
| 0.50 | 5,000,000 | 10,000,000 | 13,397,500 | 26,795,000 | 24,126,500 | 48,253,000 |
| 0.10 | 1,000,000 | 10,000,000 | 2,679,500 | 26,795,000 | 4,825,300 | 48,253,000 |
| 0.01 | 100,000 | 10,000,000 | 267,950 | 26,795,000 | 482,530 | 48,253,000 |

**Memory impact: none.** Both options need the same 400 component stream states,
19,200 bytes. The difference is work, not storage.

**Effect on replay.** Both are deterministic and both replay by resetting streams
to their initial states and re-running. Neither supports arithmetic seek to
iteration `i`, for the reason the plan gives in §5.8: under a variable-draw
sampler the stream position at iteration `i` is not an arithmetic function of `i`.

**Effect on probability-only scenario comparability — the decisive difference.**
Under **B**, the number of severity sampler *invocations* before iteration `i` is
exactly `i − 1`, independent of the occurrence path. Change only a risk's
`Probability` and the severity sequence is **unchanged**; what changes is which
iterations use it. Under **A**, changing `Probability` changes which iterations
invoke the sampler, so the entire severity sequence shifts and two runs differing
only in a probability are no longer comparable draw for draw. Attribution — "this
risk contributed more because it occurred more often", not "because it also drew
different severities" — is only available under B.

**The cost objection, examined as the authorisation requires.** Option B's
severity work at **any** probability is exactly its work at `p = 1`, and at
`p = 1` the two options necessarily coincide. So B's cost is bounded above by a
workload the design target must already support: `1.0 × 10^7` severity samples,
`2.68 × 10^7` uniforms, `4.83 × 10^7` transcendental calls. **"B may be too
expensive" is therefore not a principled reason to select A.** If the `p = 1`
workload were infeasible, that would be evidence about D6-04, the sampler, or the
implementation strategy — it would not be evidence for changing what a
probability-only comparison means.

**Selected: B — unconditional advancement, value used only on occurrence.**

**Wording correction recorded.** Revision 6 §14's "identical statistical output"
is corrected to "the same target probability law / intended distribution". A
stronger equality is **not** proved and is not claimed: the two options do not
produce the same realised sample path, the same iteration-by-iteration severity
pairing, or the same `result_digest` under the same seed. The correction is
non-semantic and is applied inline in `phase6_plan.md` §14 with that
classification stated.

---

## 7. Alternatives rejected, and why

| Rejected | Why |
|---|---|
| D6-03 timestamp-derived seed | gives no guarantee, only a likelihood, and imports clock behaviour into a reproducibility contract |
| D6-03 refuse blank | removes a documented, accepted user affordance to solve a problem the counter already solves |
| D6-05 modular expansion / any mixer | a new portability surface and a new cross-language contract, with no demonstrated requirement; the one-inequality proof shows the repeated scalar is already valid over the whole accepted domain |
| D6-14 (b) silent recalculation | mutates Phase-5 published outputs and the analytical fingerprint under the user; contradicts accepted publication and attempt-state semantics |
| D6-14 (c) combined transaction | needs new partial-failure semantics across two publication states, with no demonstrated need |
| D6-15 GUID / timestamp / COM identity | no new authority requires them; a counter satisfies the actual requirement (monotonic audit identity within one workbook) exactly |
| D6-16 Family B | **no bounded numeric Permanent-ID domain exists in any accepted contract**; collision witnesses recorded for every proposed offset |
| D6-17 a separate hashing scheme | would create a second hash authority; the accepted Phase-5 encoder is imported instead |
| D6-19 split ownership | the input contract's own note defers the domain to the RNG decision; resolving it elsewhere leaves that note undischarged and puts one rule in two files |
| D6-04 candidate A1 | 360 transcendental calls per sample for output identical to A2's 66 — an implementation rejection, not a family rejection |
| D6-04 candidate C | systematic (non-averaging) quantile bias, 9.83 MB, still requires the inverse CDF to build the table, and adds a second replay contract |
| D6-04 candidate A2 | 13.7× the transcendental work of B, for a non-cascading property that does not deliver cross-language exactness without a bit-exact VBA continued fraction |
| D6-18 option A | destroys probability-only scenario comparability, and its cost saving is against a `p = 1` workload the design target must support anyway |
| `2^76` substream matrices | the design advances sequentially and replays by reset; unused facility, second contract |

---

## 8. Limitations of the evidence

1. **Everything here is Python on Linux.** No VBA was executed. Path 2 of the RNG
   and the `MultModM` jump arithmetic *model* what VBA `Double` arithmetic would
   do; they are mirrors, not executions.
2. **Cheng measurements use 20,000 samples per shape.** Moment errors are at the
   Monte Carlo noise floor (SE ≈ `1.3e-03`), which is adequate to detect a wrong
   distribution — it caught one — but is not a precision distributional test.
3. **The acceptance-margin extrapolation is an extrapolation**, resting on four
   stated, unproved assumptions, over Python `Double` arithmetic only.
4. **Candidate C was measured on 5 of the 17 shapes**, at 4 node counts. The
   subset is declared and justified; the full sweep at 65,536 nodes did not
   complete in reasonable time.
5. **The transcendental-call counter is a proxy for cost**, not a timing. It
   ignores arithmetic, branch and memory costs, which differ between an
   inverse-CDF loop and a rejection loop.
6. **Candidate A2's per-sample count excludes hoistable per-driver setup**, and
   candidate B's *includes* it — so the A2:B ratio is, if anything, generous to
   A2.
7. **Cheng's per-sample transcendental figures are un-hoisted** and therefore
   pessimistic by about `1.0` per sample for BB shapes.
8. **`u` probes are 11 fixed points**, not a dense sweep. They include both tails,
   which is where table interpolation and CF convergence are worst.
9. **The D6-16 demonstration uses the design-target ID set** (`CL-001 … CL-200`,
   `R-001 … R-100`) plus a widened-ID probe. It does not enumerate the unbounded
   ID domain — nothing could — which is precisely why Family B is rejected by
   witness rather than by search.

---

## 9. What Step 0 does NOT prove

- **It does not prove VBA or Excel performance.** No runtime measurement exists.
  Whether 100,000 iterations complete within any threshold is Gate-B evidence.
- **It does not prove cross-language exactness of the Cheng path.** Revision 6's
  Layer-G rule stands: a full seeded Beta simulation is not required to match
  sample-for-sample across Python and VBA.
- **It does not prove VBA `Double` arithmetic is exact for the recurrence.** It
  demonstrates that the *specified* arithmetic — `Fix`-based reduction, the
  `MultModM` decomposition — is exact when performed on IEEE-754 doubles. That
  VBA performs it as specified is Gate-B evidence.
- **It does not prove the RNG is statistically sound.** It reproduces published
  MRG32k3a vectors; it runs no test battery, and none is claimed.
- **It does not constitute Phase-6 Gate A**, and it accepts nothing. Independent
  review decides.
- **It proves nothing about memory or timing in Excel.** The memory figures are
  arithmetic on data-structure sizes, not measurements of a running workbook.
- **It does not establish that any implementation exists.** No Phase-6 production
  code was written.

---

## 10. Exact semantics Step 1 must encode

`sim_contract.yaml` does not exist and was not created. When Step 1 writes it, it
must encode exactly this and nothing not decided here:

**RNG and seeding**

1. MRG32k3a with `m1 = 4294967087`, `m2 = 4294944443`, `a12 = 1403580`,
   `a13n = 810728`, `a21 = 527612`, `a23n = 1370589`,
   `norm = 2.328306549295727688e-10`.
2. Uniform combination: `if p1 <= p2 then u = (p1 − p2 + m1) · norm else
   u = (p1 − p2) · norm`; output strictly inside `(0,1)`.
3. State stored oldest-first `[s10, s11, s12, s20, s21, s22]`; matrices operate
   newest-first; the reversal at the boundary is part of the contract.
4. Scalar seed → state: `[seed] × 6` (D6-05a).
5. AUTO seed: `effective_seed = 48271^auto_nonce mod 2147483647`, `auto_nonce`
   persisted, starting at `0`, refusing at `2147483646` (D6-03b). Freshness scope
   is **one workbook**.
6. The admissible seed domain lives in `input_contract.yaml`; `sim_contract.yaml`
   **references** it and stores no copy (D6-19a).

**Streams**

7. Components: Cost Line → 1; Risk → occurrence + severity. `C + 2R`; 400 at the
   design target.
8. Assignment: sort by `(ComponentKind, PermanentId, Role)` with Permanent ID
   compared ordinally on UTF-16 code units, using the accepted Phase-5 sort key;
   stream `k` is the base state advanced by `k` applications of the `2^127` jump
   (D6-16a).
9. The two jump matrices, exactly as in §5.6, with the `MultModM` `H = 2^17`
   decomposition mandatory — a naive `Mod` is forbidden and is a silent-error
   path.
10. Advancement is sequential per stream across iterations. **No arithmetic seek
    to iteration `i`.** Replay = reset to initial state and re-run.

**Sampling**

11. Beta-PERT via Cheng, dispatched **BB when `min(α,β) > 1`, BC otherwise** —
    equality belongs to BC (D6-04b).
12. Risk severity advances **unconditionally**, once per Risk per iteration; the
    value is used only on occurrence (D6-18b).
13. Per attempt: exactly two uniforms. The definitions of attempt, uniform
    consumed and accepted sample in §5.2 are the contract's definitions.

**Identities**

14. `result_digest` exactly as §4/D6-17, over iteration-ordered totals, using the
    accepted Phase-5 encoder, with the `PCCM-RD` tag, the version, the record
    count and the iteration index all inside the hashed stream.
15. `run_id`: monotonic success counter, first successful commit allocates `1`,
    maximum `2147483647`, refusal at exhaustion, no relationship to `auto_nonce`
    or `effective_seed` (D6-15a).

**Prerequisite**

16. Phase-5 analytical results must be `CURRENT`; Phase 6 never recalculates them
    (D6-14a).

---

## 11. D6-08 remains Step-1 layout-derived

**D6-08 — the technical storage ceiling constant (`1048576 − H`) — is Class 3 and
is not a Step-0 blocker.** It is not a semantic choice; it is a constant that the
`_SimData` layout determines. It closes in Step 1, from the accepted layout,
before the contract and its validator are considered complete. Step 0 neither
decides it nor guesses `H`.

---

## 12. Decision status after Step 0

| # | Class | Decision | Status |
|---|---|---|---|
| **D6-03** | 2 | AUTO seed source / freshness / exhaustion | **CLOSED — (b)** full-period cycle, `48271` mod `2^31−1` |
| **D6-04** | 1 | Beta sampler + streams + jump arithmetic | **CLOSED — (b)** Cheng BB/BC + component streams + `2^127` jumps |
| **D6-05** | 2 | scalar seed → six-word state | **CLOSED — (a)** repeated scalar |
| **D6-14** | 2 | analytical prerequisite | **CLOSED — (a)** require Phase-5 `CURRENT` |
| **D6-15** | 2 | `run_id` semantics and exhaustion | **CLOSED — (a)** monotonic success counter |
| **D6-16** | 2 | stream assignment rule | **CLOSED — (a)** canonical sorted order, on authority |
| **D6-17** | 2 | `result_digest` canonicalisation | **CLOSED — (a)** accepted encoder, `PCCM-RD` framing |
| **D6-18** | 1 | when the severity stream advances | **CLOSED — (b)** unconditional |
| **D6-19** | 2 | Random Seed domain ownership | **CLOSED — (a)** input-contract ownership |
| **D6-20** | — | Random Seed admissible domain | accepted at plan level; **not reopened**; evidence supports it |
| **D6-08** | 3 | technical storage ceiling | **OPEN by design** — Step-1, layout-derived |

**No Class-1 or Class-2 decision remains open. D6-08 alone remains for Step 1.**

---

## 13. Static baseline protection

Production trees against the accepted planning baseline `03aa5044`:

| Path | `git diff 03aa5044 -- <path>` |
|---|---|
| `pccm/src` | **EMPTY** |
| `pccm/bootstrap` | **EMPTY** |
| `pccm/spec` | **EMPTY** |
| `pccm/builder` | **EMPTY** |
| `pccm/tests` | **EMPTY** |

No spec change was made; none is authorised in Step 0.

Accepted static baseline, re-run at the Step-0 tree:

| Check | Result |
|---|---|
| Full Python suite (`python3 -m pytest pccm/tests -q`) | **1752 passed, 0 failed** (2 warnings, 127.62 s) |
| Stage-A verifier / build (`python3 pccm/builder/build_stage_a.py`) | **351 passed, 0 failed**; "Stage A build complete." |

**No test was altered.** The suites are byte-identical to the accepted baseline,
which is what the empty `pccm/tests` diff above establishes.

---

## 14. Step-0 completion gate

| Requirement | Status |
|---|---|
| Retained evidence exists and is reproducible | yes — one command, byte-identical manifest across runs |
| D6-04 explicitly resolved | yes — §5.7 |
| D6-18 explicitly resolved | yes — §6 |
| D6-03, D6-05, D6-14, D6-15, D6-16, D6-17, D6-19 resolved | yes — §4 |
| No unresolved Class-1 or Class-2 decision | yes — §12 |
| D6-08 alone remains for Step 1 | yes — §11 |
| No Phase-6 production implementation exists | yes — `pccm/src` diff empty |
| No `sim_contract.yaml` exists | yes — `spec/` diff empty; the file was not created |
| No Windows/Excel runtime was executed | yes — Linux only; recorded in `environment.txt` |

---

## 15. Two prior figures that did not reproduce

Recorded as findings in their own right, because a number that does not reproduce
is evidence about the number.

| Prior claim | Measured | Status |
|---|---|---|
| Candidate A needs ~91 continued-fraction iterations | **14**, worst over the whole family | **withdrawn** |
| Candidate C has ~`5.9e-03` normalised error at 4096 nodes | **`9.329e-05`** | **withdrawn** |

Both prior figures were unretained desk estimates, and both were materially worse
than reality — that is, both **overstated** the case against the candidate they
described. Neither is used anywhere in this record. The rejections of A1 and C in
§5.7 rest entirely on the measured values, and had they rested on the prior
figures they would have been wrong.

The `×20` flops-per-iteration multiplier used in an earlier working estimate is
also withdrawn; no such multiplier appears in any retained number.

---

**STEP 0 — ACCEPTANCE REQUESTED**
