# PCCM Phase 6 — Step 5 authority record

Step 5 adds **Stage-A simulation artefact emission** and nothing else:

```
build/vba/modSimContract.bas    a constants-only projection of the accepted
                                simulation authorities
build/phase6_cases.json         the machine-readable conformance corpus the
                                later VBA implementation steps consume
```

Both are **generated artefacts**. They create no modelling authority: every
value in them is read from an owner that already existed.

> **CORPUS CORRECTION ROUND.** Independent review of `d0fb374` accepted the
> emission architecture and found three completion issues plus two record
> corrections: semantic numbers were emitted as strings, which made
> `allow_nan=False` a hollow guard; the corpus had no machine-readable type
> validation; each case did not carry its own version identity;
> `model_version` appeared only as a comment; and the public-API claim was
> imprecise. All five are addressed here and **no accepted number moved**. See
> §11.

**Not in this step.** No `modSimRng`, no VBA generator, no VBA sampling, no VBA
simulation engine, no VBA statistics, no VBA fingerprint implementation, no VBA
reporting, no `_SimData` materialisation, no Results publication, no Gate B, no
Windows or Excel runtime. Step 6 has not begun.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-5 executable baseline | `f571154118083e569e1fb9fbf9bf72852cc2d568` |
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-0 authority / evidence | `49effe03b88ff113a49166d31065d1f4596f2936` |
| Accepted Step-1 contract authority | `35c2467c1f0852fd6cbe5285600c96baeedca2de` and its accepted hardening history |
| Accepted Step-4 oracle head | `614be1acf0f69c16443ace5381edf6157e0f57d3` |
| Step-5 Stage-A emission | this commit |

**No spec change was needed and none was made.** `spec/`, `src/`, `bootstrap/`
and `evidence/` are byte-identical to `614be1a`, as are every accepted oracle
module (`sim_rng.py`, `sim_sample.py`, `sim_oracle.py`, `sim_stats.py`,
`calc_oracle.py`, `calc_numeric.py`, `calc_fingerprint.py`, `calc_emit.py`,
`calc_cases.py`) and every earlier phase record.

---

## 2. Files

| File | Change |
|---|---|
| `builder/pccm_builder/sim_emit.py` | **NEW** — renders both artefacts |
| `builder/pccm_builder/sim_cases.py` | **NEW** — builds the case corpus |
| `builder/pccm_builder/__init__.py` | exports `emit_sim_artifacts` |
| `builder/build_stage_a.py` | calls the emitter; the stale "EMITS NOTHING" comment corrected |
| `tests/test_phase6_stage_a.py` | **NEW** — 84 tests, of which 20 are mutation controls |
| `docs/phase6_step5.md` | **NEW** — this record |

---

## 3. `modSimContract.bas` — constants only

**182 `Public Const` declarations. Zero procedures.** (The count is reported,
not asserted as authority: the tests check the semantic set, and assert only
that at least the projected constants are present.)

A test walks every line of the module and asserts that each one is the
`Attribute VB_Name`, `Option Explicit`, a comment, blank, or a `Public Const` —
there is nowhere for an implementation to hide. A second test asserts the module
declares no `Sub`, `Function` or `Property`, and a third that no arithmetic
operator, no `Exp`/`Log`/`Sqr`, and no `Range`/`Cells`/`Worksheets` appears in
code. There is no recurrence, no jump arithmetic, no sampler, no quantile, no
digest step and no worksheet access.

### Every constant equals its owner

| Owner | Projected |
|---|---|
| `sim_contract.yaml` | versions, generator constants and state order, AUTO cycle and nonce lifecycle, component kinds and roles, stream index origin and accumulation order, jump exponent / `H` / both matrices, the locked acceptance literals, risk and contribution rules, digest framing, state / attempt / seed-mode labels, `_SimData` geometry, run-id limits, statistical method names, contingency formula and baseline |
| `input_contract.yaml` | the FIXED seed domain, the business minimum iteration count, the selectable confidence levels behind the ladder |
| `workbook.yaml` | `SIM_MODEL_VERSION`, projected from `model.model_version` |

Nothing is restated. A test compares each projected constant back to the value
its owner holds; the two that most invite a hardcoded copy — the FIXED seed
domain and the iteration minimum — are read through the same accessors the
oracle uses.

### The model version is a constant

`SIM_MODEL_VERSION` is projected directly from `workbook.yaml:model.model_version`
and is not hardcoded in the emitter. Step 1 requires a successful run to snapshot
`model_version` in its metadata, and a **comment cannot be read by the code that
has to write it** — leaving it in prose would invite the later VBA to declare a
literal of its own, which is the single thing this module exists to prevent. The
banner comment remains, but comments are not authority.

A test binds the constant to its owner, and a second renders the module from a
synthetic manifest carrying `9.9.9-test` and asserts that **exactly one**
`Public Const` line changes.

### VBA type safety

| Value | VBA type | Why |
|---|---|---|
| `m1 = 4294967087`, `m2 = 4294944443` | **Double** | above `Long` range; a `Long` projection would not compile |
| every jump matrix element | **Double** | several exceed `Long`, and a matrix carrying two VBA types would be a trap for the implementation that reads it |
| `AUTO modulus = 2147483647` | Long | fits exactly, and is projected as a Long so the rule stays a rule rather than a blanket widening |
| `FIXED seed maximum = 2147483646` | Long | fits |
| `SIM_MAX_ITERATIONS = 1048543` | Long | fits |

A test asserts that **no** constant projected as `Long` lies outside
`[-2147483648, 2147483647]`, and that every `Double` literal is the canonical
round-tripping form with no locale separator anywhere. The literal renderer is
the proven Phase-5 discipline: `repr` gives the shortest decimal that
round-trips, an integral value gets an explicit `.0`, and the output is
locale-free by construction.

### D6-11 is NOT activated

The generated module passes the **current** forbidden-construct guard with no
entry scoped and **no change to `structure_contract.yaml`**.

That required leaving two things out of the projection entirely — not as
identifiers, not as string values, not as commentary:

- the RNG **family name**, replaced by neutral `SIM_RNG_*` names that carry the
  same information;
- the future simulation **endpoint name**.

The accepted Quantile naming discipline is kept, so the globally forbidden
quantile token never appears either. Tests scan the module twice: through the
accepted `contains_construct` scanner over code, and over the **raw text**
including comments, which the code scan deliberately does not see.

### No premature registry expansion

`modSimContract` is **not** added to the Stage-B module registry, and no
`modSimRng` / `modSimSample` / `modSimEngine` / `modSimStats` /
`modSimFingerprint` / `modSimReport` entry was created.

**No authority conflict arose.** The Windows Stage-B script iterates
`$manifest.vba.modules` rather than globbing the generated directory, so an
artefact with no declared Phase-6 owner is simply not imported. The generated
module can therefore sit in `build/vba/` as a Stage-A artefact until the
implementation step that first introduces a real Phase-6 VBA module performs the
atomic D6-11 activation. A test asserts the module is absent from the registry
and that every forbidden rule is still global.

---

## 4. `phase6_cases.json` — the conformance corpus

**91 cases in 10 ordered groups.** Test data: not the runtime contract, not a
second engine, not a benchmark, not a stochastic recommendation, not workbook
data.

| Group | Cases | Subject |
|---|---|---|
| `A_rng` | 11 | seeds, states, uniforms, illegal states |
| `B_jump` | 2 | canonical stream assignment and its row-order invariance |
| `B_seed` | 13 | FIXED domain, the AUTO cycle and its exhaustion boundary |
| `C_sampler` | 19 | Uniform / Triangular transforms, Cheng BB and BC, Bernoulli |
| `D_engine` | 11 | whole seeded simulations |
| `E_digest` | 8 | the seven retained framing vectors plus the grammar |
| `F_statistics` | 13 | mean, deviation, Type-7 vectors, the ladder, constant samples |
| `G_contingency` | 4 | the reporting lookup |
| `H_domain` | 7 | the numerical domain end to end |
| `R_runtime` | 3 | behaviours needing machinery Step 5 does not build |

### No second implementation

Stated precisely, because the earlier wording was not:

- **The stochastic expectations use the accepted Step-2/3/4 public reference
  APIs** — `RngReference`, the Step-3 samplers, `prepare_simulation` /
  `run_simulation`, `result_digest`, `describe`, `contingency_at`.
- **Fixture materialisation reuses the existing Phase-5 adapters** `to_model`
  and `tolerances_from` from `calc_cases.py`. These are *not* re-exported from
  `pccm_builder.__init__`, so the earlier claim that the emitter "calls only the
  public Step-2/3/4 exports" was imprecise and is withdrawn.
- **No test module is imported.**
- **No calculation formula is duplicated**, and there is no generator, sampler,
  accumulation or statistic of the emitter's own.

Reusing those two adapters does not constitute a second calculation engine: they
turn a plain fixture payload into the accepted `CalculationModel` and
`Tolerances` value types, which is exactly the duplication Step 5 is meant to
avoid. `__init__.py` was **not** expanded to re-export them — doing so purely for
cosmetic purity would change the Phase-5 public API surface for no semantic
reason.

### The layered evidence model is preserved, not strengthened

| Policy | Cases | Meaning |
|---|---|---|
| `EXACT` | 61 | bit-for-bit equality required across implementations |
| `TOLERANCE_BOUNDED` | 23 | agreement within a bounded tolerance |
| `STATISTICAL` | 1 | distributional agreement only; sample-for-sample identity NOT claimed |
| `SAME_RUNTIME_ONLY` | 3 | a relation between two runs in one runtime, not a value to reproduce |
| `RUNTIME_ONLY` | 3 | needs publication machinery that does not exist yet |

Mapped onto the accepted plan §15: layer A and B → `EXACT`; layer C →
`TOLERANCE_BOUNDED` with individually exact rows marked case by case; layer D →
`EXACT`; layer E → draw counts and states `EXACT`, transformed values
`TOLERANCE_BOUNDED`; F1 → `TOLERANCE_BOUNDED`; F2 → `EXACT`; G →
`STATISTICAL`; G2 and G3 → `SAME_RUNTIME_ONLY`.

A case may additionally carry an **`expected_exact`** block for the fields that
are exact whatever the case policy is — a draw count and a post-sample RNG state
are exact even where the transformed sample they produced is only bounded. The
F1 whole-engine case therefore carries its `result_digest` under `expected`
(bounded), and the F2 fixtures carry theirs under `expected_exact`.

**The numeric tolerance is deliberately absent.** `sim_contract.yaml` forbids a
comparison tolerance outright, so a number invented in this corpus would be a
new authority. The class is stated; the value is left to the evidence policy
that owns it.

**Layer G is not strengthened because the Python oracle now exists.** The
Beta-containing whole-engine case is `STATISTICAL`, its digest appears only
under `python_reference`, and a mutation control asserts that promoting it to
`EXACT` is detectable.

### The exact-friendly fixtures

Two F2 fixtures, both exact at every step:

- **`engine.exact_friendly.unit_interval`** — one `Uniform(0, 1)` Cost Line with
  Quantity 1 and `Knom = Kpv = 1`. `(1−u)·0 + u·1` is `u` bit for bit, and a
  one-term accumulation cannot round, so each retained total **is** the drawn
  uniform and the digest is exact because the uniform it hashes is exact. Still
  genuinely stochastic, and there is no rejection path.
- **`engine.exact_friendly.dyadic_mixed`** — small dyadic rationals throughout:
  `512 + 4 − 8 = 508` without the Risk, `572` with it. The Uniform's populated
  Most Likely of `999` reaches nothing.

### Retained vectors, pinned independently

Every retained Step-0 vector family is emitted by calling the oracle and then
**compared in the tests against the evidence file**, not against a second call to
the implementation that produced it:

| Evidence file | Pinned |
|---|---|
| `seed_vectors.json` | 4 seed→state expansions, 7 nonce→seed mappings, the exhaustion boundary |
| `rng_vectors.json` | initial state and first five uniforms for seeds 1, 2, 12345, 2147483646 |
| `jump_vectors.json` | streams 0, 1, 7, 399, 401 — initial states and first five uniforms; both jump matrices |
| `stream_assignment_vectors.json` | all 400 components of family A (200 Cost Lines, 100 Risks), first ten and last four |
| `digest_vectors.json` | all seven digests **and** their canonical streams |
| `cheng_vectors.json` | all five cases × 24 samples — values, per-sample attempts, cumulative uniforms, post-states, initial and final states |

The seven digests are additionally pinned as **literals** in the test file, so
the corpus cannot drift silently even if an evidence file were ever regenerated.

### Semantic numbers are JSON numbers

Every finite semantic number is emitted as a **JSON number**, never as text.
`repr(...)` no longer appears in the corpus builder: `_n()` returns a Python
float and refuses a non-finite value by name, so the value stays a float all the
way to `json.dumps`.

**That is what makes `allow_nan=False` real.** `repr(float("nan"))` is the
ordinary string `"nan"` and serialises without complaint; a textual scan for
`NaN`/`Infinity` never saw it. With the numbers actually numeric, the serialiser
is a fail-loud boundary, and mutation controls plant `nan`, `+inf` and `-inf` in
a numeric leaf and require emission to fail.

**Exact binary identity has an explicit sidecar.** Where a block is `EXACT`,
every float leaf gains a sibling key `<name>_canonical` holding the **accepted
Phase-5 canonical-number encoding** — `canonical_number` from
`calc_fingerprint.py`, not a new floating format. So:

```json
"first_uniforms":           [0.12701112204657714, 0.3185275653967945, ...],
"first_uniforms_canonical": ["1.2701112204657714E-01", "3.1852756539679450E-01", ...]
```

The semantic number stays a JSON number; the sidecar is an encoding and is
labelled as one. A `TOLERANCE_BOUNDED` block gets **no** sidecar — an exact
textual form there would invite precisely the comparison the accepted evidence
model refuses.

**Retained Step-0 evidence is not rewritten.** Those files store decimals as
strings and stay exactly as they are. The emitter never reads them; at the
projection boundary the tests parse the retained string to the accepted Double
and require the corpus's JSON number to be that Double.

**741 numeric leaves were converted, and every one parses to the identical
binary64.** Verified once against the previous corpus built from `d0fb374`:
741 converted, 741 identical, **0 moved**, 0 comparison policies changed. The
durable form of the guarantee is a permanent test that the corpus is lossless and
stable under a JSON round trip.

### Numeric-looking strings that remain — and why

272 strings still parse as numbers. **Every one is a declared encoding**, and a
test asserts that no numeric-looking string exists outside a canonical sidecar:

| Category | Legitimately textual because |
|---|---|
| `*_canonical` values (all 272) | the accepted Phase-5 canonical-number encoding, named as an encoding by its key suffix |

Other textual fields do not parse as numbers at all and are unaffected: case ids,
labels, enum values, formula and grammar productions, digest hex, canonical
streams, refusal stages and reasons, and version strings such as `"1.0.0"` and
`"0.5.0"`.

### Machine-readable type validation

`validate_corpus()` runs at the end of **every build**, so a typing regression
fails Stage A rather than waiting for a reader. It is an artefact-shape
validator, not another production contract.

| Rule | Effect |
|---|---|
| generic | a string that parses as a number must sit under a key ending `_canonical` |
| integer keys | counts, indices, states, seeds, nonces, row numbers must be `int` and never `bool` |
| number keys | uniforms, samples, probabilities, α/β, means, deviations, quantiles, totals, contingencies, `p`/`h`/`f` must be `int`/`float` |
| number-list keys | states, uniform vectors, windows, distinct totals, quantile mappings must hold numbers |
| sidecars | a `_canonical` value must be text |
| declared text subtrees | `grammar` and `formula` are text by declaration — their inner names (`stream`, `h`, `lo`) collide with numeric ones by coincidence |
| finiteness | no float leaf may be non-finite |
| version identity | every case carries a complete block equal to the artefact's |

Mutating `1.25` to `"1.25"` anywhere numeric is detected, as is the string
`"nan"`.

### Every case is self-identifying

Each of the 91 cases carries:

```json
"versions": {
  "model_version": "0.5.0",
  "sim_contract_version": "1.0.0",
  "rng_version": 1,
  "sim_method_version": 1
}
```

These are **projections** of the four top-level values — literally assigned from
the same objects — not independent copies. A case lifted out of the file now says
what it applies to instead of relying on its surroundings. `FP_VERSION` is
deliberately absent: no accepted authority requires it for these cases. Tests
assert equality with the top level, and controls detect a changed one-case
`rng_version` and a removed block.

`schema_version` moves from **1 to 2** to mark the representation change.

### Determinism

No timestamp, no absolute path, no hostname, no environment data, no random AUTO
seed, no set iteration order, no `NaN`, no `Infinity`. `allow_nan=False` makes a
non-finite value a build failure rather than a non-standard token. Two emissions
are byte-identical, asserted by SHA-256 for both artefacts.

### Bounded generation

`ENGINE_ITERATIONS = 1000` — the business minimum, and the smallest run the
accepted vector families admit. The Step-0 feasibility benchmark, the
20,000-iteration analytical cross-checks and the 100,000-iteration performance
model are **retained evidence** and are not regenerated on every build. Emission
takes about two seconds and the corpus is 214 KB. A test asserts the corpus
carries no performance claim as an expected value.

---

## 5. Emitted artefacts

| Artefact | Bytes | SHA-256 |
|---|---|---|
| `build/vba/modSimContract.bas` | 13,699 | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` |
| `build/phase6_cases.json` | 251,702 | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` |

Both moved from the reviewed `d0fb374` values by design: the module gained
`SIM_MODEL_VERSION`, and the corpus gained JSON numbers, canonical sidecars and
per-case version identity. No numerical value changed.

---

## 6. What Step 5 did not disturb

The Stage-A build now emits two more files and rewrites no Phase-4 or Phase-5
generated authority. Proved by building the same tree twice — once normally and
once with only the Step-5 emission call removed — and comparing:

```
build/vba/modConstants.bas         IDENTICAL
build/vba/modCalcContract.bas      IDENTICAL
build/stage_b_manifest.json        IDENTICAL
build/phase4_scenarios.json        IDENTICAL
build/phase5_cases.json            IDENTICAL
build/phase5_gate_b_inspection.json IDENTICAL
```

The same six are also byte-identical to a build of the accepted Step-4 commit
`614be1a`.

**The workbook gains nothing.** No `vbaProject.bin`, no simulation result, no
formula, no `_SimData` row and no Phase-6 publication; `_SimData` remains
structurally empty under the existing workbook authority, and
`modSimContract.bas` is an external file that no part of the `.xlsx` contains.
Whole-workbook byte identity is **not** claimed — the Step-1 Random Seed
validation change already affects Stage A legitimately, and the build stamp
varies by design.

---

## 7. Test inventory

`tests/test_phase6_stage_a.py` — **64 tests**.

| Range | Subject |
|---|---|
| 1–6 | module shape: name, `Option Explicit`, banner, constants only, zero procedures, no algorithm |
| 7–18 | every constant against its owner: versions, generator, Long/Double safety, Double round-trip, jump matrices, acceptance literals, seed domain, iteration bounds, `_SimData` geometry, digest framing, labels, the ladder |
| 19–23 | scope: no forbidden construct in code or commentary, no endpoint or family name, no D6-11 exception, not in the Stage-B registry, deterministic |
| 24–30 | corpus shape: valid deterministic JSON, no non-finite value, no environment data, version identity, unique ids and policies, all five classes used, no Python reprs |
| 31–37 | every retained vector against the **evidence file**, plus the seven digests as literals |
| 38–50 | accepted semantics: ignored Most Likely, D6-18b, row-order, seed scope, exact-friendly fixtures, Quantity once, the full ladder, constant samples, Type-7 row policies, contingency, extreme refusals, no performance claim, bounded size |
| **51–59** | mutation controls: vectors, policies, Quantity, module constant, planted construct and procedure |
| 60–64 | the normal build emits both, prior artefacts unmoved, workbook free of Phase 6 |
| 65–69 | semantic typing: no stringified number, representative fields are floats, counts stay integers, every exact number has a canonical sidecar, a bounded block has none |
| **70–77** | non-finite rejection: `nan` / `+inf` / `-inf` refuse emission, the builder refuses first, `"nan"` and `"1.25"` fail validation, `1.25 -> "1.25"` anywhere is detected, the validator runs on every build |
| 78–79 | per-case version identity, and its two mutation controls |
| 80–82 | `SIM_MODEL_VERSION`: bound to its owner, follows a synthetic owner, module still constants-only |
| 83–84 | JSON round-trip losslessness, and the comparison-policy distribution unchanged |

### Mutation controls

| # | Mutation | Caught by |
|---|---|---|
| 51 | a changed RNG initial state | disagreement with `rng_vectors.json` |
| 52 | a changed jump-stream uniform | disagreement with `jump_vectors.json` |
| 53 | a changed result digest | disagreement with `digest_vectors.json` |
| 54 | a changed Cheng draw count | disagreement with `cheng_vectors.json` |
| 55 | a comparison policy strengthened to `EXACT` | the layer-G and layer-F1 policy assertions |
| 56 | Quantity applied twice in an expected row | the linearity assertion |
| 57 | a module constant that stops equalling its owner | the owner comparison |
| 58 | a forbidden construct planted in the module | the accepted scanner |
| 59 | a procedure planted in the module | the procedure scan |
| 70 | `nan` in a numeric leaf | emission refuses |
| 71 | `+inf` in a numeric leaf | emission refuses |
| 72 | `-inf` in a numeric leaf | emission refuses |
| 73 | a non-finite value reaching `_n()` | refused before serialisation, naming the corpus |
| 74 | the string `"nan"` in a numeric field | corpus validation fails |
| 75 | the string `"1.25"` in a numeric field | corpus validation fails |
| 76 | `1.25` turned into `"1.25"` in a contingency row | corpus validation fails |
| 79 | a changed one-case `rng_version`, and a removed version block | corpus validation fails |
| 81 | a synthetic manifest with a different model version | exactly one `Public Const` line moves |
| 84 | any comparison policy moved | the pinned distribution |

Each control mutates a **copy** and shows the accepted artefact still passes, so
none of them is vacuous. Nothing in `spec/` or `evidence/` is touched.

---

## 8. Regression

| Check | Result |
|---|---|
| Full Python suite (Python 3.11.15) | **2319 passed, 0 failed** |
| Before Step 5 | 2235 passed |
| New Step-5 tests | **+84** |
| Step-5 mutation controls | **20** (within the 84) |
| Stage-A build / verifier | **351 passed, 0 failed** |

No test was deleted, skipped or weakened.

### Phase-5 fingerprint vectors

```
fingerprint("PCCM-FP")      6551C6F365DA7F3F
fingerprint_probe(A|B)      42E49DC715F06970
fingerprint_probe(AB|)      7558FD9248656EAD
canonical_number(1/3)       3.3333333333333331E-01
```

Unchanged.

---

## 9. Runtime evidence

**No Windows or Excel execution was performed.** Everything here is Linux
Python 3.11. Step 5 emits text files; nothing runs the generated VBA, and Gate B
was not run.

---

## 10. GATE-B TEMP-DIR CLEANUP DEBT — OPEN

Carried forward unchanged from `docs/phase6_step4.md` §27.

Repeated test execution leaves `pccm-gateb-*` temporary directories under `/tmp`
and never removes them — one full suite run leaves about 385 of them, and an
accumulation of roughly 56,986 once exhausted the session's writable allowance.

**It MUST be resolved before the Phase-6 Gate-B harness extension / Windows
execution step.** It was not fixed here: the Gate-B helper is outside the
authorised Step-5 boundary and was not modified, and Gate B was not run.

---

## 11. The corpus correction round

Independent review of `d0fb374` accepted the Stage-A emission architecture and
confirmed the Step-5 suite, Stage-A, both artefact hashes, the package digest and
the byte-identity of all six prior artefacts. Five items were raised.

| # | Issue | Correction |
|---|---|---|
| 1 | ~741 semantic numbers emitted as strings via `repr` | every finite semantic number is now a JSON number; exact blocks carry an accepted-encoding canonical sidecar; retained Step-0 evidence is parsed, never rewritten (§4) |
| 2 | `allow_nan=False` was hollow while numbers were text | numbers stay Python floats to the serialiser; `_n()` refuses a non-finite value first; five mutation controls prove both boundaries (§4) |
| 3 | no machine-readable type validation | `validate_corpus()` pins integer / number / list / text / sidecar types and runs on every build (§4) |
| 4 | cases were not self-identifying | every case carries a `versions` block projected from the artefact's four values, with two mutation controls (§4) |
| 5 | `model_version` appeared only as a comment | `SIM_MODEL_VERSION` is projected from `workbook.yaml`, with an owner-binding test and a synthetic-owner control (§3) |

Plus one record correction: the public-API claim is now stated precisely (§4) —
`to_model` and `tolerances_from` are existing Phase-5 adapters, are not public
exports, and their reuse is not a second engine. `__init__.py` was not expanded.

**No accepted number moved.** 741 leaves converted, 741 parse to the identical
binary64, 0 moved, and the comparison-policy distribution is unchanged at
61 / 23 / 1 / 3 / 3. The prior six Stage-A artefacts remain byte-identical, and
the workbook gained no Phase-6 content.

**No architecture was reopened**, no D6-11 exception exists, `modSimContract`
remains constants-only, and Step 6 has not begun.

---

**STEP 5 — ACCEPTANCE REQUESTED (CORPUS CORRECTED)**
