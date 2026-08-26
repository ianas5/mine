# PCCM Phase 6 — Step-10A authority record

**Request-fingerprint authority closure. No implementation.**

```
spec/sim_contract.yaml                  the SIM extension's canonical grammar
builder/pccm_builder/sim_loader.py      that grammar, enforced token by token
builder/pccm_builder/sim_emit.py        the framing projected into modSimContract
builder/pccm_builder/sim_cases.py       the golden vectors, through the accepted encoder
docs/phase6_step10_authority.md         this record
```

**Not in this round.** `src/vba/modSimFingerprint.bas` does not exist and is not
in the module registry. There is no `SimFpBuildRequestFingerprint`, no
`SimResultDigest`, no `PCCM_RunSimulation` and no VBA of any kind. **Step 10
implementation remains unauthorised until this package is independently
accepted. Step 11 does not exist. No Windows or Excel runtime ran.**

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-6 / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Accepted Step-7 sampler authority | `f2f654eadba4f5196c795e4167b71f7002e1f727` |
| Accepted Step-8 engine authority | `39415f3` |
| Accepted Step-9 statistics head | `e760c50361f03bce4a393de64614b1cac45d7d29` |
| Step-10A request-fingerprint closure | this commit |

---

## 2. What was found, and what it was not

**Step 9 was accepted.** Nothing in it is reopened here.

The pre-Step-10 review found that `request_fingerprint` locked the SIM
extension's **semantic fields and their order** — and stopped there. It never
locked the **canonical record grammar**: how many records, which encoder each
field uses, what shape AUTO produces versus FIXED, or the token sequence itself.

`result_digest` has carried token-level authority since Step 0:

```yaml
grammar:
  stream:  'F_S("PCCM-RD") F_I(sim_method_version) section'
  section: 'F_S("RESULT") F_I(record_count) record*'
  record:  'F_I(field_count) F_I(iteration_index) F_N(total_nominal) F_N(total_pv)'
```

and the loader compares that mapping **exactly**. The request fingerprint had no
such check, so several byte-distinct streams satisfied the contract at once:

| Undecided | Alternatives the contract allowed |
|---|---|
| iterations | `F_I` or `F_N` |
| the SIM section | one record, or five one-field records |
| the AUTO seed | omitted, or blank `F_S("")`, or `F_I(0)` |
| the versions | integers, or text |
| the field names | absent, or encoded into the record |

**Step 10 must not invent one of these.** Whichever it chose would silently
become the identity of every stored request fingerprint, and a later correction
would make every stored fingerprint STALE. So the grammar is closed first, and
this round closes it.

**This is an authority and emission correction only.** It writes no VBA, chooses
no implementation strategy and adds no continuation API to the accepted encoder.

---

## 3. The locked grammar

The request stream is the accepted Phase-5 canonical prefix, byte for byte,
followed by exactly one SIM section:

```
PCCM-FP  FP_VERSION  HEADER  COST  RISK  |  SIM
```

```
sim_section  ::= F_S("SIM") F_I(1) sim_record

AUTO         ::= F_I(4) F_I(iterations) F_S("AUTO")
                          F_I(rng_version) F_I(sim_method_version)

FIXED        ::= F_I(5) F_I(iterations) F_S("FIXED") F_I(supplied_seed)
                          F_I(rng_version) F_I(sim_method_version)
```

| Field | Type | Present |
|---|---|---|
| `iterations` | `F_I` | always |
| `seed_mode` | `F_S` | always |
| `supplied_seed` | `F_I` | **FIXED only** |
| `rng_version` | `F_I` | always |
| `sim_method_version` | `F_I` | always |

**There are no encoded field names.** The names above are semantic position,
exactly as the Phase-5 HEADER record works, and the loader refuses a grammar
that quotes one.

**Every integer identity is `F_I`, never `F_N`.** A count, a seed and a version
are structural facts; `F_N` would let a version of 1 encode identically to a
Double of 1, which is precisely the distinction the two tags exist to keep.

### 3.1 AUTO means the field does not exist

Not `F_I(0)`. Not `F_S("")`. Not `null`. Not the previous effective seed. Not
the AUTO nonce. AUTO hashes **four** fields; each rejected representation hashes
**five** and is a different fingerprint — `test_10` of the golden-vector file
writes all three out and shows it.

That absence is what makes an AUTO request recomputable while C21 is blank: two
successful AUTO runs of the same current request share one
`request_fingerprint` and both stay CURRENT, even though their `auto_nonce`,
`effective_seed`, `run_id` and `result_digest` all differ.

### 3.2 One stream, one version

The extension carries **no stream tag and no stream version of its own**. It is
a section of the accepted `PCCM-FP` stream, whose tag and `FP_VERSION` belong to
the Phase-5 fingerprint authority. There is no `SIM_FP_VERSION` and no
`REQUEST_FP_VERSION`; `RNG_VERSION` and `SIM_METHOD_VERSION` are **fields inside
the record** and already own the simulation-method compatibility axes.

### 3.3 Type and presence here, admissibility there

`supplied_seed` is `F_I` and is present exactly under FIXED. Its **domain** —
`1 .. 2147483646` — stays with `input_contract.yaml` and is not restated in this
grammar. The contract records the owner (`supplied_seed_domain_owner`) and the
loader refuses any attempt to claim it.

### 3.4 The exclusions are unchanged

`effective_seed`, `auto_nonce`, `run_id` and `selected_confidence_level` remain
excluded, and the loader now additionally refuses a **grammar** that encodes one.
`result_digest`, a timestamp, `model_version`, the computed statistics, the
selected Px and the contingency are absent as well, and the analytical
fingerprint is still **not** hashed as a field.

---

## 4. Enforcement

`_validate_request_sim_grammar` in `sim_loader.py` is now as strict about the
request stream as `_validate_result_digest` has always been about the digest. It
refuses: a record count other than 1; any field type other than the locked one;
`F_N` in place of `F_I`; AUTO carrying `supplied_seed`; FIXED omitting it; a
reordered or moved field; the one-record grammar rewritten as one record per
field; encoded field names; a second SIM record; a repeated `PCCM-FP` or
`FP_VERSION` inside the extension; an excluded field encoded into a production;
and a seed domain claimed by this contract. **No tolerance field exists here,
and every prior exclusion and prefix check is retained.**

The closed-world schema was extended to describe the six new mappings, so a key
the validator does not enforce cannot be added and a key it does enforce cannot
be deleted.

---

## 5. The golden vectors

The analytical prefix is the **accepted Phase-5 case-26 reference stream**, taken
from `calc_cases.reference_stream` unchanged — 366 UTF-16 code units, digesting
to `50B6EB0E26857EA7`. Its HEADER/COST/RISK bytes are not regenerated, not
re-encoded and not hashed as a field.

Every vector is produced through the accepted `calc_fingerprint.py` encoder and
hash. **There is no second hash implementation**, and no continuation API was
added to the accepted encoder.

| Vector | SIM suffix | Request fingerprint |
|---|---|---|
| AUTO / 1000 | `S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1` | `5EAB16E15C2ECE24` |
| FIXED / seed 1 / 1000 | `S3:SIMI1:1I1:5I4:1000S5:FIXEDI1:1I1:1I1:1` | `599C95E7274759B9` |
| FIXED / seed max / 1000 | `S3:SIMI1:1I1:5I4:1000S5:FIXEDI10:2147483646I1:1I1:1` | `0010FB954CC94B53` |
| AUTO / 1001 | `S3:SIMI1:1I1:4I4:1001S4:AUTOI1:1I1:1` | `4777C8BC35F0FFEF` |

All four are pairwise distinct, and so are their streams — the distinctness is
not a hash accident. The three relations the authority names explicitly hold:
AUTO ≠ FIXED seed 1, FIXED seed 1 ≠ FIXED seed max, AUTO 1000 ≠ AUTO 1001.

`2147483646` is read from the input contract through `_seed_domain`, never
restated as a new simulation-contract bound.

**The tests pin every literal.** `tests/test_phase6_request_fingerprint.py`
carries the suffixes, their code-unit counts and the digests as constants, so an
emitter test cannot self-prove by calling the same helper twice. If a version
bumps, `test_14` fails and the literals must be re-pinned deliberately.

---

## 6. The corpus gained the family it was missing

`build/phase6_cases.json` carried `E_digest` but **no machine-readable
request-fingerprint family at all**. That omission is corrected now, before
implementation:

```
I_request_fingerprint   The request fingerprint and its SIM extension
  request_fingerprint.grammar          the locked SIM-extension grammar
  request_fingerprint.auto.1000        AUTO at the business minimum
  request_fingerprint.fixed.seed_1     FIXED at the lowest accepted seed
  request_fingerprint.fixed.seed_max   FIXED at the highest accepted seed
  request_fingerprint.auto.1001        one more iteration is a different request
```

All five are **EXACT**. Each carries the accepted per-case version block. No
timestamp, no random data, no environment data, no tolerance. **No existing case
value or comparison policy was altered** — §7.2 proves it.

---

## 7. Verification

### 7.1 Python suite

```
2798 passed, 0 failed          (1146.04s)
2798 collected
```

| Count | What |
|---|---|
| 36 | request-fingerprint authority tests — 10 contract (`test_70`–`test_79`), 19 golden-vector, 7 Stage-A projection/corpus |
| 26 | request-fingerprint mutation controls (`test_130`–`test_150`) |
| 96 | corpus cases in 11 groups (was 91 in 10) |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2736** to **2798**: +62, and the arithmetic closes exactly — 19 golden-vector
tests, 10 contract tests, 7 Stage-A tests and 26 mutation controls**. **No test was deleted,
skipped or weakened.** Two exact-inventory expectations in
`test_phase6_stage_a.py` legitimately changed and are named here so the edits
stay visible: `test_78`'s group count 10 → 11, and `test_84`'s EXACT count
61 → 66 and total 91 → 96. `test_84` additionally now asserts that all five
added cases are EXACT, so "no case moved class" is checked rather than assumed.

**Every control was checked for vacuity.** Each of the 25 contract mutations was
re-run with its refusal message captured, and every one is refused by
`request_fingerprint` — by name, at the exact rule — rather than by an unrelated
pre-existing check.

### 7.2 The accepted-artefact movement is completely attributable

Two previously accepted Step-5 artefacts moved, as the authority anticipated.
Both were rebuilt from the Step-9 sources and compared:

**`build/vba/modSimContract.bas`** — the diff is **23 added lines and nothing
else**: one section header and the 20 `SIM_REQUEST_*` constants. No existing
constant changed, moved or was removed.

**`build/phase6_cases.json`** — compared case by case:

```
cases added      5     (the I_request_fingerprint group)
cases removed    0
cases changed    0
group order      preserved; I_request_fingerprint inserted before R_runtime
top-level        case_count 91 -> 96, and nothing else
```

### 7.3 New frozen hashes

| Artefact | SHA-256 | Was |
|---|---|---|
| `build/vba/modSimContract.bas` | `42db60ea65d3f8de7a1bfbfc7a3bc2bf77a4395cb974d553419fc6c4326e085f` | `c7e7a784…10c61be` |
| `build/phase6_cases.json` | `fccd3551277e4951b5308d8b281fb3f69d258b9107519745e5a0ec79e1ca225e` | `5551606f…6ec32c` |

The old hashes are **no longer current** and this record does not pretend
otherwise.

### 7.4 Unchanged

| Artefact | SHA-256 |
|---|---|
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` |
| `src/vba/modSimEngine.bas` | `f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab` |
| `src/vba/modSimStats.bas` | `98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0` |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` |
| `build/stage_b_manifest.json` | `0c413d93a0f2d002319584e4d59ce6c36dc612cb4115afcc898f7b8801720053` |

The manifest is **byte-identical**: no registry change and no forbidden-rule
change. The module registry is the same twenty modules Step 9 closed with, and
the structured D6-11 projection is unmoved —
`MRG32k3a → [modSimRng]`, `RunSimulation → []`, `Percentile → []`.

`git diff e760c50 -- src/vba/ spec/input_contract.yaml spec/workbook.yaml
spec/calc_contract.yaml spec/driver_contract.yaml spec/structure_contract.yaml
evidence/ bootstrap/ builder/pccm_builder/calc_fingerprint.py
builder/pccm_builder/sim_rng.py builder/pccm_builder/sim_sample.py
builder/pccm_builder/sim_oracle.py builder/pccm_builder/sim_stats.py` is
**empty**.

---

## 8. Historical records are left as history

`docs/phase6_step0.md`, `phase6_step1.md` and `phase6_step5.md` are **not
rewritten**. They record what was settled and accepted at the time, gap included.
This document is the record that the gap existed, when it was found and how it
was closed; pretending it was never there would destroy the only evidence that
the review process works.

---

## 9. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed, not touched.

**Two stale descriptive `"15"` strings** in
`bootstrap/windows/phase4_functional_test.ps1` and
`docs/phase5_gate_b_harness.md`. Neither is an executable gate.

**The raising arm of `SimStatsLadderExtent`** — a genuinely never-sized VBA
array — still has no Linux execution proof and remains deferred to Gate B.

**Carried to Step 10 implementation:** the byte grammar is now locked, so the
implementation strategy is free to be chosen on its merits — a continuation API,
a rebuilt stream, or something else. This round deliberately chose none of them.

**Carried to Step 11 (unchanged from Step 9):** `modSimReport` must prove and
source-test that the ladders it uses originate from `SimStatsDescribe` and are
not mutated between description and selection; and the nominal and PV selected
values and contingencies must be staged locally and committed together.

---

## 10. Step-10A acceptance gate — self-check

| Gate condition | Status |
|---|---|
| the request fingerprint has one exact canonical grammar | §3, `test_75` |
| AUTO and FIXED have unambiguous, different record shapes | `test_73`, `test_09` |
| every field has one locked canonical type | `test_71` |
| AUTO seed means ABSENT, not a sentinel | `test_74`, `test_10` |
| the HEADER/COST/RISK bytes are untouched | `test_79`, `test_01`, `test_02` |
| `FP_VERSION` remains the only PCCM-FP stream version | `test_76`, `test_86` |
| the exact golden AUTO/FIXED fingerprints are pinned as literals | §5, `test_04`–`test_08` |
| `phase6_cases.json` carries machine-readable request vectors | §6, `test_15`–`test_17` |
| `modSimContract` projects the required framing constants | `test_85` |
| the framing is projected, not spelled in a second place | `test_86`, `test_87` |
| the accepted Phase-5 encoder was not reopened | `test_19`, §7.4 |
| 20 required mutations are controlled | 26 controls, §7.1 |
| no VBA implementation exists | `test_18` |
| no module registry or D6-11 movement | §7.4 |
| no Step 11 exists | no `modSimReport`, no publication |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B debt | **OPEN**, §9 |
