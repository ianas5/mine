# PCCM Phase 6 — Step 10 authority record

Step 10 adds **one module** and **one procedure to an accepted one**:

```
src/vba/modSimFingerprint.bas    the request fingerprint and the result digest
src/vba/modCalcFingerprint.bas   + CalcFpContinueDigest, APPENDED after every accepted line
spec/structure_contract.yaml     modSimFingerprint enters the registry
```

**Not in this step.** No simulation execution, no statistics, no AUTO nonce
allocation, no `run_id`, no state or attempt logic, no `_SimData`, no Results,
no publication and no `PCCM_RunSimulation`. `modSimReport` does not exist.
**Step 11 has not begun. No Windows or Excel runtime ran.**

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-6 / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Accepted Step-7 sampler authority | `f2f654eadba4f5196c795e4167b71f7002e1f727` |
| Accepted Step-8 engine authority | `39415f3` |
| Accepted Step-9 statistics authority | `e760c50361f03bce4a393de64614b1cac45d7d29` |
| Accepted Step-10A request-fingerprint grammar | `34a7c467a2e22c3f896cdc10487a1b3922b4536b` |
| Step-10 modSimFingerprint | this commit |

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

**SOURCE CONFORMANCE — proved now, on Linux.** Purity, the public surface, the
exact SIM suffix bytes, the request fingerprint as a **continuation** of the
analytical hash state, the result-digest framing and order, and the arithmetic
those statements describe — against the accepted Step-10A authority and the
accepted `phase6_cases.json` corpus.

**VBA EXECUTION CONFORMANCE — NOT proved, deferred to Gate B.** No VBA runtime
exists in this step. No digest here may be read as "VBA produced this
fingerprint".

### 2.1 The whole hash core is compiled from real source

Unlike Steps 7–9, where the accepted arithmetic primitives were borrowed, Step
10 transcribes the accepted hash core itself: `CalcFpReduceDouble`,
`CalcFpNormaliseCodeUnit`, `CalcFpDigestStream`, `CalcFpHex8`, `CalcFpPowerOf16`,
`CalcFpCanonicalText`, `CalcFpCanonicalInteger`, `CalcFpNumberField`,
`CalcFpField`, `CalcFpDigitsOf` and the new `CalcFpContinueDigest`,
`CalcFpHexValue`, `CalcFpHexDigitValue` — all from `modCalcFingerprint.bas`.

**Two procedures are borrowed**, each for a stated reason, each with its real VBA
signature read out of the source:

| Borrowed | Why |
|---|---|
| `CalcFpCanonicalNumber` | its second tier is the exact-integer limb machinery rebuilt at Gate B Runtime Run 2 — dynamic limb arrays and scoped rounding the transcriber does not model. Its accepted **Python** counterpart is bound instead. |
| `SimFpRetainedExtent` | reads a bound of an unproven carrier under a **scoped** `On Error`, and the engine models no error handling. The shim reproduces the **allocated** arm; the arm that RAISES is Gate-B work. |

### 2.2 The transcriber was extended, not replaced

Four mechanical additions, no second interpreter:

```
AscW                    the first UTF-16 code unit as a SIGNED 16-bit Integer
Len / Mid               made UTF-16-EXACT: code units, not Python code points
Mid$ and friends        the `$` suffix stripped before the builtin map runs
For ... Step            the counted loop with an explicit (including negative) step
modXxx.Proc(...)        a module-qualified call is the same procedure
```

`Len` and `Mid` becoming UTF-16-exact is a **fidelity increase**, not a change of
meaning: for ASCII they were already identical, and for astral characters they
were previously wrong in the same direction Python is wrong. Every Step-6,
Step-7, Step-8 and Step-9 test still passes against the extended engine.

---

## 3. Why one procedure was added to `modCalcFingerprint`

`CalcFpBuildFingerprint` returns the analytical **digest** and does not expose
the canonical HEADER/COST/RISK stream that produced it. The Step-10A authority
requires the analytical prefix bytes followed by the SIM section, and **forbids
hashing that digest as a field**. The three ways out were:

1. rebuild the analytical grammar inside `modSimFingerprint` — a second
   definition of the encoding;
2. hash the sixteen-character digest as an `F_S` field — explicitly forbidden,
   and a different stream;
3. write a second polynomial hash loop — a second authority for the mathematics.

All three are worse than the fourth. The Phase-5 digest **is** the final pair of
accumulator states — eight hex digits of `h1` then eight of `h2`, with no
finalisation transform — so

```
ContinueDigest(DigestStream(prefix), suffix) = DigestStream(prefix & suffix)
```

is not an approximation of appending to the stream. It **is** appending to it.
`test_13` proves the identity over six prefixes × seven suffixes, including
empty, ASCII, BMP and non-BMP surrogate-pair text and all four Step-10A SIM
suffixes.

**This changes no fingerprint authority.** The base, both moduli, the reduction,
the code-unit normalisation and the hex conversion are the accepted procedures,
called — not restated.

### 3.1 The addition is APPENDED, and the accepted digests still carry their literals

The new procedures sit after every accepted line, behind a unique banner. The
three frozen digest gates over `modCalcFingerprint` — its byte digest, its
pre-Run-7-rename digest and its executable-body digest — now hash the **accepted
prefix**, the file up to that banner, and **still assert their original
literals**:

```
byte digest              39e80b9ef9252a9822cd57c8ae441b67571ca3725b3d78124bd6af2ddccc4744
pre-Run-7-rename digest  9081dc05bddf052fdcb172a34eed588fef1637b89212b14a515539590e265fcf
executable-body digest   1ea6aa3ca4b9d8ce3a5b8885f6e3ba24b1cfe6da870f25ce2db88e2061084cb3
```

That is deliberately **stronger than re-pinning**. Re-pinning a frozen digest to
whatever the file now hashes to says only "this changed"; this says the accepted
bytes are identical and the only thing beyond them is the named Step-10 block.
`test_49` additionally asserts that the text after the banner declares exactly
`CalcFpContinueDigest`, `CalcFpHexValue` and `CalcFpHexDigitValue` — and nothing
else.

### 3.2 The digest is decoded, never parsed by the host

`CLng("&H" & text)` and every host hex parser are locale- and host-sensitive and
would silently accept lowercase, whitespace, a sign or an `0x` prefix. The
decoder walks the accepted uppercase `FP_HEX_DIGITS` table and compares
**ordinally**, so a representation this module never produces is never accepted
back. The accumulator stays a **Double** throughout: eight hex digits reach
4294967295, outside signed-Long range, and narrowing on the way would be exactly
the conversion `CalcFpReduceDouble` and `CalcFpHex8` exist to avoid.

`h1` and `h2` have **different** moduli, and each must be a residue of its own —
`test_18` shows that the value between the two moduli is legal as `h1` and
refused as `h2`.

---

## 4. `modSimFingerprint.bas`

7 procedures, **two** public, no `Type`.

| Procedure | Does |
|---|---|
| `SimFpBuildRequestFingerprint` *(public)* | validate, build the SIM suffix, continue the analytical hash |
| `SimFpResultDigest` *(public)* | the retained iteration digest, at `SIM_METHOD_VERSION` |
| `SimFpRequestSuffix` | the Step-10A SIM section, from projected constants only |
| `SimFpValidateRequest` | iterations, seed mode, the flag/mode agreement, the seed domain |
| `SimFpVersionedResultDigest` | the framing prefix, then one folded record per iteration |
| `SimFpDigestRecord` | one record: field count, LOGICAL index, both totals |
| `SimFpRetainedExtent` | both carrier extents, under a scoped error handler |

### 4.1 The analytical fingerprint is hash STATE, never data

`analyticalFingerprint` reaches exactly one place: `CalcFpContinueDigest`.
`SimFpRequestSuffix` **cannot see it at all** — `test_27` asserts the identifier
does not appear in that procedure — so there is no path by which it could be
framed as a field. `test_27` also shows behaviourally that hashing it as an
`F_S` field produces a different answer, and `test_28` that starting a fresh
`PCCM-FP` stream at SIM does too.

**WHAT THIS DOES NOT PROVE:** that the analytical fingerprint handed in is
CURRENT. A syntactically valid digest is not evidence that the model has not
moved since it was computed. That obligation is carried to Step 11 in §8.

### 4.2 The result digest is streamed

A run may retain 1,048,543 iterations. The framing prefix is digested once and
each record is folded in as it is built, so the canonical text alive at any
moment is **one record**. `test_46` instruments both hash entry points across a
500-record digest and asserts 501 calls with **no string longer than 200
characters** — a whole-stream build would be tens of thousands.

### 4.3 Indexing: zero-based arrays, one-based identity

Step 8 retains iteration *i* at physical element *i − 1*, and the digest index
origin is 1. Element `offset` is therefore iteration
`SIM_DIGEST_INDEX_ORIGIN + offset`, whatever the carrier's physical `LBound` is.
Encoding a physical index would make the digest depend on how the array was
declared — and because `LBound` is normally zero, that defect is **invisible in
the numbers**. `test_38` is therefore a source detector as well: it requires the
caller to hand the record its **logical** offset.

### 4.4 The production surface takes no version

`SIM_METHOD_VERSION` is the version. A caller who could choose another could
produce a digest claiming a method the run did not use. The version travels as a
parameter only into the **private** `SimFpVersionedResultDigest`, which exists so
the accepted `digest.version_2` framing vector can be source-tested — a vector no
test can reach is not a vector.

### 4.5 Source uniqueness

`modCalcFingerprint` remains the only owner of canonical field encoding, the
UTF-16 hash recurrence, the reduction and the hex conversion. `test_09` asserts
that `modSimFingerprint` reaches exactly five procedures across that boundary —
`CalcFpCanonicalText`, `CalcFpCanonicalInteger`, `CalcFpNumberField`,
`CalcFpDigestStream`, `CalcFpContinueDigest` — and `test_10` that it contains no
`Fix(`, no `Mod`, no `16#`, no modulus, no base and no reducer of its own.
`test_54` asserts no other module frames a Phase-6 stream at all.

---

## 5. Tests

### 5.1 New files

| File | Contents |
|---|---|
| `tests/test_phase6_sim_fingerprint_vba.py` | 56 Step-10 conformance tests |
| `tests/test_phase6_sim_fingerprint_vba_validation.py` | 51 tests — a baseline and 50 mutation controls |

| Group | Tests | Covers |
|---|---|---|
| A | `test_01`–`test_12` | declaration, registry, surface, purity, source uniqueness |
| B | `test_13`–`test_22` | the continuation identity, digest validation, the accepted loop untouched |
| C | `test_23`–`test_34` | the four golden request vectors, AUTO/FIXED shapes, validation, transactional refusal |
| D | `test_35`–`test_48` | the seven digest vectors, order relations, index origin, streaming, refusals |
| E | `test_49`–`test_52` | the accepted Phase-5 encoder and every locked vector |
| F | `test_53`–`test_56` | scope, the transcription boundary, the guarded bounds shape |

### 5.2 Mutation controls

50 controls: **10** on the continuation primitive, **21** on the request
fingerprint and **19** on the result digest. Each damages one of the two sources,
reruns the whole battery under a per-test time budget, and requires a **named**
detector among the refusers.

Four first-draft controls were rebuilt because they were **vacuous**, and one was
withdrawn as undetectable:

* Three source detectors were **string-blind**. `test_08`, `test_09` and
  `test_28` scanned `VbaModule.code`, which strips string literals — so a
  hand-rolled `"E+00"` exponent, an injected `"PCCM-FP"` tag and a
  `SIM_QUANTILE_*` constant all walked straight past. `test_09` and `test_28`
  now scan comment-stripped code with **strings intact**, and `test_08` scans
  **case-insensitively**, because the projected constants are spelled
  `SIM_QUANTILE_*` and a case-sensitive search for "Quantile" finds nothing.
* Substituting `Asc` for `AscW` is **not detectable on Linux**: the transcriber
  models a VBA String as the UTF-16 sequence it is and has no ANSI code page to
  narrow through. That control was replaced by one this harness can genuinely
  see — the suffix consumed one code unit short — and `Asc` versus `AscW` is
  listed as deferred Gate-B work in §9 rather than claimed here.

### 5.3 Existing tests

The new registry member invalidated thirteen inventory expectations and the new
public helper one surface expectation. **None was deleted, skipped or weakened**;
each consumes a **named** Phase-6 inventory or an explicit registry tail, and
`FINGERPRINT_PUBLIC` gained `CalcFpContinueDigest` with its reason recorded
beside it. **No numeric module-count authority was introduced anywhere.**

Five accepted gates over `modCalcFingerprint` were re-aimed at the accepted
prefix as described in §3.1, **keeping their original literals and their original
questions** — `test_44` of the check suite, the frozen-hash loop and rename proof
of the report suite, `test_64j` and `test_86b` of the source suite, and `test_21`
of the canonical-number suite, which still asks whether anything moved outside
the Run-3A boundary constructions rather than widening its allowed-token list.

One Step-10A test moved with its authorisation rather than being deleted:
`test_18` of `test_phase6_request_fingerprint.py` said "no VBA fingerprint module
exists", which was true while Step 10 was unauthorised. It now says exactly what
may exist — `modSimFingerprint` and nothing beyond it — and still refuses
`modSimReport`, `PCCM_RunSimulation`, and either public entry point appearing in
any other module.

---

## 6. Verification

### 6.1 Python suite

```
2905 passed, 0 failed          (785.56s)
2905 collected
```

| Count | What |
|---|---|
| 56 | Step-10 conformance — 12 request-fingerprint (group C), 14 result-digest (group D), 10 continuation, 20 surface/purity/preservation/scope |
| 51 | Step-10 tests — a baseline and 50 mutation controls |
| 66 + 44 | Step-9 statistics — **still green, unmodified** |
| 54 + 42 | Step-8 engine — **still green, unmodified** |
| 74 + 51 | Step-7 sampler — **still green, unmodified** |
| 51 + 32 | Step-6 RNG — **still green, unmodified** |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2798** to **2905**: +107, the 56 conformance tests and
51 controls. **No test was deleted, skipped or weakened.**

### 6.2 Stage A

```
351 passed, 0 failed
Stage A build complete.
```

**No Windows or Excel runtime ran.**

### 6.3 Module registry after Step 10

```
 1. modConstants        generated       12. modCalcFingerprint
 2. modWorkbook                         13. modCalcResolve
 3. modAppState                         14. modCalcCheck
 4. modTimeline                         15. modCalcReport
 5. modDrivers                          16. modSimContract   generated
 6. modProfiling                        17. modSimRng
 7. modInflation                        18. modSimSample
 8. modStructuralCheck                  19. modSimEngine
 9. modCalcContract     generated       20. modSimStats
10. modCalcFactors                      21. modSimFingerprint         <- new
11. modCalcAnalytical
```

No count is hardcoded anywhere; P5-M and P5-D8 remain manifest-driven.

### 6.4 Artefact hashes

| Artefact | SHA-256 | Status |
|---|---|---|
| `src/vba/modSimFingerprint.bas` | `9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda` | **new** |
| `src/vba/modCalcFingerprint.bas` | `2efbb30c6f915c04b9c07adec07e25e11f4b5bd2b98e3efa818631dc510ce847` | **changed** — the appended continuation only (§3.1) |
| `build/stage_b_manifest.json` | `01ca01a80598256f6ada218603032cd4be6c9bb9b86f452fb701dc610172ae57` | **changed, registry only** |
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` | **byte-identical** |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` | **byte-identical** |
| `src/vba/modSimEngine.bas` | `f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab` | **byte-identical** |
| `src/vba/modSimStats.bas` | `98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0` | **byte-identical** |
| `build/vba/modSimContract.bas` | `42db60ea65d3f8de7a1bfbfc7a3bc2bf77a4395cb974d553419fc6c4326e085f` | **byte-identical** |
| `build/phase6_cases.json` | `fccd3551277e4951b5308d8b281fb3f69d258b9107519745e5a0ec79e1ca225e` | **byte-identical** |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | unchanged |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | unchanged |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | unchanged |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | unchanged |

**The accepted prefix of `modCalcFingerprint.bas` still digests to
`39e80b9ef9252a9822cd57c8ae441b67571ca3725b3d78124bd6af2ddccc4744`** — its
accepted byte digest, unchanged (§3.1).

`git diff 34a7c46 -- src/vba/modSimRng.bas src/vba/modSimSample.bas
src/vba/modSimEngine.bas src/vba/modSimStats.bas spec/sim_contract.yaml
spec/input_contract.yaml spec/workbook.yaml spec/calc_contract.yaml
spec/driver_contract.yaml builder/ evidence/ bootstrap/ docs/phase6_step0.md
docs/phase6_step1.md docs/phase6_step5.md docs/phase6_step8.md
docs/phase6_step9.md docs/phase6_step10_authority.md` is **empty**.

### 6.5 The manifest movement is the registry and nothing else

Rebuilt from the Step-10A contract and compared leaf by leaf:

```
keys added      3     (.vba.modules[20].name / .generated / .responsibility)
keys removed    0
values changed  0
```

and the structured forbidden-rule projection is **bit-identical**:

```json
{"construct": "MRG32k3a",      "allowed_in": ["modSimRng"]},
{"construct": "RunSimulation", "allowed_in": []},
{"construct": "Percentile",    "allowed_in": []}
```

Step 10 required no new scoped exception and took none.

### 6.6 The golden vectors, reproduced from transcribed VBA source

Request fingerprints, over the accepted analytical prefix `50B6EB0E26857EA7`:

| Vector | SIM suffix | Request fingerprint |
|---|---|---|
| AUTO / 1000 | `S3:SIMI1:1I1:4I4:1000S4:AUTOI1:1I1:1` | `5EAB16E15C2ECE24` |
| FIXED / seed 1 | `S3:SIMI1:1I1:5I4:1000S5:FIXEDI1:1I1:1I1:1` | `599C95E7274759B9` |
| FIXED / seed max | `S3:SIMI1:1I1:5I4:1000S5:FIXEDI10:2147483646I1:1I1:1` | `0010FB954CC94B53` |
| AUTO / 1001 | `S3:SIMI1:1I1:4I4:1001S4:AUTOI1:1I1:1` | `4777C8BC35F0FFEF` |

Result digests, every accepted `E_digest` vector:

```
digest.base                      3181AF89642DE500
digest.reversed_iteration_order  4E0FEE211853E8F6
digest.nominal_and_pv_swapped    63A0E93074F0C2EA
digest.one_iteration_dropped     0CAC531732B88B2A
digest.one_ulp_perturbation      5DC1A76B56D75EF4
digest.version_2                 7E8D58C46CCDD798
digest.empty                     12ED977808313D71
```

And every locked Phase-5 vector still holds: `fingerprint("PCCM-FP")` =
`6551C6F365DA7F3F`, `fingerprint_probe(A|B)` = `42E49DC715F06970`,
`fingerprint_probe(AB|)` = `7558FD9248656EAD`, `canonical_number(1/3)` =
`3.3333333333333331E-01`, the 366-code-unit reference stream =
`50B6EB0E26857EA7`, and every collision probe and numeric encoding in
`build/phase5_cases.json`.

---

## 7. No cross-module trust claim

A syntactically valid analytical fingerprint is not proof that it is CURRENT.
Step 10 proves **canonical continuation semantics** and nothing more about the
value it is handed.

---

## 8. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed, not touched.

**Two stale descriptive `"15"` strings** in
`bootstrap/windows/phase4_functional_test.ps1` and
`docs/phase5_gate_b_harness.md`. Neither is an executable gate.

**Carried to Step 11 (`modSimReport`), not implemented now:**

1. **The request prefix must be the CURRENT analytical fingerprint.**
   `modSimReport` must recompute or obtain it from the current resolved inputs
   and pass THAT to `SimFpBuildRequestFingerprint`. It may not use the last
   successful stored analytical fingerprint unless the current one has first
   been proved identical.
2. **The digest must come from the arrays about to be published.**
   `result_digest` must be derived from the exact retained arrays Step 11 is
   committing, not from an earlier copy.
3. **Ladder provenance and non-mutation** (from Step 9): prove and source-test
   that the ladders used originate from `SimStatsDescribe` and are not mutated
   between description and selection.
4. **Nominal and PV commit together** (from Step 9): staged locally and
   committed as one, so a later PV refusal cannot publish a nominal-only result.

---

## 9. Deferred to Gate B

Beyond the standing debt, Step 10 adds real-VBA items that are **listed, not
claimed**:

| Deferred | Why it cannot be settled on Linux |
|---|---|
| `AscW` versus `Asc` | the transcriber models a VBA String as its UTF-16 sequence and has no ANSI code page. `Asc` substituted for `AscW` is invisible here (§5.2). |
| the raising arm of `SimFpRetainedExtent` | a genuinely never-sized VBA array raises 9; the shim reproduces only the allocated arm. Its source shape is pinned by `test_56`. |
| the raising arm of `SimStatsLadderExtent` | unchanged from Step 9. |
| `CalcFpContinueDigest` on real Excel | the continuation identity is proved over the transcribed source; parity with real VBA arithmetic and real `AscW` is Gate-B work, exactly as `CalcFpDigestStream` has always been. |

---

## 10. Step-10 acceptance gate — self-check

| Gate condition | Status |
|---|---|
| `modSimFingerprint` exists | yes, 7 procedures, 2 public, no `Type` |
| the request is hash-state continuation, never digest-as-data | `test_27`, `test_28` |
| continuation is owned by `modCalcFingerprint`, not duplicated | `test_09`, `test_10`, `test_12`, `test_54` |
| all four request golden vectors reproduce exactly | `test_23`, `test_24` |
| the AUTO supplied seed is absent, with no sentinel | `test_26` |
| `FP_VERSION` semantics remain Phase-5-owned | `test_28`, §3 |
| every `E_digest` vector reproduces | `test_35`, `test_36` |
| the original retained order is used | `test_40`, `test_41` |
| the digest index is 1..n independent of physical `LBound` | `test_38`, `test_39` |
| no sorting occurs | `test_40` |
| the digest is streamed with bounded canonical-string memory | `test_46` |
| the zero-record framing vector works and reads no bound | `test_37` |
| malformed carriers refuse transactionally | `test_45`, `test_47` |
| no caller-selected production digest version exists | `test_04`, `test_48` |
| every Phase-5 fingerprint is unchanged | `test_49`–`test_52`, §3.1 |
| no statistics, simulation or publication leaks | `test_08`, `test_53` |
| D6-11 scope unchanged | `test_02`, §6.5 |
| `modSimRng`/`Sample`/`Engine`/`Stats`, `modSimContract`, `phase6_cases` unchanged | §6.4 |
| no test deleted, skipped or weakened | §5.3, §6.1 |
| no Step 11 exists | no `modSimReport`, no `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B debt | **OPEN**, §8; new deferred items listed in §9 |
