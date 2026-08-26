# PCCM Phase 6 — Step 9 authority record

*(Second round: ladder-integrity hardening at the selection boundary. The
numerical implementation is not reopened and did not move — §3.4.)*

Step 9 adds **one module**:

```
src/vba/modSimStats.bas         the simulation statistics: moments, quantiles, contingency
spec/structure_contract.yaml    modSimStats enters the registry
```

**Not in this step.** No result digest, no request fingerprint, no SIM
fingerprint, no `run_id`, no AUTO nonce, no `_SimData`, no Results sheet, no
simulation or attempt state, no `PCCM_RunSimulation`, no workbook publication,
no sensitivity and no annual stochastic output. `modSimFingerprint` and
`modSimReport` do not exist. **Step 10 has not begun. No Windows or Excel
runtime ran.**

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Phase-6 planning authority | `03aa5044cb535513976f0ec3840bc332747678c8` |
| Accepted Step-5 Stage-A authority | `fa30424bfe5e782756fc44b0912c4f113781606c` |
| Accepted Step-6 / D6-11 authority | `2ec1844638badb506cdc2b133e0c7db8beb5e781` |
| Accepted Step-7 sampler authority | `f2f654eadba4f5196c795e4167b71f7002e1f727` |
| Accepted dependency-corrected head | `78f42439e5e799d860b24475167b77a1472af43a` |
| Accepted Step-8 engine authority | `39415f3` |
| Step-9 modSimStats, first round | `f44af49` |
| Step-9 ladder-integrity hardening | this commit |

`builder/pccm_builder/sim_stats.py` remains the single definition of these
semantics. `modSimStats.bas` is their VBA implementation and creates no
authority of its own. Not one line of the Python authority was edited.

---

## 2. WHAT IS PROVED NOW, AND WHAT IS NOT

**SOURCE CONFORMANCE — proved now, on Linux.** Purity, the public surface, the
sorting discipline, the scale-safe moment construction, the Hyndman–Fan type-7
position and interpolation, the ladder projection, the selected level and the
contingency subtraction — and the arithmetic those statements describe, against
the accepted Python authority and the accepted Step-5/Step-6 corpus.

**VBA EXECUTION CONFORMANCE — NOT proved, deferred to Gate B.** No VBA runtime
exists in this step. No number here may be read as "VBA produced this
quantile".

**The comparison policy is not strengthened.** Each corpus case is checked under
the policy the corpus assigns it. Nothing was promoted to a stricter policy to
make a test pass.

### 2.1 Four primitives are borrowed, and one of the module's own

`SafeSignedSum`, `SafeDivide`, `SafeMultiply` and `SafeSubtract` have accepted
Phase-5 VBA bodies whose second tier uses an exact-arithmetic UDT with dynamic
limbs and scoped error handlers — constructs the test transcriber does not
model. Their accepted **Python** counterparts in `calc_numeric` are bound
instead, and their real VBA **signatures are read out of `modCalcFactors.bas`**
through `signature_only=`, so the ByRef/ByVal call convention stays the module's
own rather than something retyped in a test. This is the Step-7/Step-8
discipline unchanged. `IsUsableDouble` is still compiled from its real source.

**One procedure of `modSimStats` itself is borrowed the same way.**
`SimStatsLadderExtent` reads a bound of an unproven array carrier under a
**scoped** `On Error`, and the transcriber models no error handling. Its
signature is read out of `modSimStats.bas`, its shim reproduces the *allocated*
arm (`UBound - LBound + 1`) so short, long and empty carriers stay covered
**behaviourally**, and its source shape is pinned statement by statement by
`test_48`. The arm that RAISES — a genuinely never-sized VBA array — has no
Linux execution proof and is **deferred to Gate B**. `test_47` names the
borrowed set exactly and asserts every other procedure was compiled from
source, so nothing can hide behind the exemption.

### 2.2 The transcriber was extended, not replaced

Two mechanical additions, no second interpreter:

```
Asc   ->  _asc    ord() of the first character
Mid   ->  _mid    1-BASED start and a LENGTH, not an end position
```

Both are needed only because the ladder label is decoded digit by digit rather
than through `Val` (§4). Everything else is the Step-8 engine unchanged, and
every Step-6, Step-7 and Step-8 test still passes against it.

**The hardening round changed the transcriber not at all.** It needed no new
construct: the one procedure it cannot execute is bound through the existing
`only=` / `signature_only=` / `extra=` mechanism.

---

## 3. `modSimStats.bas`

17 procedures, **six** public, one `Public Type`.

| Procedure | Does |
|---|---|
| `SimStatsMean` *(public)* | the scale-safe arithmetic mean, in the caller's order |
| `SimStatsSampleStandardDeviation` *(public)* | the scale-safe `n - 1` deviation, two-pass |
| `SimStatsQuantileType7` *(public)* | one type-7 quantile from an unsorted sequence |
| `SimStatsDescribe` *(public)* | count, moments, min, max and the whole ladder from ONE sort |
| `SimStatsSelectedQuantile` *(public)* | the selected confidence level, looked up in that ladder |
| `SimStatsContingency` *(public)* | selected total minus the deterministic base estimate A |
| `SimStatsUsableSequence` | finiteness and count, before any bound is read |
| `SimStatsUsableProbability` | the closed domain `0 <= p <= 1` |
| `SimStatsConstantValue` | is every observation the same Double? |
| `SimStatsUnitScale` | the largest power of two not exceeding `max(|x|)` |
| `SimStatsSortedCopy` | a private copy, then the sort |
| `SimStatsSortAscending` | bottom-up stable merge sort with one scratch buffer |
| `SimStatsQuantileSorted` | the type-7 kernel over an already-sorted sequence |
| `SimStatsValidateLadder` | proves a supplied ladder IS the owner-projected ladder |
| `SimStatsLadderExtent` | both carrier extents, under a scoped error handler |
| `SimStatsLadderLabel` | the projected label at a ladder position |
| `SimStatsProbabilityOf` | `P<number>` decoded to a probability, no `Val`, no locale |

### 3.1 The summary Type is derived reporting output only

`SimStatsMeasure` carries `Count`, `Mean`, `SampleStandardDeviation`,
`Minimum`, `Maximum`, `QuantileCount` and `Described`. Every field is a scalar
the module itself computed; it holds no sample, no sorted copy, no seed, no
stream, no prepared shape and no worksheet anchor. It is **output**, so unlike
the Step-7 prepared Beta shape there is nothing here a caller could forge into a
false success: no procedure reads it back as an input.

### 3.2 It owns no randomness and no simulation

`modSimStats` calls no `modSimRng` procedure, no `modSimSample` procedure and no
`modSimEngine` procedure. It reads no generator constant, no seed, no stream and
no iteration count. It receives a sequence of Doubles and a logical count, and
that is the whole of its input. It never touches a worksheet, never opens a
register and never publishes.

### 3.3 The selection boundary trusts nothing it is handed

The first round accepted a supplied ladder if its length was right, the selected
label was non-blank, was not the fixed rung, and **appeared somewhere in the
caller's own labels**. That last condition is not evidence. These are ordinary
VBA arrays and any caller can write to them, so a caller could insert a rung the
projection never named, select it, and be handed a value the module had no
authority for.

`SimStatsSelectedQuantile` now calls `SimStatsValidateLadder` **before** it
searches for the selected label and **before** it reads any value. That
validator is **structural only** — no sort, no quantile, no mean, no deviation,
no re-`Describe`, no iteration totals, no checksum — and it requires:

* `quantileCount = SIM_QUANTILE_COUNT`;
* both carriers physically that long, read through the guarded
  `SimStatsLadderExtent` so a malformed or never-sized array refuses **in the
  module's own words** rather than escaping as a raw subscript error (a non-zero
  `LBound` is fine and is used consistently);
* at every position, `StrComp(supplied, SimStatsLadderLabel(position),
  vbBinaryCompare) = 0`;
* every supplied value a usable finite Double — **not clamped, not recomputed**.

Because the expected labels come from the one projection-backed
`SimStatsLadderLabel`, there is still **no second label list anywhere**. The
positional equality alone refuses an inserted rung, a replaced rung, two swapped
rungs, a duplicate, a missing fixed rung, a case variation and a trailing space.

### 3.4 What this boundary CANNOT prove — and does not claim

Exactly as with Step 7's prepared UDT: structural validation can prove owner
labels, carrier shape and value finiteness. It **cannot** prove that a finite
`P50` value was not edited from 100 to 101 after `SimStatsDescribe` produced it.
Proving that needs a seal or a second quantile calculation, and **Step 9 takes
neither** — no hash, no checksum, no signature, no recomputation. `test_63`
states the limit as a test rather than leaving it as an assumption, and asserts
the module invents no seal.

The correct internal boundary, recorded:

> `SimStatsDescribe` is the authoritative constructor of the ladder carrier.
> `SimStatsSelectedQuantile` performs cheap structural validation before use.
> The ladder arrays are internal in-project derived-reporting carriers, **not**
> an externally authoritative serialised representation.

### 3.5 The ladder travels in two parallel arrays

`SimStatsDescribe` returns the labels and the values as two `ByRef` arrays sized
from `SIM_QUANTILE_COUNT` at run time. No fixed bound is written anywhere in the
module, so the ladder cannot silently drift from the projected contract, and
`SimStatsSelectedQuantile` looks its answer up in the very array `Describe`
produced rather than recomputing a second quantile from a second sort.

---

## 4. The locked rules

| Rule | Why it is stated rather than assumed |
|---|---|
| **The caller's array is never reordered** | Every sort runs on a private copy. A statistic that silently permuted its input would corrupt the engine's retained iteration totals for every later reader. |
| **One sort for eleven quantiles** | `Describe` sorts once and reads all eleven rungs off that copy. Sorting per rung is eleven times the work for the same numbers, and `test_11` counts the calls. |
| **A bottom-up stable merge, not a quadratic sort** | 100,000 iterations is the design target. The tie rule is `series(fromLow) <= series(fromHigh)` — the left run wins, so equal Doubles keep their arrival order. |
| **Moments in the caller's order** | The mean and the deviation are taken over the ORIGINAL sequence, never the sorted copy: floating-point summation is order-dependent, and reordering would make the reported mean depend on an implementation detail of the sort. |
| **The scale is a power of two** | The largest power of two **not exceeding** `max(x)`, built by exact halving and doubling. Scaling by a power of two is exact in binary floating point, so normalisation introduces no rounding of its own. |
| **The doubling test cannot overflow** | Written `candidate <= largest / 2#`, never `candidate * 2# <= largest`, which would overflow before the comparison at the top of the range. |
| **The constant-sample invariant comes first** | A sample where every observation is identical returns that value exactly as the mean, exactly `+0#` as the deviation and exactly that value at every rung — checked BEFORE any accumulation, so no summation error can appear where there is provably no dispersion. |
| **`n - 1`, never `n`** | The sample standard deviation is the estimator this model reports; fewer than two observations refuses rather than dividing by zero. |
| **No sum of squares of raw values** | The deviation is two-pass over normalised residuals. Neither `Sigma x^2` nor an unguarded original-scale residual exists in the module. |
| **Unrepresentable dispersion refuses** | A varying sample whose deviation cannot be represented is a refusal, never a reported `0#`, which would be indistinguishable from a constant sample. |
| **Type-7 is `h = (n - 1) * p`** | The accepted Hyndman–Fan definition. `Fix` is the floor here because `h >= 0` by construction. |
| **An integral `h` is returned untouched** | Not formed as `1 * low + 0 * high`. This keeps `p = 0` and `p = 1` exact at every magnitude, subnormals included. |
| **An equal bracket returns that value** | Before any interpolation runs. |
| **The convex form, not the difference form** | `(1 - f) * low + f * high`, never `low + f * (high - low)`, whose subtraction can overflow when the bracket spans the range. |
| **The probability domain is closed** | `0 <= p <= 1`, refused outside, never clamped into range. |
| **The ladder is the projection** | Every label comes from a projected `SIM_QUANTILE_*` constant in the projected order. No label is written out as a string literal anywhere in the module. |
| **The fixed rung is reported, not selectable** | It appears in the ladder and is refused as a confidence level; an unknown confidence level is refused rather than defaulting to a rung. |
| **`P<number>` is decoded digit by digit** | `Asc`/`Mid` and an explicit `magnitude * 10 + d`, then `/ 100#`. `Val` and `CDbl` on a string are locale-sensitive, and a locale must not be able to move a confidence level. |
| **Contingency is a `SafeSubtract`** | Selected total minus the deterministic base estimate A, through the accepted primitive rather than a raw `-`. |
| **A negative contingency is preserved** | Never clamped to zero. A selected total below base estimate A is a real, reportable state of the model, and hiding it would hide it from the very reader who needs it. |
| **Transactional output** | Nothing is published on a refusal: no partial ladder, no half-filled summary, no scalar left modified. |
| **A supplied ladder is proved before it is searched** | Membership in a caller-writable array is not evidence that a label is an accepted confidence level. The labels must be the owner projection, position by position. |
| **A malformed carrier refuses, it does not raise** | The one place that reads a bound of an unproven carrier is `SimStatsLadderExtent`, under an `On Error GoTo` scoped to those two reads and cleared immediately — the accepted `modCalcFactors` discipline. There is **no `On Error Resume Next` in this module**. |
| **Selection is still reporting-only** | Zero sorts, zero quantile calculations, no RNG, no simulation — proved by call counters in `test_62`, not only by reading the source. |

### 4.1 Naming discipline

The globally forbidden `Percentile` remains global with `allowed_in: []`, and
Step 9 took **no new scoped exception**. Every executable identifier here says
**Quantile**. There is no `WorksheetFunction.Percentile`, no `PERCENTILE.INC`,
no `WorksheetFunction` of any kind and no `Application` reference at all.

The local for the power-of-two scale is `unitScale`, not `scale`: `scale` is on
the accepted Phase-5 VBA reserved-identifier list, and a declaration using it
would fail `test_86`/`test_87` of `test_phase5_vba_source.py`. The merge-sort
locals (`series`, `scratch`, `runLength`, `lowEnd`, `midPoint`, `highEnd`,
`fromLow`, `fromHigh`, `target`) were chosen against the same list.

---

## 5. Tests

### 5.1 New files

| File | Contents |
|---|---|
| `tests/test_phase6_sim_stats_vba.py` | 66 Step-9 conformance tests |
| `tests/test_phase6_sim_stats_vba_validation.py` | 44 tests — a baseline and 43 mutation controls |

**Group A — declaration, surface and purity** (`test_01`–`test_09`): the
attribute and `Option Explicit`; the registry entry with nothing beyond it and
D6-11 untouched; exactly the six public numerical operations; the summary `Type`
as derived scalar output only; no workbook or environment reach; no module-level
or `Static` state; the forbidden word absent from comment- and string-stripped
code; no knowledge of seeds, streams, iterations, digests or fingerprints; and
no other module owning a statistic.

**Group B — sorting** (`test_10`–`test_15`): the caller's array is never
reordered; one sort for eleven quantiles; the standalone quantile sorts its own
copy once; the sort is a bottom-up merge and carries no quadratic pattern, no
COM sort and no library sort; it orders correctly at every shape including
signed zeros and the full magnitude range; and it finishes on 100,000 values.

**Group C — the mean and the scale** (`test_16`–`test_20`), **Group D — the
deviation** (`test_21`–`test_27`), **Group E — type 7** (`test_28`–`test_32`),
**Group F — the ladder and the selected level** (`test_33`–`test_37`), **Group
G — contingency** (`test_38`–`test_41`), **Group H — transactional output**
(`test_42`–`test_45`), **Group I — the corpus** (`test_46`, `46a`, `46b`, `46c`,
`test_47`).

**Group J — ladder integrity at the selection boundary** (`test_48`–`test_63`),
added by the hardening round:

| Test | Proves |
|---|---|
| `test_48` | the guarded-bounds helper is scoped statement by statement; exactly one `On Error GoTo` in the module and no `On Error Resume Next` |
| `test_49` | a genuine ladder still answers every selectable level with the stored value, unchanged |
| `test_50` | the fixed rung is still refused as a selector |
| `test_51` | **the defect this group exists for**: an invented rung cannot be selected by inserting it |
| `test_52` | a forged rung refuses the WHOLE ladder, not just that rung |
| `test_53` | two swapped rungs are refused |
| `test_54` | a duplicated rung is refused |
| `test_55` | a missing fixed rung is refused |
| `test_56` | a label differing only in case or whitespace is refused |
| `test_57` | a non-finite value at the selected rung is refused, at NaN and both infinities |
| `test_58` | a non-finite value at a different rung refuses the whole ladder, and nothing is clamped |
| `test_59` | a short label carrier refuses without a subscript error |
| `test_60` | a short value carrier refuses |
| `test_61` | either carrier longer refuses; an empty carrier refuses too |
| `test_62` | the whole boundary performs **zero** sorts and **zero** quantile calculations — by call counters and by source |
| `test_63` | the limit of the claim: an edited finite value is NOT detected, and no seal was invented |

Every refusal in Group J goes through `_refused`, which pre-seeds the caller's
result Double with a sentinel and asserts it is **byte-unchanged** — so
requirement 14, "selected output unchanged on every malformed-ladder refusal",
is checked on every single malformed case rather than once.

`test_47` asserts the transcription read every procedure in the module except
the single named borrowed one (§2.1), and that the borrowed one is genuinely
absent from the compiled text — so a statistic still cannot hide in a body no
test compiles.

### 5.2 Mutation controls

A conformance test that cannot fail proves nothing. Each control writes a
damaged copy of `modSimStats.bas` into a temporary directory, points the
conformance module at it, reruns the **whole** Step-9 battery under a per-test
time budget, and requires a **named** detector among the refusers. Nothing is
written to the repository.

The 32 mutations the first authorisation lists are covered by 34 controls: the
extra two are `test_17a` (a correct quadratic insertion sort) and `test_31a` (a
`Static` local), each of which isolates a distinct failure the neighbouring
control does not reach. **Every one of them is kept unchanged**; no numerical
detector was deleted or weakened.

The hardening round adds **nine** more, covering the five required mutations and
four neighbours:

| Control | Mutation | Named detector |
|---|---|---|
| `test_33` | the ladder is not validated before it is searched — the `f44af49` boundary exactly | `test_51` |
| `test_34` | the owner-label authority is removed | `test_51` |
| `test_35` | labels are checked for **membership** and not for **position** | `test_53` |
| `test_36` | a case-insensitive comparison is substituted | `test_56` |
| `test_37` | the ladder-value finiteness check is removed | `test_57` |
| `test_38` | the physical carrier-length checks are removed | `test_59` |
| `test_38a` | only the label carrier's length is checked | `test_60` |
| `test_38b` | the scoped bounds handler becomes a broad suppression | `test_48` |
| `test_38c` | validation runs but no longer **gates** | `test_51` |

Two were rebuilt during this step because the first draft was **vacuous**:

* `test_17a` originally inserted a dead scanning loop alongside the merge sort.
  It changed no number and matched none of the source detectors, so it survived
  the battery. It is now a genuine substitution — a **correct, stable**
  insertion sort that orders every shape properly — so only the source shape
  (`test_13`) and the 100,000-value practicality test (`test_15`, through the
  harness time budget) can tell it from the accepted merge. The damaged source
  was checked to transcribe and sort correctly first, so the control is not
  passing on a transcription error.
* The per-test budget was cut from 300 s to **60 s**. The whole accepted battery
  runs in **1.37 s**, the 100,000-value test included at 0.57 s, so 60 s is a
  hundredfold margin over the slowest accepted test and only a genuinely
  non-terminating mutation trips it.

### 5.3 Existing inventory tests

The new registry member invalidated twelve tests whose assumption was the prior
exact inventory. **None was deleted, skipped or weakened**; each consumes a
**named** Phase-6 inventory — `PHASE6_VBA_MODULES`, `PHASE6_HANDWRITTEN`,
`_PHASE6_MANIFEST_MODULES`, `_PHASE6_HANDWRITTEN`, `PHASE6_INVENTORY`,
`PHASE6_MODULES` — or an explicit registry tail, so the addition stays a visible
edit. **No numeric module-count authority was introduced anywhere.**

`modSimStats` moved from the "banned" list to the "authorised" list in the five
tests that policed its absence. `modSimFingerprint` and `modSimReport` remain
banned in every one of them.

**The hardening round touched no test outside the two Step-9 files.** Three
expectations inside them legitimately changed and are named here so the edits
stay visible: `test_03`'s exact private-procedure set gained the two new private
helpers; `test_47` became the stronger compiled-from-source assertion described
above; and `test_48`'s `On Error` scan reads **comment-stripped executable
code** rather than raw text, because the module states in prose which construct
it refuses to contain — the same narrowing, for the same reason, that Step 8
settled for the `Cheng` token. A real `On Error Resume Next` is executable code
and is still caught.

---

## 6. Verification

### 6.1 Python suite

```
2736 passed, 0 failed          (769.19s)
2736 collected
```

| Count | What |
|---|---|
| 66 | Step-9 conformance tests |
| 44 | Step-9 tests — a baseline and 43 mutation controls |
| 54 + 42 | Step-8 engine conformance and controls — **still green, unmodified** |
| 74 + 51 | Step-7 sampler conformance and controls — **still green, unmodified** |
| 51 + 32 | Step-6 RNG conformance and controls — **still green, unmodified** |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2626** to **2711** in the first round (+85), and from
**2711** to **2736** in the hardening round (+25: 16 Group-J
conformance tests and 9 controls). **No test was deleted, skipped or
weakened** in either round.

### 6.2 Stage A

```
351 passed, 0 failed
Stage A build complete.
```

`build/PCCM_stageA.xlsx` — 35 734 bytes, `.xlsx`, no `vbaProject.bin`, no
`.bas`/`.bin` member and no `modSim` member. `modSimStats.bas` is source VBA for
Stage B; Step 9 does not embed it. **No Windows or Excel runtime ran.**

### 6.3 Module registry after Step 9

```
 1. modConstants        generated       11. modCalcAnalytical
 2. modWorkbook                         12. modCalcFingerprint
 3. modAppState                         13. modCalcResolve
 4. modTimeline                         14. modCalcCheck
 5. modDrivers                          15. modCalcReport
 6. modProfiling                        16. modSimContract   generated
 7. modInflation                        17. modSimRng
 8. modStructuralCheck                  18. modSimSample
 9. modCalcContract     generated       19. modSimEngine
10. modCalcFactors                      20. modSimStats               <- new
```

No count is hardcoded anywhere; P5-M and P5-D8 remain manifest-driven.

### 6.4 Artefact hashes

| Artefact | SHA-256 | Status |
|---|---|---|
| `src/vba/modSimStats.bas` | `98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0` | **hardened** (was `3e5ed2ca…c330b42` at `f44af49`) |
| `build/stage_b_manifest.json` | `0c413d93a0f2d002319584e4d59ce6c36dc612cb4115afcc898f7b8801720053` | **unchanged this round** (registry-only movement was the first round's) |
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` | **byte-identical** |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` | **byte-identical** |
| `src/vba/modSimEngine.bas` | `f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab` | **byte-identical** |
| `build/vba/modSimContract.bas` | `c7e7a78406345f98a3c2d0b90d63759b765a321aee99483fadd0f411f10c61be` | unchanged |
| `build/phase6_cases.json` | `5551606f7a0add5f980601b0a2cdd246130bd1e78678fd439bd5276cd36ec32c` | unchanged |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | byte-identical |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | byte-identical |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | byte-identical |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | byte-identical |

`git diff 39415f3 -- src/vba/modSimRng.bas src/vba/modSimSample.bas
src/vba/modSimEngine.bas src/vba/modCalcFactors.bas src/vba/modCalcAnalytical.bas
spec/sim_contract.yaml spec/input_contract.yaml spec/workbook.yaml
spec/calc_contract.yaml spec/driver_contract.yaml evidence/ builder/ bootstrap/`
is **empty**, and `git diff f44af49 -- spec/ builder/ tests/phase6_vba_transcribe.py`
is **empty** as well: the hardening round changed the module, its two test files
and this record, and nothing else.

### 6.4.1 The numerical core did not move

Compared procedure by procedure against `f44af49`, these are **byte-identical**:

```
SimStatsMean                     SimStatsSortAscending
SimStatsSampleStandardDeviation  SimStatsSortedCopy
SimStatsQuantileType7            SimStatsContingency
SimStatsQuantileSorted           SimStatsDescribe
SimStatsUnitScale                SimStatsConstantValue
SimStatsUsableSequence           SimStatsUsableProbability
SimStatsLadderLabel              SimStatsProbabilityOf
```

`SimStatsSelectedQuantile` gained the validator call and nothing else;
`SimStatsValidateLadder` and `SimStatsLadderExtent` are new. No mean, sample
deviation, minimum, maximum, quantile, ladder label, contingency value or sort
count moved, and `test_49` checks the selected value against the stored rung
bit-for-bit.

### 6.5 The manifest movement is the registry and nothing else

The hardening round moves the manifest **not at all** — no registry entry and no
D6-11 authority changes, and the rebuilt hash is identical to `f44af49`'s. For
the first round, rebuilt from the Step-8 contract and compared leaf by leaf:

```
keys added      3     (.vba.modules[19].name / .generated / .responsibility)
keys removed    0
values changed  0
```

and the structured forbidden-rule projection is **bit-identical**:

```json
{"construct": "MRG32k3a",      "allowed_in": ["modSimRng"]},
{"construct": "RunSimulation", "allowed_in": []},
{"construct": "Percentile",    "allowed_in": []}
```

Step 9 required no new scoped exception and took none.

---

## 7. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed, not touched.

**Two stale descriptive strings** in
`bootstrap/windows/phase4_functional_test.ps1` and
`docs/phase5_gate_b_harness.md`, both still saying "back to 15". Neither is an
executable gate. All three must be settled before the next real Gate-B/Windows
execution.

**Carried to Step 10 and beyond:** `SimStatsMeasure` is derived output and no
procedure reads it back, so it needs no structural validator today. The first
module that accepts a `SimStatsMeasure` as an INPUT inherits the Step-7
obligation — a Boolean `Described = True` is not validation authority.

**Carried explicitly to Step 11 (`modSimReport`), not implemented now:**

1. **Ladder provenance and non-mutation.** Step 9 proves a supplied ladder has
   the owner's labels, the right shape and finite values. It cannot prove a
   finite value was not edited after `SimStatsDescribe` produced it (§3.4).
   `modSimReport` must prove — and source-test — that the ladders it uses
   originate from `SimStatsDescribe` and are not mutated between description and
   selection.
2. **Nominal and PV commit together.** The selected values and contingencies for
   nominal and PV must be staged locally and committed as one, so a later PV
   refusal cannot publish a nominal-only reporting result.

**Deferred to Gate B, in addition to the standing debt:** the raising arm of
`SimStatsLadderExtent` — a genuinely never-sized VBA array — has no Linux
execution proof (§2.1). Its source shape is pinned by `test_48`; its runtime
behaviour is Windows work.

---

## 8. Step-9 acceptance gate — self-check

| Gate condition | Status |
|---|---|
| `modSimStats` exists | yes, 17 procedures, 6 public, 1 `Public Type` |
| entirely worksheet-independent | `test_05` |
| no module-level or `Static` state | `test_06` |
| the caller's array is never reordered | `test_10`, `test_45` |
| exactly one sort for the whole ladder | `test_11`, `test_12` |
| a bottom-up stable merge, not quadratic | `test_13`, `test_14`, `test_15` |
| moments taken in the caller's order | `test_20` |
| the scale is the largest power of two not exceeding the sample | `test_19` |
| the constant-sample invariant is exact and comes first | `test_17`, `test_18` |
| the deviation divides by `n - 1` | `test_22`, `test_23` |
| no sum of squares, no unguarded original-scale residual | `test_24` |
| unrepresentable dispersion refuses rather than reporting zero | `test_25` |
| a non-finite observation is refused, not skipped | `test_26` |
| an empty sequence refuses before any bound is read | `test_27` |
| type 7 is `h = (n - 1) * p` and matches the authority everywhere | `test_28` |
| integral `h` and equal brackets are exact | `test_29`, `test_30` |
| the convex interpolation form, not the difference form | `test_31` |
| the probability domain is closed and refused outside | `test_32` |
| the ladder is the projection, in the projected order | `test_33`, `test_34` |
| the selected level is a lookup in that same ladder | `test_35`, `test_49` |
| the fixed rung is reported and not selectable | `test_36`, `test_50` |
| an unknown confidence level is refused | `test_37` |
| a caller cannot invent a rung by inserting it into a supplied ladder | `test_51`, `test_52` |
| the supplied labels equal the owner projection, in the accepted order | `test_53`–`test_56` |
| every supplied ladder value is a usable finite Double, unclamped | `test_57`, `test_58` |
| malformed carrier lengths refuse rather than raising a subscript error | `test_48`, `test_59`–`test_61` |
| valid selectable lookups are unchanged | `test_49` |
| Selected CL triggers no sort, recomputation, RNG or simulation | `test_62` |
| the structural claim is not overstated, and no seal was invented | `test_63` |
| no numerical statistic moved | §6.4.1 |
| Step-11 provenance / no-mutation carry-forward recorded | §7 |
| contingency is `SafeSubtract(selected, base A)` | `test_38`, `test_41` |
| a negative contingency is preserved, never clamped | `test_39`, `test_40` |
| transactional output, no partial publication | `test_42`–`test_45` |
| the corpus rows are honoured under their own policies | `test_46`, `46a`, `46b`, `46c` |
| every procedure is transcribed and tested | `test_47` |
| `Quantile` naming discipline, no new scoped exception | `test_07`, §4.1, §6.5 |
| no fingerprint, digest, endpoint or publication | `test_08` |
| D6-11 scope unchanged | `test_02`, §6.5 |
| `modSimRng`, `modSimSample`, `modSimEngine` byte-identical | §6.4 |
| no test deleted, skipped or weakened | §5.3, §6.1 |
| no Step 10 exists | no `modSimFingerprint`, `modSimReport`, `PCCM_RunSimulation` |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B debt | **OPEN**, §7 |
