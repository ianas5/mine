# PCCM Phase 6 — Step 12: Gate-A source closure

Step 12 adds **no production feature**. It is the final Linux review: the
repo-wide integration audit of the Phase-6 chain, the closure of every carried
SOURCE/HARNESS debt, and the frozen baseline Step 13 will execute.

**No Windows. No Excel. No simulation executed.** Nothing in this record may be
read as a runtime result.

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Step-10A request-fingerprint grammar | `34a7c46` |
| Accepted Step-10 modSimFingerprint | `09a56b2` |
| Accepted Step-11A publication authority | `5e7e7b6` |
| Accepted Step-11 modSimReport | `955976f` |
| Accepted Review Round 4A ledger correction | `93ccdcb` |
| Step-12 Gate-A closure | this commit |

---

## 2. What Step 12 closed

### 2.1 GATE-B TEMP-DIRECTORY LEAK — **CLOSED**

Open since Phase-6 Step 4. `tests/test_phase5_gate_b_harness_source.py::_emitted`
called `tempfile.mkdtemp(prefix="pccm-gateb-")` and never removed the tree. The
helper is called by more than fifty tests, so a single suite run leaked more
than fifty directories; repeated runs accumulated tens of thousands and
exhausted the writable filesystem.

It now runs inside `tempfile.TemporaryDirectory(prefix="pccm-gateb-")`, which
removes the tree on the exception path as well as the normal one. **Every
artefact is read into memory before the directory closes**, so no caller is
handed a path that outlives it — and the `"dir"` key, which was searched for and
confirmed to have **no consumer at all**, is gone. Returning it would have been
exactly that kind of dangling handle.

- `test_224` drives the real helper three times and requires the `pccm-gateb-*`
  set to be unchanged afterwards, while asserting the returned data is still
  usable after cleanup.
- `test_225` pins the source shape: the context manager is there, `mkdtemp` is
  not, and every read happens inside the context.
- `test_nc_110` is the mutation control. It **executes** a helper with the
  pre-Step-12 shape, confirms a directory really was leaked, requires the
  detector to see it, and then removes only the one directory it created.

Measured on this tree: `pccm-gateb-*` count before a full suite run **5006**,
after **5006**, delta **0**.

Nothing deletes pre-existing directories owned by another process. The 5006
already on disk are the historical leak and are left alone.

### 2.2 STALE `"15"` DESCRIPTIONS — **CLOSED**

The executable P5-M and P5-D8 checks were **already** manifest-driven and were
not touched. Only the descriptions were stale:

| Where | Was | Now |
|---|---|---|
| `phase4_functional_test.ps1` banner | "The persisted project: 15 modules BY NAME" | "The persisted project matches the manifest module set BY NAME" |
| `phase4_functional_test.ps1` banner | "the inventory back to 15" | "the inventory returned to the manifest module set" |
| `phase5_gate_b_harness.md` scenario table | "15 modules **by name**" | "The manifest module set **by name**" |
| `phase5_gate_b_harness.md` scenario table | "inventory back to 15" | "inventory returned to the manifest module set" |

- `test_226` refuses any `\d+ modules?` in the active banner or in the active
  scenario table, and requires both to describe themselves in manifest terms.
- `test_227` proves the executable logic was manifest-owned already and that no
  numeric module count exists anywhere in the scenarios file.
- `test_nc_111` plants three forms of the stale wording and requires the
  detector to see each.

**Historical Run tables keep their historical numbers.** The Run-record row that
says "fifteen modules present … under the evidence model P5-M then had" was true
when it was written and is left exactly as it is.

### 2.3 STALE RESULTS PRESENTATION TEXT — **CLOSED**

| | Was | Now |
|---|---|---|
| Subtitle | "Simulation output - no statistics are implemented in Phase 1" | "Simulation results and statistical summary" |
| Top note | "Not implemented yet. No percentiles, moments or cash flows are implemented yet." | "Phase 6 publishes run identity, summary statistics and selected-confidence reporting. Annual cash flow and reconciliation remain deferred." |

**Presentation text only.** No layout, no formula, no row moved — `test_95` pins
the accepted Step-11A geometry (label column `B`, nominal `D`, PV `F`, sections
at rows 8 and 26, run stamp 10–24, summary header 28 and body 29–44, selected
rows 46/47/48, deferred rows 51 and 54, first iteration row 34) and the
structural shape of all five selected formulas.

- `test_92` refuses the retired phrases in the manifest.
- `test_93` reads the **built** Results sheet and requires the true wording.
- `test_94` refuses any claim that sensitivity, correlation, a dashboard, an
  implemented reconciliation or an implemented annual cash flow exists, and
  requires both deferred sections to keep their deferred notes.
- `test_nc_96` / `test_nc_97` are the mutation controls for both detectors.

---

## 3. The integration review

`tests/test_phase6_integration_source.py` — **27 tests**
`tests/test_phase6_integration_source_validation.py` — **37 controls + baseline**

Nothing there duplicates a per-step suite. It proves the **seams**: that the
value one module produces is the value the next module consumes, and that no
second construction of a shared number exists anywhere in the chain.

| § | Seam | How it is proved |
|---|---|---|
| A | The Phase-5 bridge is the only door | `CalcPrepareSimulationInputs` is declared once, in `modCalcReport`, and is the only name `modSimReport` reaches there |
| A | It reuses and gates | `PrepareCurrentCalculation` < `DeriveStatus` < `CALC_STATUS_CURRENT`, in that order |
| A | It projects, never rebuilds | eight rebuild entry points asserted absent; every projected field read off the one prepared package |
| A | It writes nothing | no `.Value2 =`, no `PCCM_Calculate`, no `PCCM_CalculationStatus`, and it is not on the automation surface |
| B–D | Kernel ownership | the reporter's kernel call set is pinned to exactly seven names; engine, statistics and hash internals asserted absent; no `/`, `*` or `^` anywhere in executable code |
| B–D | Contingency | filled only from `SimStatsContingency`'s out-parameter; every carrier assignment must be exactly `= value` |
| E | Request identity | the **first argument** to `SimFpBuildRequestFingerprint` is `package.AnalyticalFingerprint` at both call sites, VBA continuations folded; the field is never reassigned; no stored snapshot, seed, nonce, run id or selector reaches the request path |
| F | Retained-array identity | one engine call fills `TotalNominal`/`TotalPv`; statistics, digest and publication all name those same two fields; no assignment or `ReDim` replaces either |
| F | No reconstruction | no sort/swap/reverse token anywhere; the engine runs exactly once |
| G | Quantile provenance | Describe < SameLadder < Contingencies < fingerprint < digest; no ladder element assigned, qualified or not |
| H | Dual bank | no candidate stage names the active bank; the selector row has exactly one namer, which reads it; exactly one procedure writes `D22:D30`, and it writes twice (commit + restore) and no more |
| I | Results boundary | no VBA module names `"Results"`, `shResults` or `Results!`; the reporter resolves only `SIM_DATA_SHEET`; the contract's five presentation flags asserted |
| J | Selected CL | absent from every module with strings intact, and from all six run/status procedures; both contract participation flags `false` |
| K | Attempt orthogonality | the derivation names no `SIM_ATTEMPT_`; blank selector is an absence, not a fourth state; the attempt writer derives |
| L | Public surfaces | Phase-6 exactly seven; Phase-4 **derived from the contract**, not restated; Phase-5 exactly its six plus the one internal bridge; no Phase-6 button |

### D6-11, repo-wide (§4 of the authorisation)

```
MRG32k3a      allowed_in = [modSimRng]     present there, absent everywhere else
RunSimulation allowed_in = [modSimReport]  present there, absent everywhere else
Percentile    allowed_in = []              absent from every executable module
Rnd(  Randomize  NPV  Worksheet_Change  Workbook_SheetChange  FinalReleaseComObject
                                           global, unscoped, absent everywhere
```

Each grant is proved **non-vacuous** (the construct really is in its owner) and
**exclusive** (`forbidden_in(name) == (name != owner)` for every declared
module). No rule is flattened, no owner list has more than one entry, and no
wildcard exists. **No module count is hardcoded**: `test_22` requires the
emitted manifest, the structure contract and the files on disk to agree as
sets, and requires every flattened entry to have a structured rule behind it.

### The controls (§14)

All fifteen required mutations are covered and non-vacuous:

| # | Mutation | Control |
|---|---|---|
| 1 | stored fingerprint instead of bridge output | integration `test_01`–`test_03` → `test_07` |
| 2 | arrays differ between stats, digest, publication | `test_04`, `test_05` → `test_09` |
| 3 | quantile carrier mutated after Describe | `test_07` → `test_11` |
| 4 | Selected CL enters the run | `test_09` → `test_15` |
| 5 | Results written by VBA | `test_10` → `test_14` |
| 6 | active bank touched during candidate publication | `test_11` → `test_12` |
| 7 | RunSimulation in a second module | `test_17`, `test_18`, `test_21` → `test_19` |
| 8 | MRG outside modSimRng | `test_19`, `test_22` → `test_19` |
| 9 | executable Percentile | `test_20`, `test_23` → `test_20` |
| 10 | bridge calls PCCM_Calculate | `test_12` → `test_04` |
| 11 | P5-LDG before P5-FIN | Gate-B `test_155a` (Round 4A, unchanged) |
| 12 | exact-type setter widens | Gate-B `test_nc_98` (Round 4A, unchanged) |
| 13 | temp-directory helper leaks | Gate-B `test_nc_110` |
| 14 | fixed module count in active wording | Gate-B `test_nc_111` |
| 15 | stale Results text returns | Stage-A `test_nc_96` |

Plus the frozen-source controls (`test_29`–`test_35`), which move an accepted
module or a generated artefact by one byte and require the hash gate to refuse.

---

## 4. Round 4A is not regressed

`bootstrap/windows/phase5_gate_b_scenarios.ps1` has **no diff**. The accepted
exact-type settlement is untouched: supported restore types remain `null`,
`System.String`, `System.Double`, `System.Boolean`; the numeric CLR aliases and
`DateTime` still refuse before any `Value2` assignment; there is no tolerance;
string comparison is case-sensitive; P5-FX still goes through the strict
comparator. The ledger order remains `P5-FIN` → `P5-LDG` → the FAIL-count
summary, with `test_154`, `test_155` and `test_155a` all green.

---

## 5. Stage-A determinism — a finding, stated plainly

`PCCM_BUILD_TIMESTAMP` pins everything **PCCM** writes. It does **not** make the
workbook byte-identical across builds, and the reason is not PCCM code:

Two consecutive builds with `PCCM_BUILD_TIMESTAMP=2020-01-01T00:00:00`:

```
47 zip members, identical member list
46 members byte-identical
 1 member differs: docProps/core.xml
   the ONLY difference inside it is <dcterms:modified>, which openpyxl
   stamps from the wall clock at save time
zip local-header date_time also varies, for the same reason
```

With `<dcterms:modified>` normalised away, the content digest of both builds is
identical:

```
2bb21393b9724f455b0e1cd047958a341d021da6ebc9659c1cae72080e660897
```

So the **semantic** content is reproducible; the residual bytes are third-party
OOXML package metadata. Closing the remaining gap would require changing
`builder/**`, which Step 12 is not authorised to do and which the environment
override cannot achieve. **This is recorded, not deferred as a runtime item** —
it is a build-tooling property, not a simulation property, and it changes no
production semantics. The workbook SHA therefore remains **build-instance
evidence only**, exactly as Step 11 §41 stated.

---

## 6. FINAL GATE-A REVIEW MATRIX

| Area | Authority | Source owner | Gate-A evidence | Gate-B required? |
|---|---|---|---|---|
| Simulation controls | `input_contract.yaml`, `sim_contract.yaml` | `modSimReport.ResolveIterations` / `ResolveSeed` | **PASS / SOURCE-CLOSED** — only the two named controls read; domain proved before narrowing; blank seed is AUTO with no sentinel | **YES** — strict Excel type reads |
| Phase-5 bridge | Step-11A | `modCalcReport.CalcPrepareSimulationInputs` | **PASS / SOURCE-CLOSED** — one door, reuses, gates on CURRENT, projects only, writes nothing | **YES** — real Excel execution against a CURRENT workbook |
| RNG | Step 6 | `modSimRng` | **PASS / SOURCE-CLOSED** — transcribed and proved; D6-11 owner | **YES** — real VBA arithmetic |
| Samplers | Step 7 | `modSimSample` | **PASS / SOURCE-CLOSED** — transcribed and proved; Cheng lives only here | **YES** — real VBA arithmetic |
| Engine | Step 8 | `modSimEngine` | **PASS / SOURCE-CLOSED** — one call, retained carriers filled directly | **YES** — real integration at scale |
| Statistics | Step 9 | `modSimStats` | **PASS / SOURCE-CLOSED** — ladder validated structurally; no statistic outside the owner | **YES** — `SimStatsLadderExtent` raising arm |
| Request fingerprint | Step 10A grammar, Step 10 | `modSimFingerprint` | **PASS / SOURCE-CLOSED** — prefix is the bridge's CURRENT fingerprint at both call sites | **YES** — `CalcFpContinueDigest`, `AscW` vs `Asc` |
| Result digest | Step 10 | `modSimFingerprint` | **PASS / SOURCE-CLOSED** — built once, over the published arrays | **YES** — `SimFpRetainedExtent` raising arm |
| AUTO nonce | Step 11A | `modSimReport.AllocateAutoNonce` | **PASS / SOURCE-CLOSED** — persisted and read back before sampling; never rolled back | **YES** — persistence/read-back on real Excel |
| Run id | Step 11A | `modSimReport` | **PASS / SOURCE-CLOSED** — headroom proved before a candidate is computed; written only in the final commit | **YES** — real counter round-trip |
| Dual-bank publication | Step 11A | `modSimReport.PublishCandidate` / `FinalCommit` | **PASS / SOURCE-CLOSED** — inactive bank only; verified; one `D22:D30` write with the bank last | **YES** — chunked write/read-back, verification, commit atomicity, restoration |
| Simulation status | Step 11A | `modSimReport.DeriveSimStatus` | **PASS / SOURCE-CLOSED** — three states plus absence; attempt-orthogonal | **YES** — derivation against a real published bank |
| Attempt axis | Step 11A | `modSimReport.RecordRefusal` / `RecordFailure` | **PASS / SOURCE-CLOSED** — refusal and failure are different records; three refuse sites, two fail sites | **YES** — real refusal and failure paths |
| Results presentation | Step 11A | Stage-A builder over `_SimData` | **PASS / SOURCE-CLOSED** — never written by VBA; geometry pinned; wording now true | **YES** — presentation against a real published bank |
| Public API | Step 11 | `modSimReport` | **PASS / SOURCE-CLOSED** — exactly seven, no button, no invented accessor | **YES** — accessor execution |
| D6-11 | Step 1 precondition, Steps 6 and 11 grants | `structure_contract.yaml` + manifest | **PASS / SOURCE-CLOSED** — both grants non-vacuous and exclusive; `Percentile` ownerless | **YES** — persisted-project enforcement of scoped `RunSimulation` |
| Application-state envelope | Phase 5 accepted | `modAppState` via `PCCM_RunSimulation` | **PASS / SOURCE-CLOSED** — envelope installed before the first fallible operation; cleanup attempted at most once; a post-commit cleanup problem does not unpublish | **YES** — `FinishOperation`, `Announce`, real failure paths |

No row reads "looks good".

---

## 7. Open items — runtime-only, and why

Every one of these can be closed **only** by real VBA/Excel. None is a
source-cleanable issue relabelled to defer it; each names the specific machine
behaviour no Linux model can supply.

| Item | Why only a runtime can close it |
|---|---|
| `SimStatsLadderExtent` raising arm | reads a bound of an unproven carrier under a scoped `On Error`; the transcriber models no error handling |
| `SimFpRetainedExtent` raising arm | same shape, same reason |
| `CalcFpContinueDigest` in real VBA | the transcribed arithmetic is proved; VBA's own execution of it is not |
| `AscW` vs `Asc` | the distinction is a host-encoding behaviour; Linux cannot exhibit it |
| `CalcPrepareSimulationInputs` real execution | needs a real Phase-5 CURRENT workbook |
| Strict simulation-control Excel type reads | `IsWholeInRange` against real Excel cell variants |
| AUTO nonce persistence / read-back | a real worksheet round-trip |
| Real engine / stats / fingerprint integration | the composed numerical chain on the real interpreter |
| Inactive-bank chunk write / read-back | COM bulk-write behaviour at scale |
| Candidate-bank verification | real `Value2` round-trip fidelity |
| Final `D22:D30` commit | atomicity as Excel actually performs it |
| Final-commit restoration | the failure path, which cannot be induced here |
| Public Phase-6 accessor execution | real invocation |
| Application-state cleanup | `FinishOperation` and `Announce` on a real host |
| Scoped `RunSimulation` persisted-project enforcement | the sweep runs against the persisted project |
| Results presentation against a published bank | real formula evaluation |
| 100,000-iteration performance | a real machine |

These are the **Step-13 evidence matrix**. Step 12 documents them; it does not
claim any of them executed, and it implements no Step-13 scenario.

---

## 8. What changed

```
MOD  tests/test_phase5_gate_b_harness_source.py   temp-dir lifecycle + 5 new tests
MOD  bootstrap/windows/phase4_functional_test.ps1 stale "15" wording only
MOD  docs/phase5_gate_b_harness.md                stale "15" wording only
MOD  spec/workbook.yaml                           two presentation strings only
MOD  tests/test_phase6_stage_a.py                 Results-text and geometry tests
NEW  tests/test_phase6_integration_source.py
NEW  tests/test_phase6_integration_source_validation.py
NEW  docs/phase6_step12.md
```

**No change to** `src/vba/**`, `builder/**`, `evidence/**`,
`spec/sim_contract.yaml`, `spec/input_contract.yaml`, `spec/calc_contract.yaml`,
`spec/driver_contract.yaml`, `spec/structure_contract.yaml`, or
`bootstrap/windows/phase5_gate_b_scenarios.ps1`. No production semantics moved.
No test was deleted, skipped or weakened.

### Accepted hashes, proved unchanged

```
modSimRng.bas          3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3
modSimSample.bas       5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877
modSimEngine.bas       f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab
modSimStats.bas        98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0
modSimFingerprint.bas  9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda
modSimReport.bas       a0b9a738b8f7346efd7f5964c311861d975075786072e2ec7b7c7773afd0c363
modCalcFingerprint.bas 2efbb30c6f915c04b9c07adec07e25e11f4b5bd2b98e3efa818631dc510ce847
modCalcReport.bas      8252b935b256b1abad9b26ca6b1d90c92c5e0d7566906308b191cd03dd6a71b3
  accepted prefix      5d4568aef01037fd2999915da87a550d02033441b8c26c80f9386d4fcf8b087f
modSimContract.bas     1d949be659d0afc3e18501a34b7d372bab3df575fc1a981cfd60dcf1f293a753
phase6_cases.json      98f835375f5b8f548172c21ae6102b50fef7e6a001e196ece0741c987d78b6d1
stage_b_manifest.json  51335e3339ab480b28760a2c58fbe83e72ce3d1be54554766857598db7272049
```

The workbook binary moved, as expected: two presentation strings changed and the
package metadata is timestamped.
