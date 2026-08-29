# PCCM Phase 6 — Step 13: the Windows/Excel Gate-B runtime harness

**Status: HARNESS SOURCE, submitted for review. NOTHING HAS BEEN RUN.**

No Excel COM session has been started, no `.xlsm` has been driven, and not one
Phase-6 procedure has executed. Everything below is a statement about SOURCE and
about the artefacts the Stage-A builder emits.

```
static / source evidence   !=   Windows / Excel runtime evidence
```

That line is the whole point of Step 13, and this document keeps it in front of
the reader rather than at the end.

---

## 1. What Step 13 exists to prove, and what it cannot

The Linux suite proves the pure kernels exactly, the transaction order
structurally, and the publication state machine as data. What it cannot do is
start Excel. Step 13 asks the questions only a real COM session can answer:

* does the accepted workbook build, open and **compile**;
* is the Phase-6 public surface callable;
* does a FIXED run publish a result, and does a repeat publish the same digest;
* **does real Excel produce the number the accepted Python oracle produces**;
* does an AUTO run consume, persist and clear correctly, and never replay;
* do the three accepted failpoints leave the workbook in the states the
  transaction says they should;
* does the durable `F21` protocol actually block and actually recover;
* and does one attempt-result string demonstrably fail to encode the physical
  allocation classification — the Step-12 correction, in persisted evidence.

---

## 2. Architecture: an extension, not a second harness

`phase6_gate_b_scenarios.ps1` is **dot-sourced into** `phase4_functional_test.ps1`,
exactly as the Phase-5 block is. It runs inside that script's one COM lifecycle,
against the one Excel instance it owns, the one workbook it opened and the one
Stage-B bootstrap it ran, and reports through the same `Add-Result`. It creates
no Excel process, no release ledger, no bootstrap invocation and no shutdown of
its own; `test_03` refuses each by name.

The Phase-4 matrix and the Phase-5 Gate-B block are **prerequisites**. `P6-PRE`
reads the results they actually produced and requires every one to be a PASS
before a single stateful Phase-6 scenario runs.

### The one change to the accepted driver

`phase4_functional_test.ps1` gains a dot-source, a preflight call, two artefact
loads, one commit capture and one scenario call. `test_02` proves no accepted
line was removed except the required-artefact list, which was extended with the
two Phase-6 paths, and `test_01` proves `phase5_gate_b_scenarios.ps1`,
`phase5_gate_b_diagnostics.bas`, `com_lifecycle.ps1` and `build_stage_b.ps1` are
line-for-line what the production baseline holds.

---

## 3. The two new build artefacts

### 3.1 `build/phase6_gate_b_inspection.json` — addresses only

The harness has to find `_SimData`. No existing build output projects it:
`stage_b_manifest.json` stops at sheets, modules, entry points, buttons,
registers and grids; `phase5_gate_b_inspection.json` is the Phase-5 projection;
`phase6_cases.json` is a value corpus in which `F21` does not appear at all. The
addresses exist in `build/vba/modSimContract.bas`, but that is VBA source, and
teaching PowerShell to parse VBA constants would put a **second reader of the
same layout authority** in the harness.

`builder/pccm_builder/sim_inspection.py` projects them through the accepted
loaders from `sim_contract.yaml`, `input_contract.yaml` and `workbook.yaml`,
with an explicit `schema_version`, a **positive key allowlist** at every level,
and identities only:

| block | carries |
|---|---|
| `sim_data.sheet` / `required_visibility` | `_SimData`, `veryHidden` |
| `sim_data.run_identity` | label/value/note columns, both bank columns, first/last row, **all 23 rows by meaning**, each row's axis (`snapshot` / `counter` / `attempt` / `derived` / `control`), and its label |
| `sim_data.pending_auto_nonce` | the F21 cell, column, row, label |
| `sim_data.iteration_records` | header row, first iteration row, footer rows, column keys, both banks' columns |
| `sim_data.summary_statistics` / `contingency_ladder` | label column, per-bank nominal/PV columns, first/last row, every row by key |
| `publication` | bank labels, the candidate-target selector map, the final-commit range and its nine fields |
| `controls` | `monte_carlo_iterations` and `random_seed`: defined name, sheet, cell, type |
| `command_surface` | the automation endpoint and the six read accessors |

**What it deliberately does NOT carry.** No bound, no tolerance, no label set.
`SIM_MIN_ITERATIONS`, the seed domain and the run-ID maximum are VALUES a
scenario compares against, so they belong to the expectation corpus; the
attempt-result and simulation-status vocabularies are model SEMANTICS — the
Phase-5 projection carried label lists in its first submission and independent
review removed them, and the same line is held here. The three failpoint stage
names are declared in **production VBA**, not in a contract, so they are not
projected either: the harness declares them once and `test_11` pins those
strings against `modSimNonce.bas` and `modSimReport.bas`.

Every emitted address is cross-checked on Linux against the generated
`modSimContract.bas` and `modConstants.bas` (`test_17`–`test_20`), and the sheet
and its visibility against the generated Stage-A workbook itself (`test_21`).

### 3.2 `build/phase6_gate_b_cases.json` — the oracle's expectations

Generated by `sim_cases.build_gate_b_cases` through the accepted
`prepare_simulation` / `run_simulation` / `result_digest`. No expected number is
written by hand, and none is written in PowerShell.

**The binding to the runtime fixture is by plan-case id.** Each parity case names
an EXISTING Phase-5 plan case; the Windows scenario drives that same case through
the same accepted `Set-Phase5Fixture` machinery `P5-AN` already uses. The two
implementations are therefore not describing "similar" fixtures — there is one
model definition, in `calc_cases.CASES`.

#### Why four cases and not one — reported, not worked around

The authorisation asks for one existing fixture that covers the required
mechanisms. **There is none.** No single Gate-B plan case reaches all four
sampling paths. Rather than invent a fifth fixture, four existing ones are used:

| plan case | mechanism | why |
|---|---|---|
| 1 | Triangular | and it is the golden analytical reference (below) |
| 6 | Beta-PERT | the only fixture exercising the prepared Cheng shape |
| 7 | Uniform | min ≠ max, so genuinely stochastic |
| 8 | Bernoulli occurrence + Triangular severity | the only family that reaches the risk occurrence primitive |

Plan case 30 — three Uniform drivers with `min == max` — is deliberately **not**
used: it is degenerate and would pass whatever the RNG did. `test_26` refuses any
parity fixture with a `min == max` driver.

#### The golden case is stronger than the other three, and the difference is stated

Plan case 1's model **is** the accepted fingerprint reference vector, so its
analytical digest comes from `reference_stream` — an authority that is already
emitted and already pinned — and its request fingerprint follows from the
accepted composition. `test_27` proves the emitted value equals
`phase5_cases.json → fingerprint.reference.digest`.

For any other model the analytical canonical stream would have to be rebuilt
field by field in Python, which would be **a second implementation of the
fingerprint field layout**. It is not done. Those three cases emit no
fingerprint, and their analytical identity is established at runtime from the
accepted `phase5_cases.json` expectations that `P5-AN` already drives. `test_46`
proves the identity is established **before** any simulation value is compared,
and `test_27` proves exactly one case claims the derivable binding.

#### Iterations, seed and exactness

1000 iterations — asserted to be `input_contract.yaml`'s own business minimum,
never restated — and FIXED seed 12345, one of the accepted RNG vectors. The
comparison policy is **EXACT** and admits no tolerance: the canonical numeric
encoder normalises the host decimal separator before hashing, so digest equality
holds on any locale. `test_31` refuses a granted tolerance; `test_10` refuses one
in the harness.

---

## 4. The scenario matrix

| ID | Proves | Evidence source |
|---|---|---|
| `P6-PRE` | Phase-4 35/35 and the Phase-5 block are intact | recorded results |
| `P6-ART` | source commit, Stage-A hash, executed `.xlsm` hash, manifest and both artefact hashes | `Get-FileHash` on the driven copy |
| `P6-CMP` | the project that compiled contains the eight `modSim*` modules — **derived from `P5-CMP`** | no second compile |
| `P6-M` | the proved inventory is the manifest's 23-module set — **derived from `P5-M`** | no second inventory |
| `P6-API` | seven procedures declared; six accessors callable; the endpoint deferred to `P6-FX1` | persisted project + `Application.Run` |
| `P6-BTN` | no shape invokes any Phase-6 procedure | shape `OnAction` |
| `P6-INIT` | a workbook that has never simulated | `_SimData` cells |
| `P6-FX1` | the first real `PCCM_RunSimulation`: bank A published, run id 1, FIXED seed recorded | cells |
| `P6-ORA` | **Excel equals the oracle** — digest, seeds, versions and the full ladder, both measures, all four cases | cells vs corpus |
| `P6-DET` | the same inputs and seed twice produce the same digest, and it is the oracle's | cells |
| `P6-FIXED-INERT` | FIXED with `Phase6AfterNoncePersisted` armed still succeeds; counter and F21 untouched | cells |
| `P6-AU1` | AUTO consumes *m*, persists *m+1*, clears F21 | cells |
| `P6-AU2` | the retry takes the next nonce with a distinct effective seed — **not** "different digest" | cells |
| `P6-BANK` | selector moves to the candidate target; commit block agrees with the published bank; the other bank untouched | cells |
| `P6-ACC` | all six accessors agree with the published authority; the status accessor moves only its two derived rows | accessors + whole-block capture |
| `P6-RF1` | a prerequisite refusal spends nothing | cells |
| `P6-PRESERVE` | the prior publication survives a refusal, field for field | cells |
| `P6-FP1` | `Phase6AfterNoncePersisted`: counter advanced, F21 clear, seed and nonce recorded, result **REFUSED** | cells |
| `P6-FP2` | `Phase6CandidateBank`: **FAILED**, selector unmoved, prior publication authoritative | cells |
| `P6-FP3` | `Phase6FinalCommit`: **FAILED**, prior shared block restored, prior bank still active | cells |
| `P6-REC1` | crash-equivalent marker whose advance DID persist → cleared, not reissued | cells |
| `P6-REC2` | crash-equivalent marker whose advance never persisted → cleared, legitimately taken | cells |
| `P6-REC3` | inconsistent durable state → retained, REFUSED, and a **second** AUTO run still blocked | cells |
| `P6-REC4` | corrupted pending marker → REFUSED, retained | cells |
| `P6-REC5` | corrupted counter → REFUSED before any allocation | cells |
| `P6-RIDMAX` | run-ID exhaustion refuses before allocation, and the captured id is restored exactly | cells |
| `P6-AXIS` | one attempt-result string does not encode the allocation classification | the persisted facts above |
| `P6-LDG` / `P6-FIN` | one result per ID; nothing skipped | ledger |

### 4.1 `P6-BANK` claims nothing that PowerShell cannot see

`Application.Run` is **synchronous**. Nothing inspects the workbook while
`PCCM_RunSimulation` is executing, and neither failpoint suspends the
transaction for inspection: `Phase6FinalCommit` fires **after** the `D22:D30`
assignment and **before** production's own restore, both inside one call, so by
the time PowerShell regains control the prior block is already back.

`P6-BANK` is therefore an **aggregate** of three independently observable states
— its own before/after pair, `P6-FP2` and `P6-FP3`. The ordering statement, that
the selector assignment happens only inside `FinalCommit` after candidate
publication, remains **source** evidence proved by the Linux transaction-order
controls, and the scenario's own report says so. `test_40` refuses any mid-call
observation claim in the file.

### 4.2 `P6-AXIS` claims nothing about a private field

`NonceConsumed` is a Private Boolean inside `modSimReport`. PowerShell cannot
observe it. What `P6-AXIS` proves is the **persisted** consequence: several
unsuccessful attempts record the same `REFUSED` while the AUTO axis around them
differs — `P6-FP1` advanced the counter, `P6-RF1` did not, `P6-REC3` was
reconciling a prior marker, `P6-REC5` never selected an identity. That is direct
runtime evidence that Last Attempt Result alone does not encode the physical
allocation classification. The `NonceConsumed` projection stays source evidence,
and `test_41` refuses any file text that claims otherwise.

### 4.3 The recovery fixtures are not all the same kind of thing

`P6-REC1`/`REC2` are **crash-equivalent** write-ahead states. `P6-REC3` is an
**inconsistent durable-recovery** state. `P6-REC4` is a **corrupted pending
marker** and `P6-REC5` a **corrupted counter**. All five are honestly
constructible by writing cells on the disposable `%TEMP%` copy; none fakes a COM
failure.

Each writes the **fewest cells possible**: `REC1` chooses *m* = counter − 1 so
the existing counter already reads *m+1*; `REC2` chooses *m* = counter; `REC3`
chooses an *m* the existing counter already disagrees with. No second cell is
corrupted to manufacture a state.

### 4.4 Restoration is part of the scenario

Every fixture-writing scenario follows one policy: **capture the typed original
→ write → verify the write took → run → capture POST evidence BEFORE cleanup →
restore → verify the restoration is exact.** The restoration check goes into the
same checklist the scenario's PASS/FAIL is computed from, so a scenario whose
restoration cannot be verified **FAILS**. It is never a note. `test_38` and
`test_39` pin both halves, and mutations 37, 38 and 39 require refusal.

---

## 5. What remains static-only after Step 13

| Not runtime-provable | Why |
|---|---|
| `PERSISTENCE_INDETERMINATE`, and therefore the `AUTO_NONCE_INDETERMINATE` production path | needs the counter's post-write verification read to disagree with the write **inside one VBA call**. There is no injection point between them and no cell edit produces it. Fabricating it would prove only that the harness can fabricate. |
| `SharedReadRaised` / `ReadRaised` genuine COM raises | `P6-REC4`/`REC5` induce the *validation* refusal honestly; the raise itself cannot be induced without breaking Excel. |
| `ClearPending` failure after a known `CONSUMED` | the exact path the Step-12 defect lived on; not inducible. `P6-FP1` proves the adjacent clean-clear case. |
| iteration ceiling 1 048 543 | not runnable. |
| the private `NonceConsumed` projection | not observable from PowerShell (§4.2). |
| the selector-write ordering inside `FinalCommit` | not observable across a synchronous call (§4.1). |
| wording / source mutation controls | source-level by nature. |

**No longer static-only:** cross-implementation parity (`P6-ORA`, D2b) and
run-ID exhaustion (`P6-RIDMAX`).

---

## 6. The static controls

`tests/test_phase6_gate_b_harness_source.py` — **48** controls, in six groups:
the accepted harness is not rewritten; the harness restates no address, name or
expected value; the failpoint and procedure names are checked copies; the
projection agrees with the generated authority; the corpus is generated, bound
and exact; the matrix is complete and fail-closed.

`tests/test_phase6_gate_b_harness_source_validation.py` — **67** mutation
controls, each requiring a **named** detector: F21 moved by a row and by a
column, the final-commit range moved, both bank column pairs swapped, four
identity rows moved, ladder rows shifted, both control defined names changed, the
sheet renamed, visibility relaxed, required keys deleted, unapproved keys added,
a vocabulary smuggled into the projection, the surface altered — and, on the
PowerShell side, six kinds of hard-coded address, a pasted digest, an admitted
tolerance, an invented failpoint, a dropped accessor, a recorded SKIP, a removed
prerequisite gate, a defanged ledger, a dropped restoration, evidence captured
after cleanup, a mid-call claim, a private-field claim, digest inequality as a
rule, a re-executed compile, a second Excel instance, a weakened preflight, a
dropped scenario, a deleted disclaimer — and on the corpus side, altered
digests, a drifting iteration count, a broken plan-case binding, a lost ladder,
an admitted tolerance, a second golden claim, drifting bounds, a disturbed
vocabulary, a degenerate fixture, a hand edit, an emptied corpus, corrupted JSON
in either artefact, and a parity comparison that runs before its identity check.

---

## 7. Running it

One command set, one working directory — the repository root:

```
python pccm\builder\build_stage_a.py
powershell -ExecutionPolicy Bypass -File .\pccm\bootstrap\windows\build_stage_b.ps1
powershell -ExecutionPolicy Bypass -File .\pccm\bootstrap\windows\phase4_functional_test.ps1 `
  *> .\pccm\bootstrap\windows\phase6_gate_b_run1.log
```

The third command runs everything: the Phase-6 block is dot-sourced into the
Phase-4 harness and runs inside its single COM lifecycle, against the disposable
`%TEMP%` copy. There is no separate Phase-6 script to invoke and no second Excel
instance.

**Prerequisite.** Importing VBA requires *Trust access to the VBA project object
model*. The scripts report it and stop if it is missing. They do not enable it,
do not lower macro security, do not edit the registry and do not add a Trusted
Location — the same refusal that has held since the first readiness run.
