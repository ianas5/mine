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

The Phase-4 matrix and the Phase-5 Gate-B block are **prerequisites**, and the
lifecycle topology is the whole difficulty of saying so.

### 2.1 Where Phase 6 sits in the lifecycle, and what can exist there

```
Invoke-Phase5GateBScenarios -> Invoke-Phase6GateBScenarios ->
PCCM_AutomationEnd -> workbook close -> Excel quit -> COM release ->
Z -> Y -> Add-Phase4FinalCompletenessResult (P5-FIN) -> P5-LDG -> P6-LDG
```

`Y` and `Z` are **post-session** lifecycle assertions, recorded after Excel is
torn down. `P5-FIN` and `P5-LDG` are post-session for the same reason. So a
prerequisite gate running inside the live session cannot demand them: the first
submission demanded the full `Get-Phase4RequiredScenarioIds` set, which includes
Y and Z, and would therefore have failed `P6-PRE` — and every stateful scenario
with it — before the first simulation ran. That is the Gate-B Run-1 sequencing
defect under a Phase-6 name.

`P6-PRE` now uses the **accepted derived partition** the Phase-5 block already
established for exactly this case: `Get-Phase4PrerequisiteScenarioIds`, which is
the required set minus the deferred finalisation cases. It proves the partition
is real (disjoint, covering, no stray), proves the deferral is real (no deferred
case has already run), and applies the same treatment to Phase 5 — every
in-session `Get-Phase5ScenarioIds` scenario recorded once and passed, `P5-P4`,
`P5-CMP` and `P5-M` passed by name, the `P5-ALL`/`P5-XX` failure channels unused,
and `P5-FIN`/`P5-LDG` **not yet recorded**.

**The 35/35 demand is not weakened.** It is still made, later and by the
already-accepted `Add-Phase4FinalCompletenessResult` / `P5-FIN`, after Y and Z
exist. Nothing in Step 13 replaces it, and `test_50` proves the Phase-6 block
does not claim to.

### 2.2 "P5-LDG has not run yet" is necessary, and not sufficient

The accepted Phase-5 guard leaves an intermediate state that a scan of recorded
results cannot see:

```
P5-X recorded once, PASS, visible in $Results
a second P5-X attempt refused and recorded in Phase5LedgerViolations
no P5 result anywhere reads FAIL
P5-LDG deferred until after Phase 6, so no verdict exists yet
```

Every Phase-5 scenario reads PASS while Phase 5 is **already known** to have a
harness-integrity violation. A prerequisite that only counted results would run
the whole Step-13 matrix on top of it.

`P6-PRE` therefore reads `Get-Phase5LedgerViolations()` directly and requires it
empty. The authority is Phase 5's own, consumed rather than reimplemented, and
`P5-LDG` still emits its verdict at the accepted point in the lifecycle — the
Phase-6 block does not emit it early, and `test_54` refuses a version that does.

### 2.2 The one change to the accepted driver

`phase4_functional_test.ps1` gains a dot-source, a preflight call, two artefact
loads, one harness-commit capture, one scenario call, one ledger-verdict call and
a summary that names Phase 6. `test_02` names every accepted line the correction
may rewrite and refuses any other removal; `test_01` proves
`phase5_gate_b_scenarios.ps1`, `phase5_gate_b_diagnostics.bas`,
`com_lifecycle.ps1` and `build_stage_b.ps1` are line-for-line what the production
baseline holds.

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
fingerprint field layout**. It is not done, and those three cases emit no
fingerprint.

**What they do instead is a check, not a note.** The first submission proved only
that the duration control matched the model and then emitted a note saying
`P5-AN` had already driven the accepted expectations — a statement about a
different scenario at a different time, which says nothing about the workbook as
it stands immediately before *this* simulation. Every parity case now
establishes **current** Phase-5 authority before the run, through the accepted
comparators rather than a reimplementation:

```
PCCM_CalculationAttemptResult = SUCCESS
PCCM_CalculationStatus        = CURRENT
PCCM_CalculationFingerprint   = PCCM_CurrentInputFingerprint
Add-Phase5AnalyticalChecks    against that plan case's emitted expectations
Add-Phase5SuccessStateChecks  against the committed calc_state record
```

Using the accepted Phase-5 expected outputs is not a second fingerprint
implementation; it is exactly the current-fixture identity check `P5-AN` is
already trusted for. The golden case keeps its independent reference-digest
check **in addition**. `test_46` proves all of it runs before
`PCCM_RunSimulation` and before the comparison, and that the comparators are not
inside the golden-case branch; `test_27` proves exactly one case claims the
derivable binding.

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
| `P6-ORA` | **Excel equals the oracle** — digest, seeds, versions and the full ladder, both measures, all four cases, each preceded by a current-fixture analytical identity check | cells vs corpus |
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
| `P6-FIN` | one result per functional ID, all PASS, nothing skipped, fixtures all restored | ledger, guarded |
| `P6-LDG` | no scenario ID was recorded twice | ledger, emitted last by the driver |

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

### 4.4 Restoration is reached on every path, and a failure is fail-closed

Every fixture-writing scenario follows one policy: **capture the typed original →
mark it written → write → verify the write took → run → capture POST evidence
BEFORE cleanup → restore → verify the restoration is exact.**

**The restore lives in a `finally`, and that is the correction.** The first
submission put it on the success path only, so any exception after the fixture
write — a COM raise from `Application.Run`, a state capture, an assertion helper,
evidence formatting — jumped straight to `catch` and left the modified `F21` or
Last Run ID in the workbook, with every later scenario running against it.
Proving a restore *call exists* is not proving it is *reached*.

Three further details carry weight:

* the fixture is marked written **before** the assignment, because a COM write
  that raises may still have changed the cell;
* the establishment verdict is **load-bearing**. A COM write can return normally
  and still leave a cell holding something other than what was asked for — a
  coerced type, a rejected value, a protected sheet. The first version recorded
  the failed check and then ran `PCCM_RunSimulation` anyway, producing
  behavioural observations against a fixture the harness had itself just proved
  was not established. The gate now throws before any state capture or
  invocation, so the sequence is *write → verify → only then production*, while
  the `finally` still restores and still verifies;
* `Restore-Phase6CellFixture` never throws — a raising write or read-back becomes
  a **failed check**, so cleanup cannot replace the original scenario failure
  with its own story, and the original is re-added to the checklist afterwards;
* if the restoration cannot be **verified**, `Set-Phase6Contaminated` latches a
  harness-integrity flag. Every remaining stateful scenario then records
  `FAIL / not attempted` instead of running, and `P6-FIN` fails on the flag.
  Lifecycle and finalisation still run: the point is to stop making claims about
  behaviour, not to stop reporting. Behavioural evidence from a state the harness
  put there is worse than none, because it looks like evidence.

`test_38`, `test_38b`, `test_38c` and `test_39` pin the design; mutations 37,
37b–37f, 38, 38b, 38c and 39 require refusal of every weakening, including the
submitted success-path-only shape.

### 4.5 Completeness and the ledger, in that order

`P6-FIN` is emitted through the **guarded** `Add-Phase6Result`, so a duplicate
attempt at it is a recorded violation. `P6-LDG` is emitted by the **driver**,
last of all, through the unguarded `Add-Result`, so the ledger can see that
duplicate and can never suppress the result that reports on the ledger.

Neither is a member of the functional set `P6-FIN` verifies: making the ledger's
verdict a precondition of the completeness verdict that precedes it is the
circular ordering Round 4A removed from Phase 5, and reproducing it here would be
fail-open in the same way.

### 4.6 Two commits, and they are not the same commit

`P6-ART` reports the **production baseline** — a pinned review authority,
`bc7949b`, checked against the Python `PRODUCTION_BASELINE` by `test_52` — and
the **runtime harness commit**, `git rev-parse HEAD`, separately and by name. The
first submission passed HEAD in as `SourceCommit` and printed it as the
production baseline, which conflated the two identities the authorisation
requires to stay distinct.

The source binding is **by blob identity, not by module name**. That a compiled
project contains a module called `modSimReport` says nothing about whose
`modSimReport` it is, so each accepted Phase-6 production module's blob id in the
runtime checkout is compared against the same path at the baseline
(`git rev-parse <baseline>:pccm/src/vba/<module>.bas`). git computes both sides
under identical attribute rules, so no line-ending or encoding difference can
masquerade as a match.

**The whole-tree statement runs from the repository root**, and that correction
matters more than it looks. `git -C <path>` runs git as though it had been
started in `<path>`, and a `git diff` pathspec resolves relative to that
directory. The first version ran with `-C <repo>/pccm` and the repository-root
pathspecs `pccm/src` and `pccm/spec`, which resolve to `<repo>/pccm/pccm/src` and
match nothing — and a pathspec matching nothing produces no diff, so `--quiet`
exited 0 and the freeze check passed whatever the tree held. Fail-open, on the
one statement the whole claim rests on. The driver now derives `$repoRoot` and
passes it; every git invocation in the block uses it; the freeze additionally
asks `git ls-tree` how many files the pathspec covers, because a freeze proved by
a pathspec that names nothing is not a freeze.

Two of the controls measure git rather than argue about it. `test_57` runs both
command shapes against a path that genuinely differs from the baseline and
requires the corrected form to detect it and the submitted form not to.
`test_58` clones the repository into a throwaway directory, edits
`modWorkbook.bas` — a production file no `modSim*` blob check covers — and
requires the corrected form to fail on it and the submitted form to miss it, so
the whole-tree claim does not rest on the eight module checks.

**No git means no pass.** If git is unavailable the harness commit stays empty
and `P6-ART` FAILS. A runtime result with no attributable revision is weaker
evidence, and recording "unknown" while passing would hand that weakness on as
though it were strength.

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

`tests/test_phase6_gate_b_harness_source.py` — **61** controls, in six groups:
the accepted harness is not rewritten; the harness restates no address, name or
expected value; the failpoint and procedure names are checked copies; the
projection agrees with the generated authority; the corpus is generated, bound
and exact; the matrix is complete and fail-closed.

`tests/test_phase6_gate_b_harness_source_validation.py` — **106** mutation
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

The harness review added a further group, one per blocker: the live prerequisite
restored to the full Y/Z-inclusive Phase-4 set, a deferral check neutered to an
empty list, the post-session Phase-5 results demanded early; the fixture unwind
moved out of its `finally` onto the success path, the written flag set only after
a successful write, a failed restoration that stops latching contamination, a
latch that does not latch, a guard that lets a stateful scenario run on
contaminated state, a raising restore that escapes, a cleanup that discards the
original failure; `P6-FIN` bypassing the guard, `P6-LDG` made a precondition of
it, the ledger verdict emitted first, the driver dropping it or emitting it early;
HEAD reported as the production baseline, a drifting baseline pin, artefact
identity passing without git, the source binding falling back to module names, a
placeholder commit; the non-golden identity reduced to a note, made conditional
on the golden case, or its currency check defanged; and the final log omitting
the Phase-6 summary.

The final pre-execution review added a third group: the pending Phase-5 ledger
read as an empty list, its count no longer required to be zero, the check dropped
entirely, and `P5-LDG` emitted early to answer the question the wrong way; the
fixture establishment verdict discarded in either scenario, captured and then
reduced to a note, or gated after the simulation instead of before it; and the
whole-tree freeze returned to the `pccm` working directory — wholesale, narrowly
on the one command, with its pathspec-coverage check removed, or by the driver
passing the wrong root.

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
