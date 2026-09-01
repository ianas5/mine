# PCCM Phase 6 — Step 13: the Windows/Excel Gate-B runtime harness

**Status: CLOSED. Run 6 executed on `a3924e0` and passed 103/0/0 — Phase-4
35/35, Phase-5 39/39, Phase-6 29/29, `P6-FIN` PASS, `P6-LDG` PASS. Windows/Excel
Gate B is accepted. The production baseline remains `79e4600`.**

Six Windows runs, and the ledger in [§8](#8-windows-run-ledger) keeps every one
of them as it happened. Run 1 aborted in the preflight. Run 2 was stopped by a
production compile defect. Run 3 proved that repair on the real VBA compiler.
Run 4 was the first execution of the Phase-6 behavioural matrix and left five
failures. Run 5's first attempt stopped before Excel on a Stage-A serialisation
defect, then executed and left one. **Run 6 is all green.** None of those red
runs is rewritten here; they are what the corrections were made against.

**Runtime-proven, and this is now the whole list.** The workbook builds, opens
and compiles on the real VBA compiler; the Phase-4 lifecycle matrix runs 35/35
with a natural shutdown and a clean COM release ledger; all 39 Phase-5 Gate-B
scenarios pass; and all 29 Phase-6 scenarios pass — the public surface, the
button, FIXED and AUTO runs, bank alternation, the no-replay invariant, the
three failpoints, all five recovery classes, the durable `F21` protocol, run-ID
exhaustion, the Step-12 attempt-result axis correction, cross-implementation
parity under the accepted Step-0 evidence policy, and the result ledger.

Every correction made after Run 4 is runtime-proven with it: `P6-ORA` on the
accepted tolerance rule with the cross-language digest diagnostic rather than a
criterion; `P6-DET` on the decoupled same-runtime replay; `P6-FP3` on the
corrected restore semantics and the repaired `SameCell`; `P6-ART` on both the
pre-open workbook capture and the two-class module identity; the LF byte
serialisation, proved by the raw Windows hashes matching the accepted ones; and
the HEAD-byte provenance binding.

**What is still source-only, and stays that way.**
[§5](#5-what-remains-static-only-after-step-13) is the bounded list, and an
all-green run does not shorten it: the genuine `PERSISTENCE_INDETERMINATE` path,
the genuine COM read raises, a `ClearPending` failure after a known `CONSUMED`,
the iteration ceiling, the private `NonceConsumed` projection, the selector-write
ordering inside `FinalCommit`, and the wording and source mutation controls.
**None of those was induced**, and Run 6 does not claim otherwise. Run-ID
exhaustion and cross-implementation parity are NOT on that list — `P6-RIDMAX`
and `P6-ORA` proved them.

```
static / source evidence   !=   Windows / Excel runtime evidence
```

That line is the whole point of Step 13. It is what kept the six runs honest,
and it is why the boundary above is still drawn now that they are done.

---

## 1. What Step 13 exists to prove, and what it cannot

The Linux suite proves the pure kernels exactly, the transaction order
structurally, and the publication state machine as data. What it cannot do is
start Excel. Step 13 asks the questions only a real COM session can answer:

* does the accepted workbook build, open and **compile**;
* is the Phase-6 public surface callable;
* does a FIXED run publish a result, and does a repeat publish the same digest;
* **does real Excel agree with the accepted Python oracle under the Step-0
  cross-implementation evidence policy** — exact on the identity and discrete
  subjects, within §10.3 on every floating row;
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

### 2.3 The authorised changes to the accepted driver

Step 13 opened with one, and the Windows runs added two more. As it stands,
`phase4_functional_test.ps1` gains a dot-source, a preflight call, three artefact
loads, one harness-commit capture, one pre-open artefact capture, one scenario
call, one ledger-verdict call and a summary that names Phase 6.

| change | why | round |
|---|---|---|
| the Phase-6 wiring — dot-source, preflight, artefact loads, scenario call, ledger verdict, summary | Step 13 itself | original |
| the pre-open artefact capture, and the two arguments that carry it | Run 4: `Get-FileHash` cannot read a workbook Excel holds open | Run-4 correction |
| the harness-commit capture moved above the pre-Excel gate, and the host-local oracle joining the required-artefact list | the gate has to refuse a stale or dirty host-local oracle before Excel starts | settlement |

`test_02` enumerates every accepted line the corrections may rewrite and refuses
any other removal; `test_01` proves
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
| `publication` | bank labels, the candidate-target selector map **as a list of `{active_bank, candidate_bank}` entries**, the final-commit range and its nine fields |
| `controls` | `monte_carlo_iterations` and `random_seed`: defined name, sheet, cell, type |
| `command_surface` | the automation endpoint and the six read accessors |

#### Why the selector map is a list and not an object

Step-13 Run 1 never reached Excel. It failed in the preflight with

```
System.Management.Automation.PSArgumentException:
Cannot process argument because the value of argument "name" is not valid.
```

The contract keys `candidate_target` by the **active bank**, and the key for
*"no bank has ever been published"* is the empty string — semantically correct,
and correct to keep. Projected as a JSON object it became a property whose
**name** was `""`, and Windows PowerShell 5.1's `ConvertFrom-Json` cannot
materialise such an object as a `PSCustomObject` at all. `-AsHashtable` is a
PowerShell 6.0 switch; the accepted runtime target is 5.1, so it is not a way
out.

The projection now emits the same mapping as entries:

```json
"candidate_target": [
  { "active_bank": null, "candidate_bank": "A" },
  { "active_bank": "A",  "candidate_bank": "B" },
  { "active_bank": "B",  "candidate_bank": "A" }
]
```

The absence moves from a JSON **property name** to a JSON **null value** — the
same fact in a shape 5.1 can read. Nothing is renamed to a sentinel like
`"BLANK"`: inventing a replacement token would put a second semantic authority
in the projection. `_candidate_target_projection` knows nothing about A, B or
which follows which; it walks whatever mapping the contract declares, in the
contract's order. `sim_contract.yaml` is unchanged.

`Get-Phase6CandidateTarget` consumes the entries, normalises a blank runtime
selector **only for comparison** with the null entry, requires **exactly one**
match, and fails closed on zero and on duplicates. `test_62` proves it
structurally: the function has exactly one `return`, it reads the matched entry,
and no bank label appears as a literal anywhere in it — so no shortcut can
answer for a state without consulting the projection.

`emit_sim_inspection` now **refuses at build time** to write any object key that
is the empty string, and `test_59` scans **both** generated Step-13 artefacts
recursively: `candidate_target` is where it happened, not the only place it could
happen. `phase6_gate_b_cases.json` was and remains clean.

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

### 3.2 The parity pair — a portable authority and host-local measurements

**Two artefacts, and they are not the same kind of thing.** `sim_cases`
traverses the accepted `prepare_simulation` / `run_simulation` / `result_digest`
once and splits the result by PORTABILITY. No expected number is written by
hand, and none is written in PowerShell.

`build/phase6_gate_b_cases.json` — **the portable case authority**, required to
be cross-platform invariant and pinned by hash. It carries which cases are
driven, their fixtures, the seed and iteration count, the bounds and vocabulary,
the accepted comparison policy, and per case the accumulation scale, the
comparison classes and the **exact discrete and identity expectations** only.

`build/phase6_gate_b_oracle_local.json` — **the host-local oracle measurements**,
deliberately not invariant and never pinned. It carries the floating summary
ladder on both measures, the deterministic base, and the `result_digest` marked
diagnostic, with the provenance that binds it to a host and a run.

D2 forced that split: the same builder produced different bytes on Windows and
Linux because one artefact carried Beta-PERT output and Cheng reaches libm.
[§10](#10-d2-the-portability-settlement) states the settlement in full.

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

#### Iterations, seed and the two comparison classes

1000 iterations — asserted to be `input_contract.yaml`'s own business minimum,
never restated — and FIXED seed 12345, one of the accepted RNG vectors.

**The comparison policy is the accepted Step-0 §10 policy, and it has two
classes.** The identity and discrete subjects §10.4 keeps exact — the request
fingerprint, the effective seed, the iteration count, the RNG and method versions
— are compared by exact equality. Every cross-language floating comparison uses
§10.3's rule for summary statistics: `rel ≤ 3e-10`, or `abs ≤ 3e-10 · S` with `S`
the accumulation scale, keyed to the scale that produced the number rather than
to the number itself.

**The `result_digest` is not a cross-language equality subject.** §10.4 keeps it
exact for **same-runtime** replay, which is `P6-DET`'s property and is enforced
there without a tolerance. In the parity comparison it is recorded as diagnostic
evidence and never checked. The earlier Step-13 rule — one exact mode, digest
included — was stronger than the policy the project had accepted, and Run 4
failed on differences §10.3 had already admitted;
[§9](#9-p6-ora-reclassified-a-harness-evidence-policy-defect) records it.

`test_31` binds the emitted policy to the Step-0 record and refuses a tolerance
KEY in `sim_contract.yaml`; `test_10` refuses a bound spelled in the harness.

---

## 4. The scenario matrix

| ID | Proves | Evidence source |
|---|---|---|
| `P6-PRE` | the **derived** Phase-4 prerequisite partition is complete and the Phase-5 block is intact — `Y`, `Z` and the final 35/35 are POST-SESSION and are proved later by `P5-FIN` | recorded results + the pending Phase-5 ledger |
| `P6-ART` | source commit, Stage-A hash, executed `.xlsm` hash, manifest and all three artefact hashes, and the host-local oracle bound to this run's HEAD | pre-open capture + `Get-FileHash` on the driven copy |
| `P6-CMP` | the project that compiled contains the eight `modSim*` modules — **derived from `P5-CMP`** | no second compile |
| `P6-M` | the proved inventory is the manifest's 23-module set — **derived from `P5-M`** | no second inventory |
| `P6-API` | seven procedures declared; six accessors callable; the endpoint deferred to `P6-FX1` | persisted project + `Application.Run` |
| `P6-BTN` | no shape invokes any Phase-6 procedure | shape `OnAction` |
| `P6-INIT` | a workbook that has never simulated | `_SimData` cells |
| `P6-FX1` | the first real `PCCM_RunSimulation`: bank A published, run id 1, FIXED seed recorded | cells |
| `P6-ORA` | **Excel agrees with the oracle under the accepted Step-0 policy** — identity and discrete fields exact, the full ladder on both measures within §10.3, all four cases, each preceded by a current-fixture analytical identity check; the digest is recorded, not compared | cells vs the portable authority and the host-local measurements |
| `P6-DET` | the same inputs and seed twice produce the same digest — **same-runtime replay, exact, and no cross-language clause** | cells |
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
`79e4600` since the `SameCell` repair, checked against the Python
`PRODUCTION_BASELINE` by `test_52` — and the **runtime harness commit**,
`git rev-parse HEAD`, separately and by name. Run 1's evidence in §8.1 names
`bc7949b`, which was the baseline at that time; that is history and stays as it
was written. The
first submission passed HEAD in as `SourceCommit` and printed it as the
production baseline, which conflated the two identities the authorisation
requires to stay distinct.

**The eight production modules do not have the same KIND of identity**, and Run 5
failed because `P6-ART` assumed they did. Seven are hand-written, tracked source
in `pccm/src/vba`, and are proved by git blob identity against the baseline.
`modSimContract` is a **generated Stage-A projection** of the accepted contracts:
it is emitted to `build/vba`, declared `generated: true` in the manifest, and has
no path in `src/vba` at any commit — so asking git for its blob returned blank on
both sides and the comparison passed a blank against a blank.

Its identity is the **canonical projection the baseline produces**: SHA-256 with
line endings normalised to LF, because the physical `.bas` is text mode for
`AddFromString` and its raw bytes are host-dependent by design.

```
accepted projection identity (derived from 79e4600)
  -> the manifest's `generated: true` entry
  -> the generated-source directory Stage B resolves from it
  -> the file this session consumed
```

`test_77` archives baseline `79e4600` into an isolated tree, runs **that**
commit's Stage-A build with **that** commit's contracts, and canonicalises the
module it emits: `daa4d27889c30eadb2ab892bcfa4e6f6bab8a137aae79a01a8d8f1e8e1c215ac`.
The harness carries that value as a checked copy. Deriving it from HEAD's
renderer would let a changed renderer bless its own changed output, which is
exactly the case that matters — the builder has legitimately moved since
`79e4600` while the production projection has not.

The source binding is **by blob identity, not by module name**, for the seven.
That a compiled project contains a module called `modSimReport` says nothing about whose
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

`tests/test_phase6_gate_b_harness_source.py` — **82** controls, in sixteen groups:
the accepted harness is not rewritten; the harness restates no address, name or
expected value; the failpoint and procedure names are checked copies; the
projection agrees with the generated authority; the corpus is generated, bound
and exact; the matrix is complete and fail-closed.

`tests/test_phase6_gate_b_harness_source_validation.py` — **223** mutation
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

Runtime Run 3 added ten controls to the Phase-5 harness suite
(`test_ev_01`..`test_ev_10`): the contract grants `RunSimulation` to exactly one
owner and the emitted manifest agrees; every scoped grant is checked **as a
grant**, with the predicate chosen from the contract rather than from memory; a
globally forbidden construct is still checked globally; the generic
module-aware scan stays load-bearing; and six mutations — the stale global
assertion restored, a wrong owner named, a check whose wording contradicts its
predicate, the contract's owner set widened, the generic scan traded for the
scoped ones, and a global rule turned into a grant.

The final pre-execution review added a third group: the pending Phase-5 ledger
read as an empty list, its count no longer required to be zero, the check dropped
entirely, and `P5-LDG` emitted early to answer the question the wrong way; the
fixture establishment verdict discarded in either scenario, captured and then
reduced to a note, or gated after the simulation instead of before it; and the
whole-tree freeze returned to the `pccm` working directory — wholesale, narrowly
on the one command, with its pathspec-coverage check removed, or by the driver
passing the wrong root.

Run 1's pre-Excel failure added a seventh: the selector map restored as an object
with the blank key, an empty key planted elsewhere in the projection, the blank
entry dropped or duplicated, a `"BLANK"` sentinel, a lost mapping, an unapproved
entry key, and — on the PowerShell side — the blank answer hard-coded, the empty
property lookup restored, and either fail-closed arm removed.

Run 3 added an eighth, and it watches this document rather than the harness.
`test_66` reads the preamble **above §1 only** and checks it against the run
ledger below it: once the ledger records a completed attempt, an unscoped denial
that anything has executed is refused; the preamble must name every attempted
run, say the execution is **partial**, bound the runtime-proven claims to what
those runs established, name `P6-ART` through `P6-AXIS` as still unexecuted, and
scope its source-only claim to Phase-6 *behaviour* rather than to Step 13 whole.
Run 4 then falsified that control's own premises, which is the lesson it now
carries: its demands are **conditioned on the ledger** rather than written down.
A run that reports a Phase-6 tally is a run in which the matrix executed, and
from that moment the preamble must say so and must stop saying the opposite —
but must also name what is still open and say the open item is unattributed,
because the danger has flipped from under-claiming to over-claiming. Its seven
mutations restore the stale preamble, restore the wording accepted for Runs 1-3,
put back the unexecuted claim on its own, sweep the source-only claim across
Step 13, report the executed matrix without its open item, take the bound off
the runtime-proven list, and report the disagreement as attributed. Section 8 is
left byte-identical in all seven: **history is not edited when the present tense
changes**, and "Neither behavioural matrix executed" stays true of Run 2 forever.

Run 4 added a ninth, one control per defect it found. `test_67` requires the two
workbook hashes to be captured before Excel opens them, from the consumed copy
rather than `build/`, under a handler that cannot abort the run, and bound to the
workbook actually opened; `test_68` requires the `P6-FP3` preservation set to be
derived from the projection's row groups rather than named by hand, the derived
status to be unchanged, and the attempt detail to name the injected stage and not
to claim the restore failed. Twelve mutations: the open workbook hashed again,
the capture aimed at `build/`, the capture allowed to throw, a missing or
malformed capture accepted, the open-workbook binding defanged; and the
hand-written skip list restored, the projection partition traded for a literal,
an empty durable set accepted, the restore claim dropped, the derived status
allowed to drift, and the candidate claim demoted back to a note.

The Run-4 production repair is controlled in the integration battery instead,
where the module lives: `test_37` and `test_38` in
`tests/test_phase6_integration_source.py` **run** the publication verify
predicate over the Variants `Value2` returns, and five mutations
(`test_72`..`test_76`) restore the one-sided blank rule, drop the blank-written
guard, make the blank branch unconditional, drop the candidate blank-write
semantics and defang the string comparison. `test_55` in
`tests/test_phase6_sim_engine_vba.py` pins the P6-ORA localisation.

The D2 settlement added a tenth. `test_69` requires the parity comparison to use
the accepted classes — exact for the identity and discrete fields, the policy for
every floating row, and the digest recorded and never checked anywhere in the
file; `test_70` requires repeatability to be same-runtime and exact and to carry
no oracle clause; `test_71` requires the two artefacts to declare their
portability, the preflight to hash the authority the evidence names, and the
runtime scenario to record the evidence's host and source revision. `test_31` no longer
demands an exact-only corpus — it binds the emitted policy to the Step-0 §10.3
table and §10.4 exact list through `builder/pccm_builder/sim_policy.py`, and
checks `sim_contract.yaml` still carries no tolerance KEY. `test_48` pins the two invariant artefacts and refuses
a pin on the host-local one.

Twelve mutations: exact cross-language digest equality restored, a floating
summary row returned to exact equality, a `1e-9` spelled in the harness, the
scale-aware floor dropped, the same-runtime digest equality weakened, the oracle
digest dependency returned to `P6-DET`, host-sensitive numbers put back into the
portable authority, the authority claiming a portability it does not have, the
evidence generated for a different authority, the preflight's artefact binding
removed, the emitted policy drifted from the record, and the policy promising a
cross-language digest again.

The settlement round added an eleventh, and it watches WORDING and PROVENANCE.
`test_45` reads the document's current architectural and scenario-definition
sections — everything except §8, which is history and is never rewritten — and
the live source banner, and refuses the superseded rule in either: an
affirmative exact-only policy, a matrix row claiming Excel equals the oracle, a
repeatability row reaching for the oracle's digest, a banner denying every
execution, a statement sending every expected value to one artefact, and `P6-PRE`
described as the pre-Excel gate when that gate is `PRE6`. `test_72` requires the
host-local measurements to carry a real 40-character commit this repository
contains, the identity pair to agree on all ten named items, and BOTH gates to
compare the revision against the runtime harness HEAD as a check over one
representation — because provenance that is only printed is attribution, and a
stale oracle from a harness-only commit passes every other check unchanged.

Fifteen mutations: each of the six superseded statements restored, and
`unavailable` as a revision, a commit the repository does not contain, the
binding weakened to a prefix, the pre-Excel comparison removed, the runtime
scenario reduced to printing, and the pair disagreeing on the supplied seed, the
schema version, the case authority and the policy authority.

A twelfth closes the gap between a commit id and a set of bytes. `git rev-parse
HEAD` says which commit is checked out and nothing about whether the tracked
files that produced a measurement — or the tracked files about to execute — are
the files that commit holds, so two paths pass a revision comparison unseen:
generate from a modified builder and revert before the run, or generate cleanly
and then modify a tracked harness file. The clean fact is now established twice.
**At generation**, `_generation_provenance` compares the tracked `pccm` subtree
against HEAD and writes `source_tree_clean` into the artefact, where reverting
the change afterwards cannot retract it; it fails closed on every path, and the
only `True` it can return is the one the diff produced. **Before Excel**, `PRE6`
requires that fact AND runs `git diff --quiet HEAD -- pccm` over the tree about
to execute, with the pathspec proved to match something so the Run-2 fail-open
cannot reappear; `P6-ART` re-asserts both inside the session. Untracked output —
an ignored `build/`, a retained run log — is not reported by `git diff` and
correctly does not make the source dirty.

`test_73` proves the mechanism and refuses a package whose measurements were not
generated from a clean tree. Eight mutations: the clean fact dropped from the
artefact, a dirty generation, the builder hardcoding `True`, the builder falling
back to `rev-parse` alone, each gate's runtime check defanged, the generation
demand defanged, and the running-tree pathspec narrowed to `pccm/src`.

A thirteenth separates two gates that share a prefix and nothing else. `PRE6` is
the pure PowerShell artefact and provenance gate and runs before Excel exists;
`P6-PRE` is a live-session scenario reading results only the session produces.
The driver's `PRE6` exception arm recorded a `P6-PRE` FAIL and exited, putting a
verdict in the ledger for a scenario that had not been reached — on the one path
where nothing else could contradict it, and `P6-FIN` proves each required ID has
exactly one result. It now reports the way the normal refusal does, to the
console and out. `test_74` refuses ANY Phase-6 result recorded before Excel
exists, refuses `PRE6` being promoted into the scenario vocabulary, and requires
the live `P6-PRE` to remain where it belongs; `test_75` reads the driver's active
banner and refuses two claims the accepted lifecycle had made false — that
`P5-P4` checks the final 35/35 when `Y` and `Z` are deferred past it, and that no
Phase-5 Gate-B run has been made.

Six mutations: the `P6-PRE` result restored to the exception arm, any other
Phase-6 result recorded there, the abort dropped, `PRE6` promoted to a scenario
ID, and each of the two banner claims restored.

`test_45` was then extended to two active regions an all-green matrix had walked
past: the `PRE6` explanatory block, which still described the inspection
projection and the parity corpus as the only two authorities the scenarios read,
and the Step-13 source battery's own module docstring, which still said the Gate-B
run had not been made. Both described the world before Run 4 and before the D2
split. Two more mutations restore each. A detector that reads only the regions it
was first pointed at goes stale exactly where nothing is looking.

Run 5 added a fifteenth. `test_77` proves the generated `modSimContract`
projection has a BASELINE-BOUND identity: it archives `79e4600` into an isolated
tree, runs that commit's Stage-A build with that commit's contracts, and requires
the canonical projection to equal the value the harness carries as a checked
copy. Eleven mutations: the projection returned to the tracked blob loop, its
identity check defanged, a manifest entry that is not `generated`, zero or
duplicate entries, the module resolved from the source tree, the expected
identity derived from HEAD, the raw Windows SHA pinned as the identity, a
hand-written blob check weakened, and the canonicaliser accepting a BOM, a bare
CR, a missing final newline, or normalising more than line endings.

**"Line endings" means LF and CRLF, and not a bare CR.** The first canonicaliser
wrote `.replace(b"\r\n", b"\n").replace(b"\r", b"\n")`, which quietly admitted a
third representation: a module whose every LF had become a lone CR hashed
identically to the accepted projection. A CR that is not part of a CRLF now
refuses the artefact, on both sides. `test_77` no longer reads the
implementation's source shape for this — the version that did REQUIRED the
offending replace, so it pinned the defect rather than catching it, and would
have gone on passing because the docstring explaining the defect contains the
string it searched for. It calls the function instead: CRLF and LF must agree; a
bare CR, a CR inside a line, a BOM and a missing final newline must each be
refused; a real change must alter the identity; and case, whitespace and
encoding differences must not.

Run 6 added a sixteenth, and it guards the one thing an all-green run makes
tempting. `test_78` reads the closure against the ledger in BOTH directions: the
closing run must name a real commit and the tallies it actually reported, every
scenario that was once open must be recorded PASS in that run's evidence and
named runtime-proven in the preamble, the status must state closure, the
production baseline must be named and must not have moved — and the static-only
boundary must survive the pass, subject by subject, with `P6-CMP` and `P6-M`
still derived from one compile and one inventory. `test_66` gains a third
ledger-driven state alongside "unexecuted" and "executed with open items":
CLOSED, where the open-item demands would otherwise force the document to
describe a state it has left.

Twelve mutations: the closing run recorded without a commit, a failure or a skip
in its row, a short Phase-6 tally, a settled scenario still reported open, a lost
PASS record, a status line that stops stating closure, a closure that does not
name its baseline, each of the six static-only subjects dropped, a claim that the
unreachable arms were induced, and run-ID exhaustion or cross-implementation
parity put back on the static-only list.

**An all-green run is where over-claiming costs most.** Every hedge looks
removable when nothing failed, and the two that must not be removed were never
about failure: the arms this harness cannot reach, and the single compile and
single inventory the derived scenarios rest on.

Two of those eleven found controls reading an `Add-Check`'s LABEL while its
predicate had been replaced by `$true` — the tracked-module blob comparison and
the manifest entry count. Both now read the joined statement. It is the third
time this shape has appeared, and it is always the same lesson: a label survives
the change that matters.

---

## 7. Running it

One command set, one working directory — the repository root:

```
python pccm\builder\build_stage_a.py
powershell -ExecutionPolicy Bypass -File .\pccm\bootstrap\windows\build_stage_b.ps1
powershell -ExecutionPolicy Bypass -File .\pccm\bootstrap\windows\phase4_functional_test.ps1 `
  *> .\pccm\bootstrap\windows\phase6_gate_b_run6.log
```

The third command runs everything: the Phase-6 block is dot-sourced into the
Phase-4 harness and runs inside its single COM lifecycle, against the disposable
`%TEMP%` copy. There is no separate Phase-6 script to invoke and no second Excel
instance.

**Step 13 is closed and this command is the record of what was run, not an
invitation to run it again.** A re-execution would be a new run under a new
number, with its own log; nothing here may be overwritten.

**The log name carries the run number**, and it is `run6` because Runs 1 to 5
already happened — see the ledger in §8. Writing to an earlier run's log would
overwrite evidence of an attempt that was made, and an aborted or blocked
attempt is still evidence: Run 1 identified the defect in §3.1, Run 2 the two in
§8.2, Run 3 the one in §8.3 and Run 4 the three in §8.4. Each authorised run gets its own log, and no
earlier one is renamed or overwritten.

**Prerequisite.** Importing VBA requires *Trust access to the VBA project object
model*. The scripts report it and stop if it is missing. They do not enable it,
do not lower macro security, do not edit the registry and do not add a Trusted
Location — the same refusal that has held since the first readiness run.

---

## 8. Windows run ledger

| Run | Source | Outcome |
|---|---|---|
| **Run 1** | `6365aeb` | **ABORTED PRE-EXCEL at `P6-PRE`.** Excel was never started. |
| **Run 2** | `849d6bf` | **VALID runtime attempt.** Reached Excel; Phase-4 35/35; a production compile defect stopped `P5-CMP`; the Phase-5 and Phase-6 behavioural matrices were NOT executed. |
| **Run 3** | `58a89f3` | **VALID runtime attempt.** `P5-CMP` PASS — the compile repair is runtime-proven; Phase-5 38/39; a stale `P5-EV` assertion blocked Phase 6; `P6-LDG` PASS. |
| **Run 4** | `6cb7f06` | **VALID runtime attempt, and the first behavioural one.** 98 passed, 5 failed; Phase-4 35/35; Phase-5 39/39; Phase-6 24/29; `P6-LDG` PASS. |
| **Run 5 (aborted)** | `5c63503` | **AUTHORISED; STOPPED PRE-EXCEL.** Stage A 351/351 on a clean Windows checkout and the generation provenance was correct — but the raw `Get-FileHash` of both invariant artefacts disagreed with the accepted hashes. Stage B was NOT run and Excel was NOT started, so no Phase-6 scenario failed: a Stage-A serialisation defect, found before runtime. |
| **Run 5** | `253b022` | **VALID runtime attempt.** Stage A 351/351, raw D1 and D2 as accepted, Stage B PASS, Excel started. 101 PASS / 2 FAIL / 0 SKIP — Phase-4 35/35, Phase-5 39/39, Phase-6 27/29. `P6-ORA`, `P6-DET` and `P6-FP3` all PASS on the corrected rules. `P6-ART` FAIL on the generated `modSimContract` treated as tracked source; `P6-FIN` derivative. |
| **Run 6** | `a3924e0` | **ALL GREEN.** 103 passed, 0 failed, 0 skipped. Phase-4 35/35, Phase-5 39/39, Phase-6 29/29. `P6-ART` PASS, `P6-FIN` PASS, `P6-LDG` PASS. Natural Excel shutdown, 4/4 transient COM releases, no emergency shutdown. **Windows/Excel Gate B accepted; Step 13 closed.** |

An aborted attempt is not nothing, and it is not renamed away when a later run
succeeds: Run 1 is what found the defect §3.1 records, and it stays in this
ledger under its own number.

**`d047eea` is not Run 2's source.** It was created after Run 2 had already
happened, and it is documentation and control work only. Run 2 ran `849d6bf`,
and the ledger says so.

### 8.1 Run 1 — aborted before Excel

**Not runtime evidence.** No Excel process was started, no VBA executed, and no
simulation behaviour was exercised. What Run 1 produced is a defect report about
this harness.

```
PRE0    PASS
PRE     PASS
P5-PRE  PASS
PRE6    FAIL   System.Management.Automation.PSArgumentException:
               Cannot process argument because the value of argument "name"
               is not valid.
        → the harness aborted before Excel, as designed
```

The cause is §3.1's empty-string object key, and the correction is described
there. The classification is worth stating plainly because the failure mode
matters more than the fix:

* a real Step-13 **evidence-infrastructure** defect;
* purely pre-Excel, a PowerShell-compatibility defect;
* **no production VBA defect is supported by it**, and none was changed;
* production baseline `bc7949b` stayed frozen throughout.

**What let it through.** 106 mutation controls and 61 source controls, all
green, none of which asked whether the only consumer of these artefacts could
parse them. Every control was about *what the projection says*; none was about
*what the reader can read*. The build now refuses the shape outright, and both
artefacts are scanned recursively rather than at the one key that failed.

### 8.2 Run 2 — reached Excel; stopped at the compile prerequisite

**A valid runtime attempt, and the first one.** Excel started, the Stage-B build
ran, and the Phase-4 structural matrix executed in full.

```
PRE0, PRE, P5-PRE, PRE6            PASS
Stage-B runtime build              PASS
Phase-4 prerequisite matrix        33/33 PASS
deferred lifecycle Y, Z            PASS
final Phase-4 matrix               35/35 PASS
natural Excel shutdown             PASS
transient COM release ledger       PASS
```

`P5-CMP` reached the correct PCCM VBProject and the correct active project
identity, found VBE command id 578 as an `msoControlButton` of Type 1, and
called `Execute()` exactly once. The control then stayed Enabled for its
five-second observation window.

**That was not a timing artefact.** The retained Run-2 workbook was opened
manually, without editing, and `Debug > Compile VBAProject` was invoked once.
The VBE produced a genuine compiler diagnostic:

```
Compile error: Argument not optional
```

selecting the call inside `modSimNonce.ReadPending`.

#### 8.2.1 Defect one — a production compile defect

`modWorkbook.IsWholeInRange` declares four arguments, the last a
`ByRef Result As Double`. Five Phase-6 call sites passed three:

| Module | Procedure | Passed |
|---|---|---|
| `modSimNonce` | `ReadPending` | 3 of 4 |
| `modSimNonce` | `ReadShared` | 3 of 4 |
| `modSimReport` | `ResolveIterations` | 3 of 4 |
| `modSimReport` | `ResolveSeed` | 3 of 4 |
| `modSimReport` | `ReadMachineLong` | 3 of 4 |

The Windows report named the first two. The full arity audit found the other
three, all the same class. The remaining fifteen call sites in
`modCalcResolve`, `modDrivers`, `modInflation`, `modStructuralCheck` and
`modWorkbook` were already correct, and a bounded audit of every Phase-6 call
into public `modWorkbook` procedures found no other malformed arity.

**VBA compiles on demand.** A procedure body nothing has reached yet can hold a
fatal call for as long as nothing reaches it, which is how five of them survived
every static control up to Run 2 — and why `P5-CMP` exists to make the compile
claim separately, once, before anything relies on it.

**What let it through.** The controls that existed proved the call was
*present*: `assert "IsWholeInRange" in body`. None proved it was *well-formed*.
The replacement walks every qualified cross-module call in the hand-written
production source — 446 of them — and checks the argument count against the
callee's own declaration, honouring `Optional` and `ParamArray`. Pinning the
five corrected lines as text would have closed five call sites and left the
class open one compiler stop away.

The repair is the minimum: each site already declared a local `Double` for the
parsed value and read it back through `TryReadDouble` on the next statement, and
that local is now passed as the fourth argument. No bound, no nonce semantic, no
recovery classification, no transaction semantic, no contract and no
`modWorkbook` signature changed. The now-redundant `TryReadDouble` is left in
place: removing it is a behaviour-preserving tidy-up, and a compile repair is
not the place for one.

#### 8.2.2 Defect two — a harness finalisation defect

After the compile prerequisite correctly prevented the behavioural matrices from
running, lifecycle finalisation reached `Z`, `Y`, `P5-FIN` and `P5-LDG`, all
PASS. Then the driver itself failed under StrictMode, before `P6-LDG` could be
emitted:

```
PropertyNotFoundException:
The property 'Count' cannot be found on this object.
```

`Get-Phase6LedgerViolations` returns `@($script:Phase6LedgerViolations)`, but a
function *returning* an empty collection emits **zero pipeline objects**, so
`$violations = Get-Phase6LedgerViolations` landed `$null` and `.Count` threw.
Zero violations is the **normal** case, so this failed on every clean run.

It is the accepted Phase-4 rule — collections are materialised at the caller —
and this call site did not follow it. `Add-Phase6LedgerIntegrityResult` now
writes `@(Get-Phase6LedgerViolations)`, and the audit that found it also wrapped
one other unwrapped assignment in the preflight. `test_64` pins the whole class
rather than the one line, and requires the PASS arm to return, the FAIL arm to
report **every** violation, and the emitted-once flag to stay.

#### 8.2.3 What Run 2 did and did not establish

**Established:** the workbook builds and opens in desktop Excel; the Phase-4
structural runtime is intact at 35/35; the compile control reaches the right
project and executes the right command; Excel shuts down naturally and the COM
release ledger is clean.

**Not established, and not claimed:** anything about Phase-5 or Phase-6
behaviour. Neither behavioural matrix executed. No Phase-6 procedure ran, no
simulation was performed, and no parity comparison was made.

### 8.3 Run 3 — compile repair proved; one stale assertion blocked Phase 6

**The second valid runtime attempt, and it settled two open questions.**

```
Phase-4 final matrix               35/35 PASS, 0 FAIL, 0 SKIP
Y, Z                               PASS
natural Excel shutdown             PASS
transient COM release ledger       PASS
P5-CMP                             PASS
Phase-5 Gate-B scenarios           39 reported, 38 passed
P6-LDG                             PASS  (28 results recorded, 0 duplicates)
```

#### 8.3.1 Closed: the production compile repair

`P5-CMP` established the correct PCCM VBProject identity, found and executed VBE
command id 578, and the compiled state settled — enabled before `True`, one
observation over 113 ms, last Enabled `False`.

**The five-site `IsWholeInRange` repair in baseline `5a5b183` is runtime-proven
on the real VBA compiler.** That is the first Step-13 claim to move from source
evidence to runtime evidence.

#### 8.3.2 Closed: the P6-LDG finalisation defect

Run 3 reached finalisation and emitted `P6-LDG — PASS`, `scenario results
recorded: 28`, `duplicate attempts: 0`. The Run-2 `$null.Count` failure is
closed, and the PASS arm has now executed.

#### 8.3.3 Open, and corrected here: a stale scoped-grant assertion

Exactly one Phase-5 scenario failed, on exactly one check:

```
P5-EV   RunSimulation is still forbidden in every module
```

Everything else in `P5-EV` passed — including the generic module-aware scan of
the **real persisted project**, which found no forbidden construct in executable
code, and the `MRG32k3a` scoped grant.

The accepted contract already reads:

```yaml
- construct: "RunSimulation"
  allowed_in:
    - "modSimReport"
```

so the assertion contradicted a grant the contract had made at Step 11. It was
true when it was written and nothing moved it. `P6-PRE` then correctly failed
closed on `P5-EV=FAIL`, and **the entire Phase-6 behavioural matrix went
unexecuted** — the fail-closed design working exactly as intended, on a premise
that was wrong.

The correction is one assertion:

```powershell
'RunSimulation is permitted in modSimReport and nowhere else'
Test-ConstructScopedTo -Manifest $Manifest -Construct 'RunSimulation' `
                       -ModuleName 'modSimReport'
```

No contract changed, no production VBA changed, `RunSimulation` stays in
`forbidden_constructs`, its `allowed_in` is not widened, and the generic
module-aware scan is untouched and still load-bearing.

**What let it through.** The harness held two statements about the same rule:
one derived from the manifest at run time, one written by hand. The derived one
moved with the contract; the hand-written one did not, and no control compared
them. `test_ev_02` now chooses the predicate **from the contract** — a construct
with owners must be checked with `Test-ConstructScopedTo` naming that owner, a
construct with none with `Test-ConstructForbiddenGlobally` — so a future grant
cannot leave a contradicting assertion behind.

#### 8.3.4 What Run 3 did and did not establish

**Established:** the compile repair, on the real compiler; Phase-4 35/35 again;
38 of 39 Phase-5 scenarios; the `P6-LDG` correction; natural shutdown and a clean
release ledger.

**Not established, and not claimed:** every Phase-6 scenario from `P6-ART` to
`P6-AXIS` is unexecuted. No Phase-6 procedure ran, no simulation was performed,
no parity comparison was made, and `P6-ART`'s runtime source binding remains
unproved. `P6-FIN` correctly failed because the matrix did not execute.

### 8.4 Run 4 — the Phase-6 behavioural matrix executed

**The first valid execution of the Phase-6 matrix, and the boundary Step 13 was
built to reach.** A real production simulation ran in real Excel.

```
Phase-4 structural matrix          35 of 35
Phase-5 Gate-B scenarios           39 reported, 39 passed
Phase-6 Gate-B scenarios           29 reported, 24 passed
P5-LDG                             PASS  (36 results, 0 duplicates)
P6-LDG                             PASS  (28 results, 0 duplicates)
natural PID exit                   True
emergency shutdown required        False
transient COM release ledger       clean
                                   98 passed, 5 failed, 0 skipped
```

Retained working copy: `C:\Users\pcd\AppData\Local\Temp\pccm-phase4-20260830-220102`.

**Five FAIL labels, three defects.** Two of the five are derivative and are not
counted twice.

| Scenario | Classification |
|---|---|
| `P6-ART` | **harness defect** — a Windows file lock, corrected here |
| `P6-ORA` | **open** — a real cross-implementation disagreement, owner unresolved |
| `P6-DET` | **derivative of `P6-ORA`** — no determinism defect exists |
| `P6-FP3` | **two defects** — one production, one harness, both corrected here |
| `P6-FIN` | **derivative** — completeness correctly refused an incomplete matrix |

#### 8.4.1 `P6-ART` — a harness defect, and not a source-binding one

```
The file '...\PCCM_stageB.xlsm' cannot be read:
The process cannot access the file because it is being used by another process.
```

The scenario hashed the executed `.xlsm` while the functional Excel instance
held it open. Nothing about the production source binding is implicated, and no
production change is justified by it.

**The fix is not to hash something else.** The artefact `P6-ART` must identify is
the disposable copy this session actually consumed; hashing `build\` instead
would identify the directory the run was seeded from and would quietly assume the
equality the check exists to test. So the same path is hashed at the only moment
it is both built and unlocked — after the Stage-B bootstrap has closed its own
Excel and **before** the functional instance opens anything — and the captured
record is passed into the scenario. `P6-ART` now requires two well-formed
captures, binds the captured `.xlsm` to `$Workbook.FullName`, and fails closed
without them. It no longer reads either workbook. The blob-by-blob binding to the
production baseline and the separate harness-HEAD identity are untouched.

That capture is the **one authorised change to the accepted Phase-4 driver** in
this round, and `test_02` names it line by line, including the demand that it sit
above the point where the driver starts driving the workbook.

#### 8.4.2 `P6-FP3` — a production defect and a harness defect

**The production defect.** The endpoint announced

```
the final commit did not complete (Phase6FinalCommit) AND the previous shared
block could not be restored. The publication selector cannot be guaranteed and
requires recovery.
```

while the sheet showed the restore had physically succeeded: the active bank was
still the prior bank, `last_run_id` was unchanged, the prior publication was
intact and the candidate was unpublished.

`BuildCommitBlock` writes `vbNullString` into the blank fields of a **candidate**
block. `FinalCommit` captures the previous block with `Range.Value2`, which
returns `Empty` for those same fields. Both are verified with one predicate, and
that predicate recognised only the built spelling of blank:

```vba
If IsEmpty(written) Then
    SameCell = (VarType(wanted) = vbString And Len(CStr(wanted)) = 0)
```

so `Empty` against `Empty` — a blank correctly restored over a blank — was false.
The correction adopts the rule `modCalcReport` has carried since Phase 5: blank is
a **set**, tested on both sides. No tolerance, no weakened verification, the
candidate `vbNullString` semantics retained, and the bank and run-id comparisons
still exact. Every caller was audited; `modCalcReport`'s own `SameCell` already
held the correct rule and was not changed.

`test_37` reads the predicate out of the module and **runs it** over the Variants
`Value2` returns. Its model of `IsNumeric` deliberately *raises* on `Empty`, so
the predicate has to settle every blank case before reaching a coercion Linux
cannot adjudicate.

**The harness defect.** The preservation set was too broad. After `FinalCommit`
returns `False`, production runs `RecordFailure` → `WriteAttemptBlock`, which
rewrites the whole attempt and status range in one assignment — every `attempt`
row and both `derived` rows, the second of which is a fresh `Now`. Run 4 failed on

```
status_evaluated_at   before 46264.92150462963
                      after  46264.921527777777
```

which is correct bookkeeping, not a restore failure. The scenario now derives the
durable set **from the projection's row groups** — the shared rows production's
attempt range does not contain, which is the run identity counter and the
publication selector — so a row that changed group in the contract moves between
the two demands on its own instead of leaving a stale list behind.

It also now demands what the failure must positively record: the derived status
is unchanged (nothing was published, so the state over the same inputs and the
same published bank must not move), the stamp is still populated, the detail
names the injected stage, the candidate is not the published bank, and **the
detail does not claim the restore failed**. That last one is what makes the
production repair load-bearing at runtime: without it the repair could regress
and `P6-FP3` would still be green.

#### 8.4.3 `P6-DET` and `P6-FIN` — derivative, and not defects

`P6-DET` proved what it exists to prove: both runs announced success and **both
published the same result digest**. Its only failed assertion was that the
repeated digest equals the **oracle's**, which is `P6-ORA` restated. FIXED
repeatability is runtime-proven and no determinism or RNG change is justified.
`P6-FIN` failed because required scenarios above it were red, which is
completeness working. Neither is touched.

#### 8.4.4 What Run 4 did and did not establish

**Established:** a real production Phase-6 simulation executed; the public
surface, the button, publication and bank alternation, the AUTO lifecycle and its
no-replay invariant, all five recovery classes, the durable `F21` protocol,
run-ID exhaustion, the Step-12 attempt-result axis correction and the result
ledger; FIXED repeatability; Phase-5 39/39; a natural shutdown with a clean
release ledger.

**Not established, and not claimed:** cross-implementation numerical parity. The
`P6-ART` source binding is still unproved at runtime — Run 4 never reached the
check that would have proved it. The impossible-to-induce COM clauses remain
static-only and are not upgraded.

---

### 8.6 Run 6 — all green, and Gate B accepted

**The closing run.** Executed on harness `a3924e0`, production baseline
`79e4600`, from a clean checkout.

```
Stage A                            351 passed, 0 failed
raw phase6_gate_b_inspection.json  83eff35f...30111573   as accepted
raw phase6_gate_b_cases.json       6a9d8678...c495af5c   as accepted
modSimContract canonical           daa4d278...e1c215ac   as accepted
source_revision                    a3924e0e691b9db215245f764a36d51be14af6e2
source_tree_clean                  true
Stage B                            23 modules, 14 CodeNames, 5 buttons
Phase-4 structural matrix          35 of 35
Phase-5 Gate-B scenarios           39 reported, 39 passed
Phase-6 Gate-B scenarios           29 reported, 29 passed
P6-FIN                             PASS  (every required ID once, none skipped)
P6-LDG                             PASS  (28 scenario results, 0 duplicates)
                                   103 passed, 0 failed, 0 skipped
```

Fixture restoration completed on every scenario that established one.

The retained working copy is
`C:\Users\pcd\AppData\Local\Temp\pccm-phase4-20260901-233556`.

#### 8.6.1 What Run 6 closed

Every correction made after Run 4, proved on real Excel:

| Correction | Scenario | Found by |
|---|---|---|
| the accepted Step-0 evidence policy, digest diagnostic not criterion | `P6-ORA` PASS | Run 4 |
| same-runtime replay decoupled from the Python oracle | `P6-DET` PASS | Run 4 |
| the `SameCell` blank-restore repair and the derived preservation set | `P6-FP3` PASS | Run 4 |
| the pre-open workbook identity capture | `P6-ART` PASS | Run 4 |
| the two-class production module identity | `P6-ART` PASS | Run 5 |
| LF byte serialisation of the invariant artefacts | raw Windows hashes as accepted | Run 5 (pre-Excel) |
| the HEAD-byte generation and runtime tree binding | `P6-ART` PASS | settlement |

`P6-ART` also passed the whole `pccm/src` + `pccm/spec` freeze against `79e4600`,
the executed `.xlsm` identity, the oracle source-revision binding, the host-local
declaration, portable D2 invariance and D1 provenance. The generated
`modSimContract` passed as a projection: one manifest entry, `generated: true`,
the Stage-B source actually consumed, and the canonical identity the baseline
produces. Its raw Windows SHA `CC74EEC4...` is recorded and remains
host/text-mode diagnostic.

`P6-CMP` and `P6-M` stay **derived** from `P5-CMP` and `P5-M` — one compile,
command id 578, one observation, and one inventory. Run 6 adds no second compile
and no second inventory.

#### 8.6.2 What Run 6 did NOT establish

The [§5](#5-what-remains-static-only-after-step-13) list is unchanged. An
all-green run proves the scenarios that ran, not the arms that cannot be reached:
no genuine `PERSISTENCE_INDETERMINATE`, no genuine COM read raise, and no
`ClearPending` failure after a known `CONSUMED` was induced, and none is claimed.
Reading a pass as though it covered them would be the exact over-claim this
document has refused for six runs.

---

## 9. `P6-ORA` reclassified: a harness evidence-policy defect

**Not a production defect, and not an oracle defect.** Run 4's failure was Step
13 comparing under a rule stronger than the one the project had already accepted.

### 9.1 The governing clauses, quoted

`docs/phase6_step0.md` §10, *Numeric comparison tolerance — settled as an
evidence policy*:

> **§10.1 Ownership.** The tolerance is not a simulation-runtime contract and
> does not belong in `sim_contract.yaml`. The engine never compares two Doubles
> for approximate equality at runtime: replay comparison is by `result_digest`,
> which is **exact** … A tolerance exists only when two **implementations** are
> compared — which is oracle, Gate-A and Gate-B evidence. **Single owner: the
> Phase-6 oracle and evidence policy.** `sim_contract.yaml` stores **no tolerance
> at all**, so the rule cannot come to live in two files.

**§10.3, the policy:**

| Subject | Rule |
|---|---|
| individual Uniform / Triangular / PERT-rescale transformed samples | `rel ≤ 1e-12`, or `abs ≤ 1e-12 · s`, `s = max(\|a\|,\|m\|,\|b\|)` |
| deterministic Cheng vector outputs | `rel ≤ 1e-11` |
| F1 per-iteration no-Beta end-to-end totals | `rel ≤ 3e-10`, or `abs ≤ 3e-10 · S` |
| **summary statistics compared cross-language** | **`rel ≤ 3e-10`, or `abs ≤ 3e-10 · S`**, same accumulation-scale floor |

**§10.4, what stays exact** — MRG32k3a state and uniform values; jump state;
Bernoulli occurrence decisions; proposal and draw counts where the arithmetic
path is fixed; and **same-runtime G2/G3 `result_digest`**.

That last qualifier is the defect. Step 13 read `result_digest` and not
`same-runtime`. Gate B is the *cross-implementation* comparison §10.1 says the
tolerance exists for, and Step 13 gave it the one rule §10 had reserved for a
replay inside a single runtime.

### 9.2 Run 4's numbers, against the accepted envelope

The two differences the classification quoted:

```
oracle 143.4549368738345      oracle 1081.4363960870785
VBA    143.45493687383447     VBA    1081.4363960870783
```

Both are **inside** `rel ≤ 3e-10` by roughly six orders of magnitude — they are
one- and two-ULP differences, and §10.3's basis line says the measured worst
expression-order gap is 1 ULP with about 4,500× headroom. The accepted evidence
model anticipated exactly this class and admitted it.

Case 8 needs no second mechanism either. The digest covers **every** retained
iteration exactly; a summary statistic does not. Perturbing one retained value by
one ULP, for each of the 1000 iterations of each case:

| case | perturbations that change the digest | of which the summary layer cannot see |
|---|---|---|
| 1 | 1000 / 1000 | 996 (99.6%) |
| 6 | 1000 / 1000 | 986 (98.6%) |
| 7 | 1000 / 1000 | 984 (98.4%) |
| 8 | 1000 / 1000 | 994 (99.4%) |

### 9.3 The corrected comparison

**Exact, §10.4:** the request fingerprint, the effective seed, the iteration
count, and the RNG and method versions. Unchanged, and still by exact equality.

**Under the policy, §10.3:** every published floating row — mean, sample standard
deviation, minimum, maximum, all eleven quantiles and the deterministic base, on
both measures. Still mandatory and still every row: what changed is the rule, not
the coverage. A ladder that agreed to nine significant figures and then diverged
in the tenth is not a failure; a wrong ladder still is.

**Diagnostic, and never a criterion:** the `result_digest`. It is recorded in the
evidence with both values so a disagreement can be described.

**The numbers are not spelled in the harness.** §10.1 gave the tolerance one
owner, so `builder/pccm_builder/sim_policy.py` is that owner in code, the
emitted authority carries the policy, and the harness reads it. A `1e-N` literal
in PowerShell would be the second owner §10.1 forbade — `test_10` refuses one.

**The floor is scale-aware.** `S` is `max |contribution|` over the drivers
summed, computed from the prepared model by arithmetic over contract values, and
emitted per case as `accumulation_scale`. It is keyed to the scale that produced
the number, not to the number's own magnitude.

### 9.4 `P6-DET` decoupled

Run 4 proved this scenario's own property: both runs announced success and
**both published the same result digest**. That is the same-runtime replay
property §10.4 does keep exact, and it is not weakened by a hair — it is still
exact equality, and a blank pair can no longer pass as equal. The clause that
asked whether the repeated digest equalled the *Python* oracle's has been
withdrawn: it is a cross-language question, it belongs to `P6-ORA`, and it is
diagnostic there.

**Run-4 `P6-DET` evidence therefore establishes FIXED repeatability**, and the
red label was the stale clause, not the property.

### 9.5 What is still not known

The corrected comparison has **not run**. Whether real Excel lands inside the
accepted envelope on every row of every case is a Windows question and Run 4 did
not ask it — the old rule failed before the new one would have been reached.

Also unresolved, and deliberately not instrumented: whether the discrete subjects
§10.4 keeps exact — the MRG32k3a stream, jump state, Bernoulli decisions, draw
counts — are in fact identical. Run 4 gives indirect evidence that they are: case
8's summary statistics agreed bit-for-bit, which a diverged uniform stream could
not produce. A diagnostic to establish it directly is described in §10.5, and it
is **not** the "first ULP difference" hunt proposed before this reclassification:
that hunt assumed cross-language bit equality was required, which §10 disproves.

---

## 10. D2: the portability settlement

### 10.1 The classification

**A confirmed portability defect in the evidence artefact, not in production.**
The Windows Run-4 tree and a Linux tree at the same source generated different
bytes for `phase6_gate_b_cases.json`:

```
Windows  8C0D021FE42E10727E918D7BF099C39C46EFAE5690DD56426A5587A05E6A67E7
Linux    8C17DF7CD0EAA685151BCA683219A536D01EBDB0EDCD8BBF80993532B20B8726
```

The mechanism is not in doubt. Beta-PERT is sampled by Cheng BB/BC, which calls
`log` and `exp`; CPython delegates both to the platform libm; Windows UCRT and
glibc disagree in the last ULP. Of the four parity cases, **Beta-PERT is the only
one that reaches a transcendental** — Triangular uses `sqrt`, which IEEE-754
requires to be correctly rounded, and Uniform and Bernoulli use neither — and
Beta-PERT is exactly the case that differed.

**One number in the earlier package is corrected by this.** The previous round
recorded case 6's oracle digest as "unreconciled", suggesting a transcription
slip. It was not: `37ED4B3D7A271A52` is what the Windows tree genuinely
generated, and `7CBBB70842889648` is what Linux generates. Both are correct on
their own host. That is the defect.

**Nothing produced by the workbook is affected.** The engine performs no
approximate comparison at runtime and no published number comes from a tolerance
test. This is about how the evidence was represented.

### 10.2 The architecture

The one artefact is split by **provenance**, not by convenience.

**`build/phase6_gate_b_cases.json` — the portable case authority. REQUIRED to be
cross-platform invariant.** It identifies which parity cases are driven, their
existing Phase-5 plan-case fixtures, the seed and iteration count, the bounds and
vocabulary, the accepted comparison policy, and per case: the sampling mechanism,
the analytical-identity pointer, the accumulation scale, the comparison classes,
and the **exact discrete and identity expectations** — effective seed, iteration
count, RNG and method versions, and for the golden case the calculation and
request fingerprints. Every value is an integer, a canonical hash over exact
inputs, or arithmetic over contract decimals. It declares
`portability.cross_platform_invariant: true` and names its companion. It is
pinned by hash, and that pin is now a true statement.

**`build/phase6_gate_b_oracle_local.json` — host-local oracle evidence.
DELIBERATELY NOT cross-platform invariant, and it says so.** It carries, per
case, the oracle's floating measurements — the full summary ladder on both
measures, the deterministic base, and the `result_digest` marked diagnostic.

Provenance fields, all required by the preflight:

| Field | What it fixes |
|---|---|
| `generated_for.authority` / `.sha256` | which case authority these numbers were produced against |
| `model_version`, `sim_contract_version`, `rng_version`, `sim_method_version`, `iterations`, `supplied_seed` | the run identity the numbers describe |
| `source_revision` | the tree that produced them, from git, or a plain `unavailable` |
| `host.system` / `.release` / `.machine` / `.python_implementation` / `.python_version` / `.float_repr_style` | which host's libm these numbers belong to |
| `generated_at_utc` | when |
| `evidence_policy_authority` | the policy they are to be compared under |
| `portability.cross_platform_invariant: false` | that they must never be hash-frozen |

**Lifecycle.** Stage-A writes the authority first, hashes it, then writes the
evidence naming that hash — so the pair provably comes from one build. On the
Gate-B host that happens **before Excel starts**, and the numbers are independent
of anything VBA produces. The driver requires the file with the other artefacts
and copies it into the disposable run root, so it survives in the retained
working copy beside the log.

**Binding.** `PRE6` — `Invoke-Phase6CoveragePreflight`, the pure artefact gate
that runs before Excel is started, and NOT `P6-PRE`, which executes later inside
the live session — refuses if the file is missing, if the SHA-256
it names does not match the authority beside it, if any version disagrees, if the
authority does not claim invariance or the evidence does, if the provenance is
incomplete, or if any case has no measurements. `P6-ART` re-asserts the same
binding inside the session and records the evidence's host, source revision and
target authority into the run log.

**Why not the rejected shortcuts.** Rounding or truncating to force the two hosts
to hash alike would discard real information to protect a number. Hardcoding one
host's Beta-PERT values would make that host silently authoritative. Deleting
Beta-PERT would remove the only case that exercises Cheng. Weakening the
algorithm is not on the table. None of these was chosen, and the old SHA was not
restored: the authority now hashes to
`6a9d86784ff1f29195b23c85ee4445e133a4cb283da0c3834afe4048c495af5c`, and it is a
different file because it is a different, smaller, honest claim.

### 10.3 Which artefacts must be invariant

| Artefact | Cross-platform invariant | Pinned by hash |
|---|---|---|
| `build/phase6_gate_b_inspection.json` | **required** — addresses and names only | yes |
| `build/phase6_gate_b_cases.json` | **required** — no transcendental-derived value | yes |
| `build/phase6_gate_b_oracle_local.json` | **no, by design** | **never** — a control refuses a pin |

### 10.3a The serialisation contract, and why content invariance was not enough

Run 5's pre-Excel check found that the pinned hashes did not reproduce on a clean
Windows build at the accepted HEAD:

```
phase6_gate_b_inspection.json   raw eac55e72...  LF-normalised 83eff35f...
phase6_gate_b_cases.json        raw dee31593...  LF-normalised 6a9d8678...
```

246 and 247 line endings, and no other difference. **The content was invariant;
the bytes were not.** `Path.write_text` opens in text mode with `newline=None`,
so Python translates every newline into `os.linesep` on the way out — a no-op on
Linux, and CRLF on Windows. The emitter had never written the bytes it hashed.

The invariant artefacts are now written through `write_lf_artifact`: **UTF-8, no
BOM, LF, a retained final LF, written as bytes.** The contract is enforced at the
point of writing rather than assumed, and the function returns what it reads back
from disk, so `generated_for.sha256` describes the physical file rather than the
string the emitter happened to hold. `phase6_cases.json` carries a pinned SHA too
and had the same latent defect; it is corrected with them. The host-local
measurements use the same writer for hygiene — that is not a claim of invariance,
and they are still never pinned.

**The generated `.bas` modules deliberately keep text mode.** `AddFromString`
consumes them on Windows, where CRLF is the separator VBA expects; emitting them
as LF bytes would be a portability "fix" that broke the import it was meant to
protect.

No pinned hash moved: the LF bytes were always what Linux wrote, so the accepted
values are exactly what the corrected emitter produces. A Linux "generate then
hash" control could never have caught this, which is why `test_76` also checks
the source shape and makes the text-mode path fail if it is used at all.

### 10.4 What the transcription still shows, and what it does not

Every parity case runs twice on Linux — through the accepted Python oracle, and
through the statements `modSimRng`, `modSimSample` and `modSimEngine` write down
— and agrees bit-for-bit over all 1000 retained iterations of both measures and
over the digest (`test_55`). That says the two *algorithms* are the same
algorithm. It says nothing about whether either host's floating execution matches
the other's, and D2 shows that even the Python oracle does not match itself
across hosts. Both facts sit inside the accepted envelope.

### 10.5 The diagnostic Run 5 would need, and its honest classification

Only three questions are worth instrumenting, and none of them is "find the first
ULP difference":

* **the discrete subjects §10.4 keeps exact** — that the MRG32k3a stream, the
  jump state, the Bernoulli decisions and the draw counts are identical, because
  those admit no tolerance at all;
* **that Cheng has not branched** where the accepted policy expects the
  arithmetic path to stay fixed;
* **that no cross-language difference exceeds the accepted envelope**, which the
  corrected `P6-ORA` already answers for every published row.

**A correction to the previous proposal.** It claimed such a diagnostic merely
"observes existing state". That was wrong. The persisted publication holds
iteration **totals**; the uniforms, sampler intermediates and draw counts are not
in the workbook. Obtaining them requires **bounded transient diagnostic
instrumentation**, and it must be classified as such: a temporary observation
surface, removed before any accepted build, adding **no production authority** and
changing nothing production computes. It is not proposed here and is not
implemented here.

