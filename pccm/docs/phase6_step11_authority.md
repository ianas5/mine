# PCCM Phase 6 — Step-11A authority record

**Publication / Results / orchestration authority closure. No implementation.**

```
spec/sim_contract.yaml                  two-bank publication, transaction order,
                                        persisted statistics, the bridge, the
                                        settled public surface
spec/workbook.yaml                      the physical Results / _SimData shell
builder/pccm_builder/sim_loader.py      all of it, enforced, and cross-validated
builder/pccm_builder/sim_emit.py        the coordinates projected
builder/pccm_builder/sim_cases.py       the publication-state vectors
builder/pccm_builder/spec_loader.py     the shell parsed
builder/pccm_builder/workbook_builder.py  the shell materialised
builder/pccm_builder/verify.py          the shell's formula cells enumerated
docs/phase6_step11_authority.md         this record
```

**Not in this round.** `src/vba/modSimReport.bas` does not exist and is not in
the registry. There is no `PCCM_RunSimulation`, no `CalcPrepareSimulationInputs`
implementation and no VBA of any kind. **Step 11 implementation remains
unauthorised. No Windows or Excel runtime ran.**

---

## 1. Authority chain

| Role | Commit |
|---|---|
| Accepted Step-9 statistics authority | `e760c50361f03bce4a393de64614b1cac45d7d29` |
| Accepted Step-10A request-fingerprint grammar | `34a7c467a2e22c3f896cdc10487a1b3922b4536b` |
| Accepted Step-10 implementation head | `09a56b2405614d55f5dc4e7028398bdef10b0afb` |
| Step-11A publication closure | this commit |

---

## 2. The three gaps, closed

### 2.1 `prior_successful_publication_survives` was aspirational

The contract asserted it; the physical design could not deliver it. Writing a new
distribution over the published rows and only then stamping the Run ID cannot
survive a COM failure half way through a million rows, and no rollback of a
million rows is a transaction anybody should attempt.

**Two banks, one active.** A candidate success is written **entirely into the
inactive bank**, verified there, and published by one small final write that
moves `active_bank`. A failure at any earlier point leaves the active bank
physically untouched and the corrupted candidate with **no semantic standing at
all** — nobody reads it, because `active_bank` still names the other one.

```
candidate_target:   ""  -> A       first success
                    A   -> B
                    B   -> A
```

No third bank. No temporary worksheet. No duplicate workbook. **The second bank
consumes COLUMNS, never rows** — both banks share the row axis, so `H` is still
33, the first iteration row is still 34, and the technical ceiling is still
1,048,543. `active_bank` took the row-30 **spacer**; the axis did not grow.

The transaction order is locked end to end, and the switch is the last step of
the last write:

```
prepare_phase5_inputs_and_require_current
validate_pre_allocation_prerequisites
allocate_auto_nonce_when_auto
run_simulation_and_statistics_in_memory
build_request_fingerprint_and_result_digest_in_memory
choose_inactive_bank
write_candidate_snapshot_to_inactive_bank
write_candidate_summary_to_inactive_bank
write_candidate_contingency_to_inactive_bank
write_candidate_iterations_to_inactive_bank
verify_inactive_bank_against_staged_package
final_commit_shared_block_including_active_bank      <- D22:D30, one write
```

`D22:D30` is exactly the nine **shared** rows in row order, ending with
`active_bank`. The prior block is captured before the write and restored if it
fails, so a failed commit leaves the run id and the published bank exactly where
they were.

**Run ID is allocated by the commit and by nothing else.** The candidate is
`last_run_id + 1`, held locally, so a refused or failed attempt burns no
identity. Headroom is checked **before** the AUTO nonce is consumed: there is no
reason to burn a random sequence for a run that can never be identified or
committed.

### 2.2 `Results derives from _SimData` was not implementable

`_SimData` persisted the raw iterations but not one statistic, while
`modSimStats` is the single owner of those numbers. Recomputing them on the sheet
would be a second statistics engine; recomputing them in VBA on every read would
make Results a calculator.

**So the run persists them.** A banked summary block carries the mean, the sample
deviation, the minimum, all eleven rungs, the maximum and the deterministic base
A, for nominal and PV, produced by `modSimStats` from the **same retained arrays**
that are written to that candidate iteration bank.

**And the whole contingency ladder is precomputed.** Selected Confidence Level is
reporting-only and may move without a rerun, so a publication holding only the
selected rung would force either a rerun or a worksheet subtraction the moment it
did. Every rung's contingency goes through `SimStatsContingency` before the
candidate bank may commit, and if **any** of the twenty-two values is
unrepresentable the publication is refused and the active bank does not move. The
fixed rung is persisted too, so the ladder is structurally uniform.

**No rung label is written into `sim_contract.yaml`.** The selectable ladder
belongs to `input_contract.yaml`; a rung is identified here by its projected key
and its label is resolved through the accepted ladder authority. The first draft
of this block spelled the labels out, and the accepted `test_63` caught it —
which is exactly what that test is for.

### 2.3 Results had no schema, and would have been a second write

Results is now **presentation**: materialised once by Stage A, part of no
transaction, computing nothing. Every cell is a lookup into the **active** bank
or the live selector. That also removes the failure mode where `_SimData` commits
and a second Results write then fails, leaving a published distribution the
workbook does not show.

The locked minimum layout: Run Stamp rows 8–24, Summary Statistics rows 26–44,
the three selected reporting rows 46–48, and the deferred Annual Cash Flow and
Reconciliation placeholders at 51–52 and 54–55.

---

## 3. Every Results formula is a lookup, and that is enforced

`workbook.yaml` owns **where** each label and formula sits; `sim_contract.yaml`
owns **which** fields exist and what they mean. The simulation loader
cross-validates the two, so a row that names a field the contract does not have,
omits one it does, labels one differently, or reads the wrong `_SimData` row is a
**build failure**.

The same cross-check enforces the property the design rests on:

* no formula may call `AVERAGE`, `STDEV*`, `PERCENTILE*`, `QUARTILE`, `MEDIAN`,
  `OFFSET`, `INDIRECT`, `RAND*`, `NOW` or `TODAY`;
* no formula may subtract — contingency is persisted, never derived on the sheet;
* **every banked formula must read the active-bank selector**, and every shared
  one must not;
* each selected reporting row must `MATCH` the selector inside the persisted
  ladder and `INDEX` the answer out of it, guarded by a `COUNTIF` over the
  **selectable** sub-range so the fixed rung is not selectable and an unknown
  selector displays blank.

An invalid selector blanks those three rows and does **nothing else**: the
simulation stays CURRENT, no rerun is required, and no `UNSELECTED` state exists.

---

## 4. One Phase-5 bridge, and a settled public surface

`CalcPrepareSimulationInputs` is locked as the one reusable surface into the
accepted Phase-5 preparation — owner `modCalcReport`, reusing the existing
private `PrepareCurrentCalculation`, requiring Phase 5 **CURRENT**, writing
nothing to `_Calc`, updating no Phase-5 metadata, duplicating no factor
mathematics, succeeding on a zero-driver model, and returning the seven prepared
values. It is an **internal** cross-module API: not an endpoint, and its name does
not begin with `PCCM_`.

The contract also records the obligation Step 10 carried forward: the analytical
fingerprint it returns is the **current** one, never a stored last-successful
value.

`read_accessor_names_settled` moves from `false` to `true`, with an exact closed
list — chosen before the module exists precisely so the implementation cannot
invent a name:

```
PCCM_RunSimulation                          (the endpoint, not an accessor)
PCCM_SimulationStatus
PCCM_SimulationRequestFingerprint           the STORED one, from the active bank
PCCM_CurrentSimulationRequestFingerprint    the RECOMPUTED one
PCCM_SimulationResultDigest
PCCM_SimulationAttemptResult
PCCM_SimulationAttemptDetail
```

The two fingerprint accessors' semantics are locked **word for word**, for the
same reason the three simulation-state definitions are: one reads a stored value
and the other recomputes one, and a prose drift between them is the difference
between a correct staleness answer and a wrong one.

---

## 5. Verification

### 5.1 Python suite

```
2954 passed, 0 failed          (1394.24s)
2954 collected
```

| Count | What |
|---|---|
| 49 | Step-11A publication tests — 27 authority and shell, 22 mutation controls |
| 16 | `J_publication` corpus cases, all EXACT |
| 112 | corpus cases in 12 groups (was 96 in 11) |
| 351 | Stage-A builder/verifier checks, 0 failed |

Collection moved from **2905** to **2954**: +49, the whole Step-11A test file**. **No test was deleted,
skipped or weakened.** Four accepted expectations legitimately changed and are
named here: `test_62` (the accessor names are now settled), `test_66` (the run
identity gained `active_bank` at the spent row-30 spacer, and the note column
moved to H to make room for bank B), and `test_78` / `test_84` of the Stage-A
suite (12 groups, 82 EXACT of 112). `test_84` additionally now asserts that all
sixteen added cases are EXACT.

**Every control was checked for vacuity.** One of them survived its first draft:
the "Results pins one bank" control passed because the cross-validator only
counted how many formulas mentioned the selector. It now requires **each** banked
formula to read it by name, and each shared one not to.

### 5.2 The row ceiling did not move

```
reserved rows H         33          unchanged
first iteration row     34          unchanged
max iterations          1048543     unchanged
```

`active_bank` took the row-30 spacer; both iteration banks share rows 34 upward
and differ only in columns (B/C/D against F/G/H). A control plants a second bank
that buys itself a header row and requires refusal.

### 5.3 Artefact movement

| Artefact | SHA-256 | Status |
|---|---|---|
| `build/PCCM_stageA.xlsx` | `88be5508f00e41f47b14e1f304f12ddd926a73b398db15da97be62b598ac4980` | **changed** — Results / `_SimData` shell only |
| `build/vba/modSimContract.bas` | `1d949be659d0afc3e18501a34b7d372bab3df575fc1a981cfd60dcf1f293a753` | **changed** — new coordinates projected |
| `build/phase6_cases.json` | `98f835375f5b8f548172c21ae6102b50fef7e6a001e196ece0741c987d78b6d1` | **changed** — `J_publication` added |
| `build/stage_b_manifest.json` | `01ca01a80598256f6ada218603032cd4be6c9bb9b86f452fb701dc610172ae57` | **byte-identical** |
| `src/vba/modSimRng.bas` | `3d7c2cb365df03ccf73722f39b0c10e8964381e7cdd243732381dac7638257e3` | byte-identical |
| `src/vba/modSimSample.bas` | `5553198289bd98a7c84025868ac03c9f8ec95da3c01b23249c0da57d77901877` | byte-identical |
| `src/vba/modSimEngine.bas` | `f1283fe7d5d2ffcc5345dab9a00f68d3685b787563d104f50a886c5ed409abab` | byte-identical |
| `src/vba/modSimStats.bas` | `98bd21b227047d04e6847e554e027b339cf01dfb1112c1539a9e334966233be0` | byte-identical |
| `src/vba/modSimFingerprint.bas` | `9e6ad972fe59ead9e34c7d65b807dd0f2ca1cb1b29bfa71b377a4eb8f65cdfda` | byte-identical |
| `src/vba/modCalcFingerprint.bas` | `2efbb30c6f915c04b9c07adec07e25e11f4b5bd2b98e3efa818631dc510ce847` | byte-identical |
| `build/vba/modConstants.bas` | `5d10d0074478d0dfaccf329d161333c75833aa78612b3a3e925ff519be227425` | unchanged |
| `build/vba/modCalcContract.bas` | `251cac5e1bcbd67461126529fe133291d651ac7be4c42d804c10b77fa1b8c6b9` | unchanged |
| `build/phase4_scenarios.json` | `219b255205bf470037f1dbe71106d238e1f52acc1dcf6266ded125617acb9c04` | unchanged |
| `build/phase5_cases.json` | `f79d64efcc1795d7686cb877c2f824e8c6e9827f403ceb63f024552e888c98e7` | unchanged |
| `build/phase5_gate_b_inspection.json` | `9cc1b007faec52839aa87c4d06be827aa182b88a43a41cf0f483cde2b9004e8f` | unchanged |

The manifest is **byte-identical**: no registry entry and no forbidden-rule moved.
The registry is the same twenty-one modules Step 10 closed with, and the
structured D6-11 projection is unmoved — `MRG32k3a → [modSimRng]`,
`RunSimulation → []`, `Percentile → []`.

### 5.4 The workbook gained a shell, not a result

The materialised `_SimData` carries every label, both iteration-bank headers, the
summary and contingency labels, and the three initial values `0`, `0`, `NONE`.
It carries **no** iteration row, **no** snapshot value, **no** request
fingerprint, **no** result digest, and `active_bank` is blank. Results carries
its labels and its lookup formulas and nothing else. `test_25` scans the whole
sheet for `PCCM-FP`, `PCCM-RD`, `CURRENT`, `STALE` and `SUCCESS` and requires
none of them.

---

## 6. Carried forward

**GATE-B TEMP-DIR CLEANUP DEBT — OPEN.** Not addressed, not touched.

**Also deferred to Gate B, unchanged:** the two stale descriptive `"15"` strings;
the raising arm of `SimStatsLadderExtent`; the raising arm of
`SimFpRetainedExtent`; `CalcFpContinueDigest` and `AscW` on real Excel.

**Carried to the Step-11 implementation:**

1. `modSimReport` must obtain the **current** analytical fingerprint through
   `CalcPrepareSimulationInputs` and pass THAT to
   `SimFpBuildRequestFingerprint` — never the stored last-successful value.
2. The result digest must come from the exact retained arrays being published.
3. Ladder provenance and non-mutation between `SimStatsDescribe` and selection.
4. Nominal and PV staged locally and committed together.
5. The bridge's source tests must prove it contains no factor mathematics and
   only projects fields from the accepted existing preparation.

---

## 7. Step-11A acceptance gate — self-check

| Gate condition | Status |
|---|---|
| two banks and one explicit active-bank authority | `test_01`, `test_02` |
| a candidate never overwrites the active bank | `test_02`, `test_28` |
| the active-bank flip is the final success commit | `test_03`, `test_04`, `test_31`, `test_32` |
| the row ceiling is unchanged | `test_07`, `test_08`, `test_48` |
| the full Step-9 summary is persisted | `test_09`, `test_44` |
| the full contingency ladder is persisted | `test_11`, `test_45` |
| Results contains no statistics or contingency arithmetic | `test_14`, `test_37`, `test_38` |
| Results selects only from persisted active-bank values | `test_15`, `test_16`, `test_39` |
| Selected CL stays reporting-only | `test_17`, `test_43` |
| the exact minimum Results layout is locked | `test_21`, `test_40`, `test_41` |
| one locked Phase-5 bridge surface | `test_18`, `test_46` |
| the public accessor names are settled | `test_19`, `test_47` |
| Stage A contains the complete EMPTY publication shell | `test_22`–`test_25` |
| no source VBA Step-11 implementation exists | `test_20` |
| D6-11 and the registry unchanged | §5.3 |
| no Windows/Excel runtime ran | none was authorised and none was run |
| Gate-B debt | **OPEN**, §6 |
