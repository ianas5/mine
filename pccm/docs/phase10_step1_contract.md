# PCCM — Phase 10 Step 1: hardening and documentation contract

```text
PHASE 10 — CONTRACT SETTLED
Implementation is NOT started by this record.
Phases 7, 8 and 9 are not reopened.
```

Phase 10 is the last phase before delivery. It adds no cost, risk, percentile,
sensitivity or annual mathematics. Its whole job is to make an accepted model
**operable, protectable, recoverable, measurable, identifiable and documented**
by someone who did not build it.

This record settles that scope before any of it is written.

---

## 1. Scope

1. Methodology finalisation
2. user-facing error handling and recovery guidance
3. workbook protection and safe-edit boundaries
4. Reset Results
5. Repair Profiling
6. performance tuning and final performance evidence
7. version / build stamp
8. final documentation surfaces required before delivery

Nothing else is in Phase 10.

---

## 2. The user-operable command surface

**A delivered PCCM workbook must not require the user to invoke a macro through
`Application.Run` or the Macro dialog.** Until now every operation past the five
Phase-4 structural buttons has been driven by a harness. That is an acceptance
arrangement, not a delivery one.

**Six commands are authorised**, in the Setup command area, through the same
declared-button architecture the Phase-4 buttons already use — declared in
`structure_contract.yaml → buttons.definitions`, created by `build_stage_b.ps1`,
with `OnAction` read back to prove the binding.

| # | Caption | Entry point | New? |
|---|---|---|---|
| 1 | **Calculate** | `PCCM_Calculate` | existing endpoint, newly bound |
| 2 | **Run Simulation** | `PCCM_RunSimulation` | existing endpoint, newly bound |
| 3 | **Run Sensitivity** | `PCCM_RunSensitivity` | existing endpoint, newly bound |
| 4 | **Run Annual Cash Flow** | `PCCM_RunAnnualStochastic` | existing endpoint, newly bound |
| 5 | **Reset Results** | `PCCM_ResetResults` | **new** |
| 6 | **Repair Profiling** | `PCCM_RepairProfiling` | **new** |

The caption is user-facing and the entry point is technical; command 4 is the
clearest case and the two are deliberately allowed to differ.

**Binding a button changes no semantics.** Sensitivity and the annual step
remain **explicit separate operations**. Neither runs as part of a simulation,
and a successful run stays successful if a later analysis of it fails — which is
the accepted Phase-7 reason they were never part of a run. **No Phase-7 contract
is reopened by giving an existing endpoint a button.**

---

## 3. `Workbook_Open` — authorised

The project's **first** `ThisWorkbook` event handler is authorised, with minimal
scope: **call one protection owner** to re-apply the declared sheet protection
with `UserInterfaceOnly:=True`, because that flag does not survive a file close.

`Workbook_Open` must not: calculate, simulate, derive state, publish, run
business logic, take a destructive action, or show a dialog on success.

**Protection configuration has ONE owner.** Re-applying `Protect` at the head of
every command instead — to avoid the event — is explicitly rejected: it would put
the same rule in a dozen places and make its correctness a matter of
remembering.

**Failure-safe.** An open handler that throws must not leave the application
altered. On any failure it restores `ScreenUpdating`, `Calculation` and
`EnableEvents`, discloses the failure once, and leaves the workbook usable and
unprotected rather than half-protected. This is a required Windows scenario.

---

## 4. Reset Results

**Clears every result publication:**

- deterministic calculation publication and last-successful fingerprint;
- calculation last-attempt presentation and history;
- simulation publications, both banks, and the active publication identity;
- simulation last-attempt presentation and history where applicable;
- annual distribution and profile publications and their stamps;
- sensitivity publication.

**Preserves every input and every identity continuity:**

- Setup, Config, Cost Lines, Risk Register, Inflation, Cost Profiling, Risk
  Profiling;
- the applied timeline inputs;
- **permanent-ID counters** — resetting them would reissue identities;
- **run-id / nonce monotonic identity state** — it is an anti-replay guarantee,
  and a discarded run must never be re-creatable by a future one;
- seed history and identity fields required for non-collision;
- build metadata.

**Live states immediately afterwards are DERIVED, never written:**

| Axis | After reset |
|---|---|
| calculation, valid model | `NOT CALCULATED` |
| calculation, invalid model | `INVALID` |
| simulation | blank — no publication exists |
| annual | `NOT PRODUCED` |
| sensitivity | not produced |

**No state string is forced to obtain these outcomes.** They fall out of the
existing owners once the publications are gone; forcing one would make Reset a
second opinion about state, which §13 forbids.

The command is destructive-confirmed (`ConfirmDestructiveChange`), deterministic,
idempotent, input-preserving, and **reversible only by re-running** the relevant
operations. No output is reconstructed.

### 4.1 What "preserves inputs" means for acceptance

Every **declared editable / business input** and every **permanent identity
value** compares exactly before and after.

**Not** a binary workbook-byte comparison. Calculation metadata, volatile
presentation and recalculation artefacts may legitimately move across a reset,
and a byte comparison would fail on those while proving nothing about the inputs
it exists to protect.

---

## 5. Repair Profiling

**The problem.** The profiling grids are keyed by driver permanent id. Add and
delete operations, or a structural mishap, can leave a profiling row missing,
orphaned or misaligned. The checker *reports and refuses* — by design, it "never
repairs" — so today the user has no route back except by hand.

| Class | Condition | Action |
|---|---|---|
| repairable | a valid driver has **no** profiling row | create it, id only, **weights blank** — a blank is an unmade assumption, not a zero |
| repairable | a profiling row references a driver that **no longer exists** | remove the orphan |
| repairable | rows present and unique but **out of register order** | reorder; **weights travel with their id** |
| repairable | grid narrower or wider than the applied duration | extend with blanks; trim **only** empty columns beyond the duration |
| **ambiguous — refuse** | two rows carry the same id with **different** weights | refuse, naming the id |
| **ambiguous — refuse** | trimming would discard a **non-empty** weight | refuse, naming id and column |
| **ambiguous — refuse** | a row's id is unreadable | refuse |
| **ambiguous — refuse** | weights present but do not sum to 1 | refuse — a business assumption, not a structural defect |

**Preserved:** every semantically attributable user weight.
**Regenerated:** row existence, ordering, the id column, grid width, layout.
**Never changed:** a valid weight, anywhere.

### 5.1 A structural repair is not a semantic change

A repair that only restores alignment or order **while preserving the
driver-to-weight mapping has changed no model input**, and must not pretend
otherwise.

- **Do not force `STALE`.**
- The existing fingerprint owner decides, as it does for everything else.
- If semantic inputs are unchanged, **the calculation fingerprint must remain
  identical** — and that is a control, not a hope.
- If a repair necessarily changes a semantic input, the ordinary existing
  fingerprint and state rules apply, with no special case.

---

## 6. Protection

| Sheet | Role | Protection |
|---|---|---|
| Setup | input | protected; the declared editable inputs unlocked; the applied-timeline block stays locked |
| Config | admin | protected; list masters unlocked; locked model constants stay locked |
| Cost Lines, Risk Register | input | protected; table body unlocked; id column, headers, structure locked |
| Inflation, Cost Profiling, Risk Profiling | input | protected; rate and weight cells unlocked; id column and layout locked |
| Dashboard, Results, Sensitivity, Model Check | output | protected, nothing unlocked |
| Methodology | reference | protected, nothing unlocked |
| `_Calc` | technical | stays `hidden` — auditable by design — and protected |
| `_SimData` | technical | stays `veryHidden` and protected |

- **Workbook structure: protected.** The sheet order is an architectural lock the
  build already enforces; leaving structure unprotected lets a user break at run
  time what the build refuses to break.
- **Passwordless.** It stops accident, which is the entire goal. A sheet password
  is not a security control, and one kept in source would be worse than none.
- **`UserInterfaceOnly:=True`** so macros still write, re-applied by
  `Workbook_Open` (§3).
- **Maintenance stays possible**: every `Protect` call is generated from one
  declared table, and an unprotect path exists for the maintainer.

> **Passwordless protection is an accidental-edit control. It is not a security
> claim and must never be described as one.**

---

## 7. Methodology

Eleven sections replace the empty *"Methodology Sections"* placeholder. The
audience is a cost engineer or a reviewer, never a developer.

| § | Section |
|---|---|
| M1 | deterministic calculation basis, carry convention, reporting currency |
| M2 | three-point estimating — Triangular / PERT / Uniform, ordering, central basis |
| M3 | risk occurrence × severity |
| M4 | FX, inflation and discounting — rates are inputs, never derived |
| M5 | cost profiling — weights distribute, they never scale |
| M6 | Monte Carlo simulation — sampling, iterations, AUTO / FIXED seed modes |
| M7 | percentile interpretation and contingency |
| M8 | annual cash flow and the selected-Px profile — see 7.1 |
| M9 | sensitivity and Spearman interpretation |
| M10 | zero-variance drivers — undefined, not zero |
| M11 | `CURRENT` / `STALE` / `INVALID` / `NOT CALCULATED`, and live versus persisted |

Methodology explains. It never becomes a second opinion about state: M11
describes the vocabulary and defers the current answer to Model Check.

### 7.1 M8 — the corrected wording

Two different things, which must not be conflated. The Step-1 draft blurred them
and is corrected here.

**(a) Annual simulated distributions.** The percentile ladder of each project
year, taken across iterations. Adding up a per-year Px across the years does
**not** give the reported total Px, and such a set of year values **is not a
profile of anything** — each year's figure comes from a different iteration.

**(b) The selected-Px annual profile.** The annual vector belonging to the
reported total-cost percentile. Because a reported Px is generally a blend of two
order statistics, the profile applies the **same** lo / hi / f the reported total
used to the annual vectors of those exact two iterations:

```text
Profile_Px(y) = (1 - f) * AnnualVector_lo(y) + f * AnnualVector_hi(y)
```

**This profile DOES sum to the selected total Px** — `Sum_y = reported Px`, by
linearity, subject only to floating-point round-off, and it degenerates to a
single iteration's vector when `f = 0`.

> Methodology must **not** say "the annual profile does not sum to the total Px".
> That contradicts the accepted selected-profile identity. It must say which of
> the two objects is being discussed, every time.

---

## 8. Error handling and user recovery

No new state machine. Every row is an existing owner's existing outcome, given a
user-facing surface.

| Failure class | Refuses / warns | User sees | Where | Recovery | Results |
|---|---|---|---|---|---|
| invalid cost-line / risk input | refuses | ERROR, Subject = permanent id, live reason | Model Check + command message | correct the named driver, Calculate | prior publication retained, marked; nothing overwritten |
| broken three-point ordering | refuses | as above | as above | as above | as above |
| missing FX | refuses | ERROR, **blank Subject** — model-wide | as above | add the currency to the reference set | as above |
| missing / invalid profiling | refuses | ERROR, Subject where the owner holds one | as above | fix the weights, or **Repair Profiling** | as above |
| invalid timeline | refuses | ERROR, blank Subject | as above | Apply / Update Timeline | as above |
| invalid probability / quantity | refuses | ERROR, Subject = id | as above | correct the named driver | as above |
| simulation request invalid, calculation `CURRENT` | refuses the run | ERROR — an independent fault | as above | correct iterations or seed | prior publication retained |
| simulation `INVALID`, calculation not `CURRENT` | context | **INFO**, not counted | Model Check | fix the calculation row | — |
| calculation `STALE` | warns | WARNING | Model Check | press Calculate | last publication usable, historical |
| simulation `STALE` | warns | WARNING | Model Check | re-run the simulation | published run usable, not current |
| annual `HISTORICAL` / `OTHER Px` | warns | WARNING as the owner reports it | Model Check, Results | re-run the annual step | annual output historical, retained |
| sensitivity unavailable | context | **INFO** — an optional output | Model Check, Sensitivity | run sensitivity | absent, not defective |
| structural fault | refuses at the structural gate | ERROR, one row per fault | Model Check | repair the named structure | calculation blocked |
| unexpected runtime failure | refuses, restores, discloses | operation, reason, restore note | modal | retry; report the message | see 8.1 |

**One vocabulary, one owner.** `modAppState` already owns `ReportResult`,
`ReportFailure`, `ConfirmStructuralChange` and `ConfirmDestructiveChange`, with
automation hooks so a harness drives them without a human. Phase 10 reuses it and
adds no dialog mechanism.

**A gap Phase 10 closes.** `PCCM_Calculate` and `PCCM_RunSimulation` currently
tell the user nothing — no module outside `modAppState` contains a `MsgBox`. They
persist a status and a detail and return. Their outcome is routed through
`ReportResult` / `ReportFailure`. That is presentation only and changes no state.

### 8.1 The transactional-safety claim, stated precisely

An unverifiable global claim that no legacy VBA path can ever partially write is
**not** made. The required claim is:

- every **new** Phase-10 destructive or structural command is transactionally
  safe;
- **shared existing** command owners retain their already-accepted restoration
  contracts, unchanged;
- application state — `ScreenUpdating`, `Calculation`, `EnableEvents`, and
  protection — is restored on failure;
- **no newly introduced partial-write path is permitted.**

---

## 9. Version, build stamp and release

**One surface: the Methodology `Build Metadata` block.** It already exists and is
generated, never typed.

| Row | Owner | Phase 10 |
|---|---|---|
| PCCM Model Version | `workbook.yaml → model.model_version` | → **1.0.0** |
| Build Phase / Release | `workbook.yaml → model.build_phase` | see 9.2 |
| Builder Version | `BUILDER_VERSION` | → **1.0.0** |
| Build Timestamp (UTC) | `resolve_build_timestamp()`, overridable by `PCCM_BUILD_TIMESTAMP` | unchanged |
| Source Manifest / Input Contract / Driver Contract / Structure Contract versions | their own spec files | unchanged |
| **Source Revision** | — | **ADD**: short git hash plus a clean/dirty marker, captured at Stage-A build |

### 9.1 1.0.0, and two authorities that stay independent

For this first production delivery **both** reach `1.0.0`. That coincidence is
this release only.

> **Model version and builder version remain INDEPENDENT authorities after
> release.** A later builder-only change need not move the model version, and a
> later model contract release need not mechanically equal the builder version.

No version literal is duplicated. The Dashboard may mirror the block; it must not
restate it.

### 9.2 The build-phase string

The user-visible value today reads `"Phase 5 - Calculation Workspace (Gate A:
source)"`, which is stale with Phase 9 closed. This is a **Phase-10 metadata
finalisation item and reopens no earlier phase.** During implementation it may
identify Phase 10; the **final delivery artifact must present a release-oriented
value, preferably `Release 1.0 — Production`.** One generated owner; no competing
phase or release label is typed anywhere else.

---

## 10. Performance

**Measured, separately:** **A** Calculate · **B** Monte Carlo simulation ·
**C** Sensitivity · **D** annual replay and reporting · **E** workbook
recalculation and UI responsiveness.

**Models:** Small = 20 drivers / 10 years · Medium = 100 / 25 · Large = 300 / 40.

### 10.1 The practical matrix — not a Cartesian product

| Model | 10,000 | 50,000 | 100,000 |
|---|---|---|---|
| **Small** | required | required | required where the operation meaningfully depends on N |
| **Medium** | required | required | where practical; **may be omitted for Sensitivity** if runtime evidence shows it adds little delivery value |
| **Large** | required | **maximum** for Simulation, Sensitivity and annual replay | **not required** |

**Calculate and ordinary workbook recalculation are iteration-independent** and
are measured once per model size, not repeated for every N.

> Every omitted combination is recorded explicitly as **not required /
> impractical**, never left to read as missing evidence.

### 10.2 What every timing must carry

Excel version and bitness · Windows version · CPU · RAM · local versus synced
path · whether other workbooks are open.

**Cold** = the first operation after file open, reported separately. **Warm** =
the median of three subsequent runs. Only warm figures are compared across
builds. Timings are **not user-visible**: they go to the acceptance report, not
to a sheet.

### 10.3 Baseline and regression

**No absolute pass/fail threshold is contracted before the baseline exists.**
The Phase-10 benchmark creates the first delivery-oriented measured baseline.
After it:

- **1.5×** a comparable warm baseline → investigate;
- **2×** a comparable warm baseline → **blocks delivery** unless explained and
  accepted.

### 10.4 The Phase-7 timings are historical

`docs/phase7_closure.md` records 20×10k ≈ 6.4–7.6 s, 100×10k ≈ 31.4 s, 300×10k ≈
105.4 s as **the operator's report, explicitly not confirmed by the repository** —
only the scenario shape in `phase7_timing_scenarios.ps1` is re-establishable.
**That record is quoted, never edited.** Phase 10 adds new, separately labelled
end-user benchmarks whose numbers its runner captures into its report.

---

## 11. Distribution readiness

### 11.1 Verified, expected and unsupported

**REQUIRED / OFFICIALLY SUPPORTED**

- Windows desktop Excel;
- macros enabled;
- **the exact Excel version and bitness family actually exercised by the final
  Windows acceptance run** — named in the manual once that run exists, and not
  before.

**COMPATIBLE BY INSPECTION, NOT INDEPENDENTLY VALIDATED**

- **Office bitness.** No `Declare` statement exists anywhere in `src/vba`, so no
  Win32 pointer surface is used. 32-bit Office is *expected* to be compatible
  where no capacity limit is exceeded. It is **not claimed as accepted unless it
  is actually tested.**
- **Minimum desktop Excel release.** The built workbook uses 23 worksheet
  functions — `ABS AND CHAR COUNTIF COUNTIFS EXACT FIND IF IFERROR INDEX LEN
  MATCH MAX MID NA NOW OR SUBSTITUTE SUM SUMIF SUMPRODUCT T TRIM` — of which the
  newest, `IFERROR` and `COUNTIFS`, are Excel 2007. **The formula surface
  therefore imposes no requirement beyond Excel 2007.** That is an inspection
  result about formulas, not a support claim; the supported minimum is whatever
  acceptance actually runs.

**UNSUPPORTED**

- Excel Online; Excel mobile; LibreOffice; any environment where VBA cannot
  execute.

**Excel for Mac: unsupported / unvalidated.** It is not called supported unless a
final compatibility audit proves every Windows-specific dependency absent **and**
it receives suitable execution evidence.

> The final manual states **VERIFIED** support separately from **EXPECTED**
> compatibility. The two are never merged into one list.

### 11.2 Packaging and portability, from source

| Check | Finding |
|---|---|
| local-path dependence | none — no `ThisWorkbook.Path`, no UNC, no URL in any module |
| external links | none in VBA; the built workbook must be asserted link-free as a control |
| packaging | `.xlsm`, `FileFormat 52`, enforced and read back by `build_stage_b.ps1` |
| macro trust | required — the workbook is inert without it |
| references / add-ins | none beyond the default project references |
| hidden sheets | `_Calc` hidden and auditable; `_SimData` veryHidden |
| regional settings | already handled — `HostDecimalSeparator()` derives the host separator at runtime |
| copy and rename | must be proved safe; nothing keys off the file name |

---

## 12. Documentation deliverables

A standalone **25-section user manual** — purpose; sheet-by-sheet guide; initial
setup; timeline; cost lines; risks; inflation; cost profiling; risk profiling;
deterministic Calculate; simulation; annual outputs; sensitivity; Dashboard;
Model Check; the state concepts; common errors; recovery; Reset Results; Repair
Profiling; performance expectations; protection and safe editing; version
identification; troubleshooting; delivery and distribution notes.

Plus a **one-page Quick Start** and a **~15-minute demonstration / rehearsal
script**.

The manual is written **after** implementation: its Reset, Repair and protection
sections describe surfaces that do not exist yet.

---

## 13. No new business logic

Hardening must not silently change **costing mathematics, risk mathematics,
percentile semantics, sensitivity semantics, annual profile identity,
fingerprints, existing state contracts or the Model Check taxonomy.**

| Guard | Exposure | Verdict |
|---|---|---|
| costing, risk, percentile, sensitivity mathematics | none — no scope item touches a kernel | safe |
| annual profile identity | Reset **clears** annual output | safe — clearing is not redefining |
| calculation fingerprint | Repair may change grid shape | safe **and visible** — §5.1: unchanged semantics means an unchanged fingerprint, and that is a control |
| simulation request fingerprint | none | safe |
| existing state contracts | Reset clears **publications**; live states stay derived | safe |
| Model Check taxonomy | no new check id, group, severity or vocabulary word | safe |

**Shared owners Phase 10 will touch, declared in advance:** `modAppState` (reuse
only), `ThisWorkbook` (new, §3), the profiling structure owner, and the button /
entry-point tables in `structure_contract.yaml`. Each is declared in the Phase-7
§1.1 and Phase-8 production tables under the existing discipline.

> **Any correction discovered in a shared owner stops implementation and is
> surfaced before proceeding.**

---

## 14. Acceptance matrix

| | Scenario | Static | Windows |
|---|---|---|---|
| A | protection allows every intended input | declared-vs-unlocked table | write each input cell |
| B | protection blocks unintended structural edits | — | locked cell, row insert, sheet delete all refused |
| C | macros operate under protection | `UserInterfaceOnly` + re-apply on open | full Calculate / Simulate cycle while protected |
| D | Reset clears outputs, preserves inputs | — | every declared input and identity value compares exactly (§4.1) |
| E | Reset produces correct live and persisted states | evaluator | `NOT CALCULATED` / `INVALID`, blank simulation, `NOT PRODUCED` |
| F | Repair repairs a repairable defect | fixture matrix | delete one profiling row; repair; weights preserved |
| G | Repair refuses an ambiguous defect | fixture matrix | duplicate id, differing weights → refusal names the id |
| G2 | a structural-only repair moves no fingerprint | fingerprint control | fingerprint identical before and after (§5.1) |
| H | Methodology and version surface render | block and emitter controls | eleven sections present; metadata rows non-blank; revision matches |
| I | opens and recalculates cleanly after protection | — | open, recalculate, **no persisted cell rewritten** |
| J | a distribution copy runs from a different path and name | — | copy, rename, run |
| K | no external or local path dependency | link scan on the built file | — |
| L | performance benchmarks | scenario shape | §10.1 matrix |
| M | an unexpected runtime error leaves state clean | injected-failure control | app state and protection restored; no partial write |
| N | **the six commands are reachable as buttons** | button and entry-point tables | each button present, `OnAction` correct, each invokable |
| O | **`Workbook_Open` failure is safe** | — | injected failure leaves the application state clean and the workbook usable |

**One new runner**, the P8-Z / P9-1 shape, plus one performance runner. **The
Phase-7, Phase-8 and Phase-9 suites are not re-run** unless implementation
touches a shared owner — which §13 says it must not.

---

## 15. Expected implementation surfaces

**Spec** — `workbook.yaml` (Methodology sections, release label, protection
table, source-revision row); `structure_contract.yaml` (six buttons, two new
entry points, protection declaration).
**Builder** — `workbook_builder.py` (Methodology emitter, source revision); a
protection emitter.
**VBA** — `modReset`, `modRepair`, a protection owner, the `ThisWorkbook` open
handler, and outcome routing in `modCalcReport` / `modSimReport`.
**Bootstrap** — `build_stage_b.ps1` (six buttons, protection); a Phase-10
hardening runner; a performance runner.
**Docs** — this record; the user manual; the quick start; the demo script.
**Tests** — `test_phase10_*` and a runner-source suite per new runner.

---

## 16. Status

```text
PHASE 10 STEP 1 — CONTRACT SETTLED
Implementation is NOT started by this record.
No Windows run has been executed for Phase 10.
```
