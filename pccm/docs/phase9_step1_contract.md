# PCCM — Phase 9 Step 1: the Model Check contract

```text
CONTRACT / DESIGN ONLY.
P9-2 has NOT started. No Phase-9 production, spec or builder change exists.
```

Phase 9 finalises the **Model Check** sheet: the aggregation and presentation of
warnings and blocking issues, and the advisory when Monte Carlo iterations are
below 10,000.

Model Check is a **presentation of facts other modules already own**. It is not a
new state machine, and it creates no refusal.

---

## 1. Authoritative owners

| Fact | Owner | Model Check |
|---|---|---|
| Structural / input validation | `modStructuralCheck.ValidateStructure`, via `PCCM_StructuralReport` | **aggregates** — one row per fault |
| Calculation state | `modCalcReport.DeriveStatus` (private, pure) | **mirrors** — through the adapter authorised in §8 |
| Calculation attempt | `PCCM_CalculationAttemptResult` / `…Detail` | **mirrors** (INFO) |
| Simulation state (live) | `modSimReport.DeriveSimStatus` → `SimReportDerivedStatus` → `PCCM_ResultsSimulationState` | **mirrors** |
| Simulation publication | run stamp, active bank, `PCCM_SimulationRequestFingerprint` | **mirrors** (INFO) |
| Annual distribution / profile / Px / years | `modSimAnnualStore.PCCM_AnnualDistributionState` and its three siblings, on the read path → the four `modResultsState` adapters | **mirrors** |
| Sensitivity availability | the Sensitivity sheet's own `availability_formula` | **mirrors** — the whole sentence |
| Iteration recommendation | `input_contract.yaml → inputs.monte_carlo_iterations` | **reads** — §9 |

**Not displayed:** fingerprints, result digests, seeds, nonce lifecycle, bank
identity, iteration-level data. Machine identity is not user-facing validation.

### 1.1 The two axes stay apart

The calculation axis is `NOT CALCULATED / CURRENT / STALE / INVALID`. The
simulation axis is `CURRENT / STALE / INVALID`, with a **blank** meaning no
publication exists. They are separate rows with separate vocabularies. P8-3 cost
a Windows run to the assumption that they were one axis; this contract does not
repeat it.

---

## 2. Severity taxonomy

The manifest already reserves *"the aggregate PASS / WARNING / ERROR status"*.
That vocabulary is adopted, not replaced.

| Severity | Meaning |
|---|---|
| **ERROR** | the requested operation is genuinely invalid or impossible **under an existing contract** |
| **WARNING** | the model is usable; something requires an operator decision or action |
| **INFO** | context. Never a judgement, never counted |

**ERROR rows come only from owners that already refuse.** Model Check displays a
refusal; appearing here never creates one.

### 2.1 Severity assignment

| Condition | Severity |
|---|---|
| Structural / input fault from `ValidateStructure` | **ERROR** |
| Calculation `INVALID` | **ERROR** |
| Calculation `NOT CALCULATED` | **WARNING** — an operator action is outstanding, not a defect |
| Calculation `STALE` | **WARNING** |
| Simulation `INVALID` **while calculation is CURRENT** | **ERROR** — see §2.2 |
| Simulation `INVALID` while calculation is not CURRENT | **INFO** — see §2.2 |
| Simulation `STALE` | **WARNING** |
| Annual `HISTORICAL` or `OTHER Px` | **WARNING** |
| Iterations below the recommendation | **WARNING** |
| No simulation published | **INFO** |
| Annual `NOT PRODUCED` | **INFO** |
| Sensitivity not produced for this run | **INFO** |
| Calculation / simulation `CURRENT`, year count, run id, last attempt | **INFO** |

**Optional outputs are INFO.** A valid, current model that has simply not been
simulated is not defective, and Model Check does not report it as though it
were. No existing contract makes simulation, the annual step or sensitivity
mandatory.

### 2.2 When simulation `INVALID` is a fault of its own

Derived from source, inventing nothing:

```text
DeriveSimStatus rule 1   current prerequisites do not resolve  ->  INVALID
CalcPrepareSimulationInputs refuses unless the calculation is CURRENT
```

So whenever the calculation is `NOT CALCULATED`, `STALE` or `INVALID`, the
simulation is `INVALID` **as a consequence of that** — the same root cause,
already reported by the calculation row.

> **The rule.** Simulation `INVALID` is an actionable **ERROR** if and only if
> the calculation state is `CURRENT`. Otherwise it is displayed faithfully as
> **INFO** and is not counted.

When the calculation is `CURRENT` and the simulation still cannot form a
request, the cause is independent — the iteration or seed request itself does not
resolve — and that is a genuine, separately actionable contract failure.

**The live value is always shown as it is.** `INVALID` is never rendered blank
because no publication exists, and blank is never rendered `INVALID`. Publication
existence is a separate INFO row, exactly as P8-Z proved them separate.

---

## 3. The actionable-count rule

> A row is **actionable** if it is ERROR or WARNING. INFO rows are never counted.
>
> A mirrored state that is a **consequence** of a fault already reported by
> another counted row is INFO, not a second count. Two rows are counted
> separately only when they are two independently actionable contract failures
> with two different remedies.

**Overall Status reconciles exactly to the actionable rows.** The counts in the
summary equal the ERROR and WARNING rows in the register — a reader can add them
up. §2.2 is the first application of this rule, not an exception to it.

---

## 4. Aggregation

- **One row per issue.** No collapsed categories.
- **Ordering, deterministic:** severity (`ERROR`, `WARNING`, `INFO`) → `group` in
  the declared order (`Structure`, `Inputs`, `Calculation`, `Simulation`,
  `Annual`, `Sensitivity`) → `check_id` ordinal. Never worksheet position, never
  discovery order.
- **Duplicate suppression by `check_id` + `subject` only.** Two drivers failing
  one rule are two rows; one rule reported twice for one subject is one row.
- **Owners are never coalesced.** Every row carries its `source`.
- **Live, not publication-based.** Model Check answers *what is true now*.

### 4.1 Row window and overflow

```text
row_window = 100
```

- The first 100 rows in the contracted order — deterministic, not arbitrary.
- The **total issue count is authoritative** and is reported whether or not it
  exceeds the window.
- If the total exceeds 100 the summary states, in contract-owned wording:
  **`Showing the first 100 of <M> checks.`**
- **No silent truncation.** No formula may imply only 100 issues exist: the
  counts are of *all* issues, and only the register is windowed.
- Unused rows emit `NA()` — the accepted no-data representation. An empty slot is
  not a check, exactly as an empty tornado slot is not a category.

---

## 5. UI layout

Consistent with existing PCCM sheets: labels in B, values from D, notes at F;
`freeze_panes: A6`; **no charts**.

```text
Section: Overall Model Status              rows 8-14
  B8   Overall Status         D8   PASS | WARNING | ERROR
  B9   Errors                 D9   count of actionable ERROR rows
  B10  Warnings               D10  count of actionable WARNING rows
  B11  Information            D11  count of INFO rows
  B12  Checks evaluated       D12  total, authoritative
  B13  Evaluated              D13  live, volatile
  B14  Disclosure             D14  "" | "Showing the first 100 of <M> checks."

Section: Validation Results                header 17, first row 18, 100 rows
  B  Check ID   C  Group   D  Severity   E  Subject   F  Message   G  Guidance
```

- **Severity vocabulary:** `ERROR` / `WARNING` / `INFO`.
- **Subject:** the offending permanent id, sheet or input name; blank when the
  issue is model-wide.
- **Guidance:** the remediating action, one sentence.
- **Healthy state:** `PASS`, counts `0 / 0 / n`, and one INFO row stating the
  model is valid — not an empty grid, which reads as "never run".
- **Formatting:** severity as text; counts `#,##0`; no colour that carries
  meaning the text does not.

---

## 6. Live versus persisted

| Row | Kind |
|---|---|
| Structural faults | **live** |
| Calculation state | **live** — via the §8 adapter |
| Last calculation attempt | **persisted**, labelled *(last attempt)* |
| Simulation state | **live** |
| Simulation status *(last evaluated)* | **persisted**, labelled |
| Published run id / iterations run | **historical output** |
| Annual distribution / profile / Px / years | **live** derivations over persisted stamps |
| Sensitivity availability | **persisted comparison**, mirrored whole — blind to model drift, so it is paired with the live simulation state |
| Iterations advisory | **live** — the current requested value |

**A persisted `CURRENT` may never masquerade as live.** Every persisted row
carries its qualifier in the label, and **no persisted value feeds the summary**.

---

## 7. Precedence

```text
any actionable ERROR                        ->  ERROR
no ERROR, any actionable WARNING            ->  WARNING
no ERROR, no WARNING                        ->  PASS
```

INFO rows never affect the summary.

---

## 8. AUTHORISED — the live calculation-state adapter

`PCCM_CalculationStatus` calls `WriteStatusBlock` (`modCalcReport.bas:210`),
which writes `C19:C20`. Excel forbids a cell-called function to write, so a Model
Check cell calling it would show `#VALUE!` — the P8-1 defect exactly.

The derivation is already pure: `DeriveStatus` (lines 647-675) and
`PrepareCurrentCalculation` write nothing. Only the persistence is the caller's.

**Authorised, following the accepted P8-1/P8-3 split:**

- expose the existing pure derivation read-only as
  **`modCalcReport.CalcReportDerivedStatus`**;
- add one thin volatile adapter, **`PCCM_ModelCheckCalculationState`**, in
  **`modResultsState`** — the module that already owns every worksheet-safe
  presentation adapter, and therefore the stronger existing owner;
- **`PCCM_CalculationStatus` is untouched**; no persistence moves; no derivation
  is duplicated.

This is a later Phase-9 integration correction to a Phase-7-owned module and
**must follow the declared-production-correction discipline** when implemented —
the Phase-7 record's §1.1 table and the Phase-8 declared-corrections table, with
its removals named if any.

---

## 9. The <10,000 advisory

### 9.1 Threshold ownership — audited

`input_contract.yaml → inputs.monte_carlo_iterations` says, in its own note:

> *"1000 is the locked hard minimum. No upper limit is imposed; **the <10000
> advisory belongs to Model Check**."*

**The advisory is already assigned to this sheet by the accepted input
contract.** Phase 9 implements an obligation that already exists.

**But the value has no machine-readable owner.** 10,000 appears only as
`default: 10000` and inside the user-facing `prompt` string.

- The prompt is **not** parsed. Advisory logic must never read prose.
- `default` is **not** the recommendation. They coincide today and are different
  concepts: a default is what a new workbook starts with; a recommendation is
  advice about a value the user chose.

**P9-2 adds one machine-readable field** to that input's declaration —
conceptually `recommended_iterations: 10000` — as the single owner. It creates no
new authority; it makes an existing one executable. **No literal 10,000 may
appear in a formula, a builder or a test**: every consumer reads the field.

### 9.2 The advisory contract

| | |
|---|---|
| Threshold | the contract's `recommended_iterations` |
| Trigger | **current requested** iterations (`inpMonteCarloIterations`) `<` threshold |
| Live or persisted | **live** |
| Severity | **WARNING** |
| Message | `Monte Carlo iterations are below the recommended 10,000; simulation precision may be lower at this setting.` |
| Guidance | `Increase Monte Carlo Iterations on Setup to 10,000 or more before the next simulation run.` |
| Clears | automatically when requested ≥ threshold |
| Does the threshold itself warn? | **No** — strictly `<` |
| Before any simulation? | **Shown** — it is advice about the request |
| Invalid model? | **Coexists.** Never suppressed; an ERROR simply outranks it |
| Stale request? | **Shown**, from the current requested value |
| Refusal? | **Never.** The hard minimum of 1000 is unchanged and untouched |

---

## 10. Worksheet-safety audit

Every path Model Check may call from a cell, classified transitively.

| Path | Writes? | Verdict |
|---|---|---|
| `PCCM_StructuralReport` → `ValidateStructure` | **none** — `modStructuralCheck` contains no assignment to `.Value`/`.Value2`, no `ClearContents`, no `.Delete`, no `ListRows.Add` and no application-state change, module-wide | **SAFE as it stands.** No split needed |
| `PCCM_CalculationStatus` | **writes `C19:C20`** via `WriteStatusBlock` | **UNSAFE** — never called from a cell; §8 adapter instead |
| `CalcReportDerivedStatus` (proposed) → `DeriveStatus` → `PrepareCurrentCalculation` | none | **SAFE** |
| `PCCM_ResultsSimulationState` → `SimReportDerivedStatus` → `DeriveSimStatus` | none | **SAFE** — established at P8-1, proved live at P8-3 |
| The four `modResultsState` annual adapters → `modSimAnnualStore` read path | none | **SAFE** — established at P8-1 |
| Sensitivity availability | a worksheet formula over `_SimData`; no VBA | **SAFE** |

**Requirement for P9-2:** no worksheet-called path may write a cell, persist a
status, consume a nonce or run id, call a command endpoint, or mutate workbook or
application state. A control must prove this transitively over the actual call
graph — the P8-3 `SIM_REPORT_CALLERS` detector is the precedent.

---

## 11. Scenario matrix

Actionable counts are `E / W` (errors / warnings). INFO rows are not counted.

| | Scenario | Summary | Counts | Register |
|---|---|---|---|---|
| **A** | untouched workbook | **WARNING** | 0 / 1, or 0 / 2 | WARNING calc `NOT CALCULATED`. INFO live simulation `INVALID` — shown faithfully, **not counted**, its cause is the calculation row (§2.2). INFO no publication. INFO annual `NOT PRODUCED`. INFO sensitivity no publication. The advisory **only if** the current requested iterations are below the recommendation |
| **B** | valid calculated model, no simulation | **PASS**, or **WARNING** if below the recommendation | 0 / 0 or 0 / 1 | INFO calc `CURRENT`. INFO no publication. INFO annual `NOT PRODUCED`. INFO sensitivity not produced. The optional outputs do **not** warn |
| **C** | current simulation, ≥ recommendation | **PASS** unless an independent warning exists | 0 / 0 | INFO calc and simulation `CURRENT`. **No advisory** |
| **D** | current simulation, < recommendation | **WARNING** | 0 / 1 | Exactly one WARNING — the advisory. Everything else INFO |
| **E** | request drift | **WARNING** | 0 / 2+ | WARNING simulation `STALE`. WARNING annual `HISTORICAL` or `OTHER Px` as the owners report them. INFO persisted status *(last evaluated)*, never as live. Advisory iff currently below |
| **F** | invalid triangular input | **ERROR** | 1 / n | ERROR from the structural/input owner, naming the permanent id. Calculation `INVALID` displayed; **same root cause**, so it does not double-count (§3). Simulation `INVALID` INFO (§2.2). Advisory may coexist |
| **G** | request restored to ≥ recommendation | per the other owners | −1 warning | The advisory row is **gone**; every other state warning remains exactly as its owner reports it. The advisory's departure changes no other row |
| **H** | multiple blocking issues and warnings | **ERROR** | n / m | Every independently actionable ERROR listed, deterministically ordered, none suppressed by another. Counts equal the listed rows |
| **I** | duplicate warning source | as applicable | 1 per `check_id`+`subject` | One rule twice for one subject = one row. Two subjects failing one rule = two rows |
| **J** | healthy | **PASS** | 0 / 0 | One INFO row stating the model is valid. Not an empty grid |

Windows acceptance later: **one dedicated minimal runner**, the P8-Z shape. It
re-runs no Phase-8 suite.

---

## 12. P9-2 implementation plan

1. **Contract first.** Add `recommended_iterations` to the input contract (§9.1).
   Declare the check register — ids, groups, severities, messages, guidance — and
   the Model Check layout in `spec/`.
2. **Projection.** `phase9_model_check_inspection.json`: addresses, the 100-row
   window, the severity and group vocabularies, the disclosure wording, the
   threshold. Its validator refuses a missing field, exactly as the P8-3 one does.
3. **The authorised adapter** (§8), with its declaration.
4. **Builder.** Render the two sections as mirrors and adapters. Every unused
   register row `NA()`. No literal threshold anywhere.
5. **Static controls and mutations.** Precedence; the actionable-count rule;
   ordering; the duplicate rule; overflow disclosure; the advisory's strict `<`
   boundary; the two-axis separation; and the transitive worksheet-safety proof.
6. **Only then** the minimal Windows runner.

Steps 1-5 are Linux-settleable. Step 6 is not started by this plan.

---

## 13. Files expected to change in P9-2

```text
spec/input_contract.yaml                        + recommended_iterations
spec/workbook.yaml                              + phase9_shell.model_check
builder/pccm_builder/phase9_model_check.py      NEW projection + validator
builder/pccm_builder/workbook_builder.py        render the two sections
builder/build_stage_a.py                        emit the projection
src/vba/modCalcReport.bas                       + CalcReportDerivedStatus
src/vba/modResultsState.bas                     + PCCM_ModelCheckCalculationState
docs/phase7_closure.md                          declare the src correction
tests/test_phase8_charts.py                     declare the src correction
tests/test_phase9_model_check.py                NEW
bootstrap/windows/phase9_p1_model_check.ps1     NEW — step 6, not P9-2
```

---

```text
P9-2 HAS NOT STARTED.
No Phase-9 production, spec or builder change exists in the tree.
```
