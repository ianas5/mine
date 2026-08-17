# PCCM — Phase 4 Gate B, run 5 (target Windows / Excel)

```text
27 passed, 1 failed, 7 skipped
PHASE-4 FUNCTIONAL TEST FAILED: O
```

**Gate B is NOT ACCEPTED** — seven independent scenarios have still not run. But
this run is materially successful, and one thing in it is settled.

Runs 1–4 are recorded in `phase4_gate_b_run1.md` … `run4.md`.

---

## What passed on the target

| | |
|---|---|
| `PRE0` checklist factory | PASS |
| `PRE` collection shape | PASS |
| Stage-B bootstrap | PASS |
| `A` build, CodeNames, modules, buttons | PASS |
| `A1` VBA automation surface | PASS |
| `A2` Data Validation baseline | PASS |
| `B` Permanent Cost Line IDs | PASS |
| `B2` Real ListObject reorder | PASS |
| `C` Permanent Risk IDs | PASS |
| `D0` Seeded keyed Inflation Profile | PASS |
| **`D-J.1` … `D-J.10`** | **ALL PASS** |
| `K` Profiling synchronisation | PASS |
| `K2` Percentage ownership across a real reorder | PASS |
| `L` Runtime failure containment | PASS |
| `M` Growth past reserved capacity, incl. validation-baseline comparison | PASS |
| `N` Removed Config inflation profile | PASS |
| `V` Generated year cell presentation | PASS |
| `Y` Transient COM releases | PASS |
| `Z` Natural Excel shutdown | PASS |

### The collision-safe header correction is proven on real Excel

`D-J.1` through `D-J.10` all pass, including `D-J.3` — the start-year shift
`2028…2032 → 2030…2034` that produced `20272, 20282, 20292, 20302` at run 4. The
two-pass `modWorkbook.SetHeaderBlock` rename holds against real Excel across every
timeline transition in the fixture, in both the forward path and the rollback.

`K2` passing alongside it means profiling percentages stayed with their permanent
IDs across a real `ListObject.Sort` with a live timeline, and `Z` and `Y` mean the
whole run ended with a natural Excel exit and clean COM release.

---

## The single failure: `O` — a harness defect, not a model defect

`O` intends to prove that blank and numeric zero are not destructive data but
pasted text is. To isolate that, it first neutralises the final project-year cell.
It did so by looping **every physical body row** of both profiling grids:

```powershell
for ($r = 1; $r -le (Get-TableRowCount ... $costGrid ...); $r++) { ... -Value 0 }
for ($r = 1; $r -le (Get-TableRowCount ... $riskGrid ...); $r++) { ... -Value 0 }
```

That writes numeric zero into the **unkeyed reserved rows**, which is precisely the
structural orphan the model exists to refuse:

```text
[no_orphan_structural_data]
tblRiskProfiling row(s) 3, 4, 5, 6, 7, ... hold data but carry no key.
```

The real sequence was:

1. baseline structurally clean;
2. the harness writes 0 into unkeyed Risk Profiling rows;
3. the harness writes `PASTED TEXT` into a keyed Cost Profiling row;
4. `PCCM_ApplyTimeline` begins;
5. `PreMutationCheck` sees the unkeyed zeros;
6. **Apply is refused before `BuildSummary` and the destructive assessment**;
7. no prompt exists, so `O` reports its two prompt checks as failures;
8. the success-path cleanup clears only the Cost text;
9. the unkeyed Risk zeros survive;
10. `P`, `Q`, `R`, `S`, `T`, `U`, `W` correctly SKIP on contaminated prerequisites.

**The model behaved correctly throughout.** The orphan protection did exactly what
it is for, and the operation never reached the assessment `O` was testing. Nothing
in this evidence indicates a defect in `IsDataCell`, `CountDataBeyond`,
`PreMutationCheck`, orphan semantics, profiling synchronisation, `SetHeaderBlock`
or driver logic, and none of those changed in this patch.

---

## The corrections

**1. Keyed-only neutralisation.** Each profiling grid is walked row by row and the
tail cell is written **only where the row carries a permanent ID**. An unkeyed row
is not touched at all.

**2. Each grid's own width.** `$fixedRisk = $riskGrid.fixed_columns.Count` is
introduced, and `O` indexes each grid by its own contract-derived count. The two
grids happen to have two fixed columns each today; the Risk write was riding on
that coincidence and would have moved sideways if the schemas ever diverged.

**3. The fixture is proved clean before Apply.** After the tail cells are
neutralised, `PASTED TEXT` is placed and the entered duration reduced — but
**before** `PCCM_ApplyTimeline` — `PCCM_StructuralReport` must be empty. Pasted
text in a *keyed* cell is business-invalid but not a structural orphan, and `O`
needs the workbook coherent so the operation reaches `CountDataBeyond`. **The
run-5 package fails this assertion**, which is the point.

**4. The outcome is asserted, not inferred.** `PCCM_AutomationResult` is captured
and must be exactly `OK|cancelled`, checked *before* the prompt claims. Reading
only the prompt is what let a refusal masquerade as a missing warning.

**5. Cleanup is failure-safe and exact.** Original tail values are captured as
plain data keyed by **grid + permanent ID**, never by physical row position. A
`finally` restores every touched ID to its exact captured value — a blank comes
back **blank**, not as numeric zero — restores entered = applied, and re-reads the
structural report. No structural synchronisation is called merely to erase
residue. Cleanup problems are noted and keep the gate failed, because the result
is reported *after* cleanup.

**6. The unkeyed rows are proved untouched.** Their tail cells are captured before
anything is written, asserted unchanged after the fixture is built, and asserted
unchanged again after cleanup. The claim is that `O` never writes to them — not
that it tidies up afterwards.

**7. `O` ends with its own clean assertion**, and `P`'s `Test-CleanStructure`
gate remains as independent defence.

---

## Direct-write fixture audit

Every scenario write into Cost Profiling, Risk Profiling, Inflation, Cost Lines or
the Risk Register, classified:

| Scenario | Write | Guard |
|---|---|---|
| `B2` | register Description marker | `$row[0] -ne ''` — keyed |
| `D-J` | profiling percentages | **was unguarded — wrote to row 1 on the assumption it was keyed.** Now resolves the first keyed row |
| `D-J` | deliberate blank in a second keyed row | **was `-RowIndex 2` on the same assumption.** Now resolves a second keyed row |
| `D-J` | inflation rates | row found by matching the profile name — named |
| `K2` | profiling percentages, register markers | `$row[0] -ne ''` — keyed |
| `N` | inflation rate | row found by matching the profile name — named |
| `O` | profiling tail cells | **fixed** — keyed only |
| `T` | orphan fixtures | **deliberate, and must remain** — an unkeyed row holding data is the invariant under test |

Two latent cases beyond `O` were found and closed: the `D-J` seeding wrote to rows
1 and 2 by position. Those rows happen to be identified at that point, but the
fixture rested on it, which is the same class of assumption that poisoned run 5.

---

## Regressions

| Test | Proves |
|---|---|
| `test_44o1` | no structural-table write outside a helper sits without a permanent-ID guard, scenario by scenario, with `T`'s deliberate orphans excepted; and the exact run-5 blind-loop shape is rejected by name |
| `test_44o2` | `O` zeroes keyed rows only, and indexes each grid by its own fixed count |
| `test_44o3` | the fixture is proved structurally clean after it is complete and before Apply |
| `test_44o4` | `OK\|cancelled` is asserted exactly, before the prompt claims |
| `test_44o5` | cleanup is in a `finally`, restores exact captured values, restores blanks as blanks, calls no synchronisation, and reports its own failure into the verdict |
| `test_44o6` | the unkeyed rows are captured first and asserted unchanged twice, and the later independent gates survive |

All six fail against the run-5 package.

---

## Status

Gate B remains **NOT ACCEPTED**: `P`, `Q`, `R`, `S`, `T`, `U` and `W` have not yet
run. Everything before them now passes on the target, and the model source is
unchanged by this round.
