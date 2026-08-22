# PCCM — Phase 4 Gate B, run 4 (target Windows / Excel)

**Outcome: NOT ACCEPTED.** No structural runtime scenario set has passed.

Run 4 is the first to reach and complete the structural runtime harness. Runs 1–3
are recorded in `phase4_gate_b_run1.md`, `run2.md` and `run3.md`; their evidence
stands unchanged.

```text
16 passed, 18 failed, 0 skipped
FAILED: D-J.3, D-J.4, D-J.5, D-J.6, D-J.7, D-J.8,
        K2, M, N, O, P, Q, R, S, T, U, W, Z
```

**The 18 failures are not 18 defects.** They come from a small number of causes;
most are cascades and harness-precondition faults. That distinction is the main
correction in this round: a harness that turns one defect into ten apparent ones
cannot be used to judge a model.

---

## What passed on the target

| Step | Result |
|---|---|
| `PRE0` checklist-factory prerequisite | PASS |
| `PRE` collection-shape preflight | PASS |
| `A` Stage-B bootstrap, complete | PASS |
| `A1` VBA automation surface — **the real VBA automation surface answered** (this row read "the real VBA project compiled and was callable" until Phase-5 Runtime Run 7 disproved the compile half; see `phase5_gate_b_harness.md`) | PASS |
| `B` Permanent Cost Line IDs | PASS |
| `B2` Real ListObject reorder | PASS |
| `C` Permanent Risk IDs | PASS |
| `D0` Seeded keyed Inflation Profile | PASS |
| `D-J.1` First timeline application | PASS |
| `D-J.2` Duration increase 3 → 5 | PASS |
| `D-J.9` Degenerate inflation span | PASS |
| `D-J.10` Rejected: Base Year later than Start Year | PASS |
| `K` Profiling synchronisation | PASS |
| `L` Runtime failure containment | PASS |
| `V` Generated year cell presentation | PASS |

`A1` is the significant one: three runs were spent getting the VBA project to
compile at all, and it now does.

---

## Failure classification

| Scenario | Class | Cause |
|---|---|---|
| `D-J.3` | **REAL MODEL DEFECT** | collision-unsafe header rename (below) |
| `D-J.4` … `D-J.8` | cascade | oracle compared against a predecessor state the workbook never reached after D-J.3 rolled back |
| `K2`, `N`, `W` | harness precondition | D-J.10 deliberately leaves an INVALID entered triple; their Apply was rejected before they tested anything |
| `O` | harness precondition | requires Applied Duration ≥ 2; D-J.9 deliberately ends at duration 1 |
| `P`, `Q`, `R`, `S`, `T`, `U` | contamination | N left `HARNESS TEMP PROFILE` in Config with no inflation row; that one orphan failed everything after it |
| `M`, `Q`, `R` | harness assumption | the Data Validation probe asked the wrong question of the COM API |
| `T` | harness fixture | Add-then-Delete cannot produce a free row, so the fixture wrote at row index 0 |
| `Z` | harness bug | the Excel identity variable was clobbered by a loop iterator |

**One** of those is a model defect.

---

## Root cause 1 — collision-unsafe header renaming (REAL, in the model)

Before `D-J.3` the profiling headers are `2028 2029 2030 2031 2032`; the start-year
shift requires `2030 2031 2032 2033 2034`. The rename was a single sequential pass:

```vb
For i = 1 To NewCount
    target.ListColumns(fixedCols + i).Name = CStr(CLng(StartYear) + i - 1)
Next i
```

Renaming column 1 to `2030` meets a column still named `2030`. Excel does **not**
refuse — it appends a digit. The target reported the characteristic corruption:

```text
20272, 20282, 20292, 20302, 2031, 2032
```

### The fix

One collision-safe primitive, `modWorkbook.SetHeaderBlock`, and it is now the only
place in the model that renames a header:

1. capture the desired final headers as plain `String` data;
2. verify they are unique among themselves, and do not collide with a column
   **outside** the block, which this operation may not rename;
3. generate **deterministic** temporary names — a reserved `PCCM_TMP_HDR_` prefix
   plus the column index, with an incrementing suffix — checked against every
   current name, every desired name and each other. Not random;
4. first pass: rename every affected column to its temporary name;
5. second pass: rename every affected column to its final name;
6. verify every final name is **exactly** what was requested, case-sensitively;
7. raise if any differs. Excel's auto-disambiguation is never accepted.

Used by all three header-mutating paths:

- `modProfiling.SetYearColumns`
- `modInflation.SetYearColumns`
- `modWorkbook.RestoreTable` — **especially** this one. A failed operation could
  otherwise corrupt headers while supposedly rolling them back, and `RestoreTable`
  wrote header cells one at a time, which is the same collision-unsafe rename.

No independent renamer survives: `test_45l` fails on any `ListColumn.Name`
assignment outside the primitive and on any `HeaderRowRange.Cells(...).Value`
write anywhere.

Whether a sequential pass collides is **direction-dependent**, which is the trap:
shifting the same block back down, left to right, happens to vacate each target
before it is needed and survives by luck. `test_45j` demonstrates both — the
forward shift corrupts, the downward shift does not, and a reversal (which
`RestoreTable` can face) corrupts immediately.

---

## Root cause 2 — one failure poisoned the sequential oracle chain

`D-J` is a state machine: `D-J.5`'s expected headers are computed from the state
`D-J.4` was supposed to leave. After `D-J.3` rolled back, steps 4–8 were compared
against a predecessor that never existed, and five cascades were reported as five
independent behavioural defects.

Now: every step asserts its **outcome contract** and prints the exact VBA outcome —

| Step kind | Required outcome |
|---|---|
| expected rejection | `FAIL\|…` |
| expected cancelled change | `OK\|cancelled` |
| expected accepted change | `OK\|…`, and not `OK\|cancelled` |

and once any step fails, the remaining steps are **SKIPPED** with the reason. The
Gate result still fails; the diagnosis is now truthful. The run does **not** stop
globally — independent scenarios still execute.

---

## Root cause 3 — state hygiene after the intentional invalid-timeline test

`D-J.10` correctly proves prevalidation rejects Base Year 2040 with Start Year
2035, and that invalid triple remains in the entered cells afterwards. `K2`, `N`
and `W` then called Apply without changing it, so each was rejected before testing
anything:

```text
Apply / Update Timeline was rejected...
Base Year 2040 is later than Project Start Year 2035.
```

`Sync-EnteredTimelineToApplied` copies applied → entered for all three fields and
**verifies** the copy. It is called explicitly by `K2`, `N`, `T` and `W`, and
reported as a check so a failed normalisation is visible. It is **not** applied
globally: the entered/applied difference is exactly what `D-J` tests, and
`test_44n3` fails if the D-J loop ever calls it.

---

## Root cause 4 — `O` had an impossible precondition

`O` tests destruction of the final profiling year during a shrink, so it needs
Applied Duration ≥ 2. `D-J.9` deliberately ends at duration 1 and `D-J.10` is
rejected, so `O` could not pass in a perfectly correct workbook.

`O` now establishes its own baseline, taken from the **oracle fixture** rather than
invented: the first accepted step whose applied duration is at least 2 (Base 2026 /
Start 2028 / Duration 3). It then verifies the apply succeeded, the applied triple
is the requested one, entered equals applied, and the structural report is clean —
and only then seeds `PASTED TEXT` and runs the destructive-cancel test.

---

## Root cause 5 — `N` contaminated every later scenario

`N` added `HARNESS TEMP PROFILE` to Config; its Apply was rejected (root cause 3),
so no inflation row was created, and

```text
[inflation_profile_rows]
Config profile 'HARNESS TEMP PROFILE' has no row in tblInflation.
```

failed `P`, `Q`, `R`, `S`, `T`, `U` and `W`.

Two fixes. `N` normalises the timeline first, so its Apply succeeds. And the
fixture is removed in a **`finally`** — cleared from Config, synchronised, and the
structural report re-read, with any residue noted — whether `N` itself succeeds or
not. A disposable fixture must never outlive its scenario.

Beyond that, every independent scenario after `N` (`P`, `Q`, `R`, `S`, `T`, `U`,
`W`) now calls `Test-CleanStructure` first. If the workbook is already faulted, the
scenario reports **prerequisite contamination as a SKIP** instead of running its
assertions as though the failures were its own. Contamination is never hidden — it
is printed.

---

## Root cause 6 — `K2`'s order assertion was too weak

It claimed "the profiling grid order follows the reordered register" but compared
**sorted** sets, which proves membership, not order. At run 4 Apply was rejected,
the grid never synchronised, and the assertion passed anyway.

It now captures `profileOrderAfter` and asserts **exact** equality against the
register order, with neither side sorted. Only then does the positional-percentage
assertion mean anything: the register order changed, the grid order changed to the
same order, each ID kept its own seeded percentage, therefore the positional
sequence changed. `K2` also normalises entered → applied before its Apply.

---

## Root cause 7 — the Data Validation check asked the wrong question

The probe read `$cell.Validation`, touched `.Type`, and treated "no exception" as
"this cell carries a user restriction". That is not a safe contract: Excel returns
a `Validation` object either way, and a type of `xlValidateInputOnly` restricts
nothing.

**No model change was made on this evidence.** Instead the harness is calibrated.
New scenario `A2` captures a validation fingerprint — `Type`, `Formula1`,
`Formula2` — for every column of an **untouched Stage-A** row of both registers,
before any driver is added. It then asserts the fingerprint actually
**distinguishes** a validated user column from the model-controlled ID column; a
baseline of "nothing anywhere" would make every later comparison vacuously true.

`M`, `Q` and `R` compare their runtime-grown and restored rows against that
baseline, per column. The ID assertion is phrased as the contract states it — *the
ID cell has no constraining user validation, matching the model-controlled
baseline* — not as a claim about the COM property. Coverage is not weakened: the
validated user columns are still named individually.

If a rerun shows the grown ID cell genuinely carries a different constraining rule
from the Stage-A one, **then** the model needs fixing. `Validation.Delete` has not
been added pre-emptively.

---

## Root cause 8 — `T` did not guarantee an unkeyed blank row

`T` tried Add-then-Delete to free a row, but `DeleteDriver` removes the `ListRow`
when more than one exists, so the table returned to the same count, `$freeRow`
stayed `0`, and the fixture wrote at row index zero.

`Add-BlankTableRow` now adds one blank `ListObject` row directly, with the same
leaf-before-parent COM ownership as every other reader, and deliberately allocates
**no** permanent ID — an unkeyed row holding data is the invariant under test, so
it must not be created through `PCCM_AddCostLine`. `T` then verifies the row
exists, its ID is blank and every cell is blank, writes `ORPHAN DESCRIPTION`, runs
Add, verifies the refusal and that the row is untouched, and removes the fixture.

---

## Root cause 9 — the Excel identity variable was clobbered

```powershell
$id = Get-ExcelIdentity ...        # the process identity
...
foreach ($id in $ids) { ... }      # PowerShell loop variables are not block-scoped
...
Wait-ExcelExit -Identity $id       # $id is now a driver identifier string
```

```text
PropertyNotFoundException: The property 'ProcessId' cannot be found on this object.
```

Renamed to `$excelIdentity` in the harness, and `$buildExcelIdentity` /
`$verifyExcelIdentity` in the bootstrap. Loop variables carry semantic names —
`$driverId`, `$riskId`. `test_44n10` audits the **whole** script rather than the
two loops that happened to be noticed: every variable assigned from
`Get-ExcelIdentity` must be named for it, must never be a `foreach` iterator, and
must never be assigned anything else.

---

## Status

Gate B remains **NOT ACCEPTED**. One real model defect is fixed; the rest of this
round makes the harness capable of telling the truth about the next run. No
structural runtime scenario set has passed.
