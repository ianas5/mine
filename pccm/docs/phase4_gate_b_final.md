# PCCM — Phase 4 Gate B, final acceptance run

```text
Working copy:
C:\Users\pcd\AppData\Local\Temp\pccm-phase4-20260817-190848

35 passed, 0 failed, 0 skipped

PHASE-4 FUNCTIONAL TEST: ALL CHECKS PASSED
```

**This is the acceptance run.** Every Gate-B scenario executed. Nothing failed and
nothing was skipped, so no structural-runtime scenario remains unexecuted.

```text
PHASE 4 — GATE B ACCEPTED
PHASE 4 — ACCEPTED / CLOSED
```

---

## Evidence

| Stage | Result |
|---|---|
| Linux/static tests, Phase 1–4 | **502 / 502 PASS** |
| Post-build verification | **181 / 181 PASS** |
| Windows functional scenarios | **35 / 35 PASS**, 0 FAIL, 0 SKIP |

The first two were established on the submitted package before the Windows run;
the third is the target result above. No Windows result here has been rerun,
inferred or invented.

---

## What the real Windows / Excel target demonstrated

**Prerequisites, before Excel is started**

- `PRE0` checklist-factory prerequisite
- `PRE` collection-shape prerequisite

**Build path**

- Stage-B `.xlsx → .xlsm` bootstrap
- owned Excel COM lifecycle
- FileFormat 52
- all 14 worksheet CodeNames
- all 8 Phase-4 VBA modules
- all 5 Phase-4 buttons
- fresh-instance persistence verification

**Runtime surface**

- real VBA compilation and callability, through `A1`
- Stage-A Data Validation baseline, through `A2`

**Permanent identity**

- permanent Cost Line IDs and non-reuse
- permanent Risk IDs and independent non-reuse
- real `ListObject` reorder identity preservation

**Timeline**

- first timeline application
- duration growth
- overlapping start-year shift
- destructive duration shrink, cancelled
- destructive duration shrink, accepted
- Base Year moved earlier and later
- combined Base / Start / Duration transition
- degenerate inflation span
- invalid Base > Start rejection

**Profiling and inflation**

- profiling synchronisation by permanent ID
- profiling percentage ownership across real reordering
- Config / Inflation profile structural handling
- non-numeric profiling destructive-data assessment

**Failure containment**

- logical rollback after an injected runtime failure
- Add rollback after mutation
- Delete rollback after mutation
- Excel application-state restoration
- oversized pasted timeline rejection

**Structural integrity**

- table growth beyond reserved capacity
- presentation and Data Validation preservation
- orphan / unkeyed structural-data refusal
- counter corruption handling
- generated-year presentation contract
- representation-ceiling behaviour

**Shutdown**

- natural functional Excel-process shutdown
- clean transient COM release

---

## How the gate was reached

Runs 1–5 are recorded in `phase4_gate_b_run1.md` … `run5.md` and are **not**
rewritten. They are the failure-discovery evidence, and each issue they exposed
gained regression coverage before the final rerun.

| Run | Exposed | Class |
|---|---|---|
| 1 | PowerShell empty-collection caller semantics — a function returning an empty collection emits zero pipeline objects, so `$x.Count` threw under `Set-StrictMode` | harness |
| 2 | mutable `ArrayList` factory semantics — an empty enumerable returned from a factory arrives as `$null` | harness |
| 3 | VBA declaration-section errors — module-level declarations written after the first executable procedure | **model source** |
| 4 | collision-unsafe Excel `ListObject` header renaming | **model** |
| 4 | several harness state and precondition defects — a poisoned sequential oracle chain, an invalid entered timeline left behind by a deliberate test, an uncleaned fixture, a membership-not-order assertion, an unsafe Validation assumption, an unreachable blank-row fixture, a clobbered process-identity variable | harness |
| 5 | Scenario O fixture contamination — the fixture wrote numeric zero into unkeyed reserved profiling rows, which is the orphan the model is designed to refuse | harness |

Two of these were genuine defects in the model source: the misplaced VBA
declarations found at run 3, and the collision-unsafe header rename found at run
4. **The rest were defects in the Windows harness itself**, not in the model, and
they are recorded as such. In particular, run 5's failure of scenario `O` was its
own fixture breaking the invariant under test; the model's orphan protection
behaved exactly as designed and stopped the operation before the assessment `O`
was written to exercise.

Every one of them — model and harness alike — is now covered by a static
regression that was verified to fail against the package that exhibited it.

---

## Status

```text
PHASE 4 — GATE B ACCEPTED
PHASE 4 — ACCEPTED / CLOSED
```

Phase 5 has not begun.
