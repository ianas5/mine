# PCCM VBA source

Authoritative source for the production VBA. These are plain text files under
version control; the `.xlsm` is a build artifact Stage B assembles from them.

## Phase 4 module inventory

| Module | Responsibility |
|---|---|
| `modWorkbook.bas` | Access primitives: sheets, defined names, list objects, blank handling, table snapshots. No structural policy. |
| `modAppState.bas` | Application state save/restore, structured failure reporting, confirmation prompts, the harness automation hooks. |
| `modTimeline.bas` | Apply / Update Timeline: prevalidation, one combined old→new delta, orchestration, logical restore. |
| `modDrivers.bas` | Permanent ID allocation and the Add / Delete Cost Line and Risk operations. |
| `modProfiling.bas` | Profiling grids: year columns by project-year index, rows synchronised by permanent ID. |
| `modInflation.bas` | Inflation grid: year columns by calendar year, profile rows synchronised by name. |
| `modStructuralCheck.bas` | Focused Phase-4 structural revalidation. Reports; never repairs. |

`modConstants.bas` is **not** here. It is generated from
`spec/structure_contract.yaml` into `build/vba/modConstants.bas` on every Stage-A
build, because a hand-written copy would be a second definition of every sheet
name, table name, defined name, ID prefix and limit in the model. The inventory
test fails if a hand-written `modConstants.bas` ever appears in this directory.

## What is deliberately absent

- Any `Worksheet_Change`, `Workbook_SheetChange` or other input event handler.
  Structural operations are command-driven so there are no hidden side effects
  and the entered-vs-applied distinction stays explicit.
- Any cost, risk, escalation, FX, NPV, EMV, simulation, RNG, Model Check,
  sensitivity or results logic. `spec/structure_contract.yaml → vba.forbidden_constructs`
  lists the constructs the static tests assert are absent.
- Class modules and document modules. Neither is needed for structural runtime.

## How this is checked without a compiler

VBA cannot be compiled on Linux, so `tests/test_phase4_stage_b_source.py` reads
it as text, after separating code from commentary and from string literals:

- every declared entry point exists, and no orphan `PCCM_` macro does
- every SCREAMING_CASE constant referenced here is emitted by `modConstants`
- every `modXxx.Member` reference resolves to a real procedure in that module
- no forbidden construct appears in code

None of that proves the runtime is correct. Only a clean Windows functional run
does.
