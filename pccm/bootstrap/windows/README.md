# PCCM Stage B bootstrap (Windows / Excel COM)

Production build path from the generated Stage-A `.xlsx` to the macro-enabled
`.xlsm`, plus the Phase-4 functional test harness.

This directory is deliberately separate from `pccm/readiness/windows/`, which
holds the disposable Excel COM smoke test. That readiness gate is closed; its
script is a throwaway diagnostic, not production build code, and the two are not
mixed. What *is* shared is the lifecycle **policy** it proved — reimplemented
once here, in `com_lifecycle.ps1`.

## Files

| File | Role |
|---|---|
| `com_lifecycle.ps1` | The COM ownership policy, dot-sourced by both scripts below. One implementation, so they cannot drift apart. |
| `build_stage_b.ps1` | `.xlsx` → `.xlsm`: CodeNames, VBA import, buttons, save, reopen in a fresh instance, verify. |
| `phase4_functional_test.ps1` | Phase-4 functional test matrix A–W, run against a disposable copy. |

## Inputs

Run the Stage-A build first. It emits everything these scripts read:

    python3 pccm/builder/build_stage_a.py

    build/PCCM_stageA.xlsx        the workbook
    build/vba/modConstants.bas    generated VBA constants
    build/stage_b_manifest.json   CodeNames, modules, buttons, entry points, file format
    build/phase4_scenarios.json   oracle-derived expected timeline shapes

Neither script restates a sheet name, CodeName, macro name, button caption or
expected year. They read the manifest and the fixture, so a contract change flows
through without a PowerShell edit.

## Order

    powershell -ExecutionPolicy Bypass -File .\build_stage_b.ps1
    powershell -ExecutionPolicy Bypass -File .\phase4_functional_test.ps1

The functional test runs the bootstrap itself, against a temporary copy under
`%TEMP%`, so it never touches `build/`.

Gate-A source review is approved. Five Gate-B runs have happened on the target.
Everything up to and including scenario `N` now passes; the seven independent
scenarios `P`, `Q`, `R`, `S`, `T`, `U`, `W` have **not yet run**:

- **run 1** completed the whole bootstrap — including natural shutdown of two
  Excel instances — then threw in its final reporting block
  (`../../docs/phase4_gate_b_run1.md`);
- **run 2** aborted in the preflight, before Excel was started, on a broken
  checklist factory (`../../docs/phase4_gate_b_run2.md`);
- **run 3** passed both preflights and the whole bootstrap, then hit
  `Compile error: Variable not defined` on the first `Application.Run` — a real
  VBA source defect, module-level declarations written after the first executable
  procedure (`../../docs/phase4_gate_b_run3.md`);
- **run 4** compiled and ran the whole structural harness — 16 passed, 18 failed.
  **One** real model defect (collision-unsafe ListObject header renaming); the
  rest were cascades and harness-precondition faults
  (`../../docs/phase4_gate_b_run4.md`);
- **run 5** — 27 passed, 1 failed, 7 skipped. `D-J.1`…`D-J.10` all passed, so the
  collision-safe header correction is **proven on real Excel**. The one failure,
  `O`, was its own fixture writing numeric zero into unkeyed reserved rows — the
  orphan the model is designed to refuse — and its residue correctly skipped the
  seven scenarios after it (`../../docs/phase4_gate_b_run5.md`).

All are fixed and the package is ready for one rerun. The model source is
unchanged by the run-5 patch.

Importing a VBA module is **not** compiling it. Scenario A can be entirely green
while the project does not build; Excel compiles on the first `Application.Run`,
which is what scenario `A1` now names.

## Prerequisite

Importing VBA requires **Trust access to the VBA project object model**
(File → Options → Trust Center → Trust Center Settings → Macro Settings). That is
a one-time choice a person makes on the machine.

These scripts report the prerequisite and stop if it is missing. They do **not**
enable it, do not lower macro security, do not edit the registry, and do not add
a Trusted Location — the same refusal that has held since the first readiness run.

## COM lifecycle policy

Unchanged from the run that closed the readiness gate:

- `Marshal.ReleaseComObject` only. `FinalReleaseComObject` is prohibited.
- The release count is reported, never interpreted as failure, never looped to zero.
- No generic COM stack, release plan or object graph. Every long-lived object has
  a named variable and an explicit release point.
- Leaf before parent; `Workbook.Close` before releasing the workbook;
  `Application.Quit` before releasing the application.
- Diagnostic collections hold plain data only, never an RCW.
- **Collections are materialised at the caller**, not inside the helper. A
  PowerShell function returning an empty collection emits zero pipeline objects,
  so `$x = Get-Something` lands `$null` and `Set-StrictMode` turns `$x.Count` into
  a `PropertyNotFoundException`. Every caller writes `@(Get-Something)`. Scalar
  helpers are left alone. This ended Gate-B run 1 — see
  `../../docs/phase4_gate_b_run1.md`.
- **A test fixture must not break the invariant it is testing.** Numeric zero is
  non-destructive data only in a KEYED profiling cell; an unkeyed row must stay
  entirely blank. Every scenario write into a structural table is guarded by a
  permanent ID, except `T`'s deliberate orphan fixtures. This failed Gate-B run 5.
- **A container factory is not an element producer.** `New-Checklist` returns ONE
  mutable `ArrayList`, emitted with `Write-Output -NoEnumerate`; wrapping it in
  `@(...)` would satisfy the rule above and break every caller, because `@(...)`
  yields an `object[]` with no `.Add()`. `New-ReleaseLedger` is safe by returning
  a `PSCustomObject`, which is not enumerable. This ended Gate-B run 2.
- **A producer of rows emits one pipeline object per row**, through
  `Write-RowObject` (`Write-Output -NoEnumerate`). Not `return $rows`, which
  unrolls a single row into its cells, and not `return ,$rows`, which emits the
  whole table as one object and leaves the caller's `@(...)` one level too deep.
  The harness proves this on the target with a pure-PowerShell preflight before
  Excel is started, and aborts the run if the shape is wrong.
- **Ownership starts at the assignment, not at the first successful use.** Every
  acquired object reaches a release on the exception path as well as the success
  path. An inline early release is allowed only when the enclosing `finally` also
  releases the same variable if it is still non-null.
- The acceptance criterion is **actual clean Excel shutdown**. A forced stop is an
  emergency path, only ever applied to a process this script created and can still
  positively identify by HWND-derived PID, process name and start time — and it is
  never reported as a pass.
