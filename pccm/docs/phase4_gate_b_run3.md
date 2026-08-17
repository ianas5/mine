# PCCM — Phase 4 Gate B, run 3 (target Windows / Excel)

**Outcome: NOT ACCEPTED.** No structural runtime scenario has run.

Run 3 is the first to reach the **real VBA execution path**, and it produced a
genuine VBA source defect rather than a harness bug. That is progress, and it is
also the reason it stopped.

Runs 1 and 2 are recorded in `phase4_gate_b_run1.md` and
`phase4_gate_b_run2.md`. Their evidence stands unchanged.

---

## What passed on the target

| Step | Result |
|---|---|
| `PRE0` checklist-factory prerequisite | PASS |
| `PRE` collection-shape preflight | PASS |
| Stage-B bootstrap, complete | PASS |
| Stage-A open | PASS |
| `SaveAs` `.xlsm`, FileFormat 52 | PASS |
| All 14 worksheet CodeNames applied | PASS |
| All 8 Phase-4 VBA modules imported | PASS |
| All 5 command buttons created | PASS |
| Workbook saved | PASS |
| Build Excel process exited **naturally** | PASS |
| Fresh Excel process reopened the `.xlsm` | PASS |
| 14 CodeNames / 8 modules / 5 buttons persisted | PASS |
| Verification Excel process exited **naturally** | PASS |
| All transient COM releases clean | PASS |

Both preflights introduced after runs 1 and 2 did their job: `PRE0` proved the
checklist factory, `PRE` proved the row-emission contract, and the bootstrap ran
end to end with clean COM shutdown.

## Where it stopped

The harness then made its **first real VBA invocation**:

```powershell
$excel.Run('PCCM_AutomationBegin', $true, '')
```

Excel opened the VBA editor and raised:

```text
Compile error:
Variable not defined
```

on the identifier `gAutomationActive`, inside:

```vb
Public Sub PCCM_AutomationBegin(ByVal ConfirmReply As Boolean, ByVal FailAfterStage As String)
    ClearAutomation
    gAutomationActive = True
```

**No structural runtime scenario has passed.** Scenario B onward never ran.

---

## Root cause

`modAppState.bas` has `Option Explicit`, and its declaration section correctly
began with `AppStateSnapshot`, `OperationResult` and `MSG_TITLE`. But the five
automation globals were written **far later in the file**, after
`ConfirmDestructiveChange`:

```vb
Public gAutomationActive          As Boolean
Public gAutomationConfirmReply    As Boolean
Public gAutomationLastPrompt      As String
Public gAutomationFailAfterStage  As String
Public gAutomationLastResult      As String
```

VBA has no "module-level statement anywhere in the file". Everything **before the
first executable procedure** is the declaration section; everything after it is
procedure bodies. Those five sat after the first procedure, so under
`Option Explicit` the compiler reached `PCCM_AutomationBegin` with
`gAutomationActive` undefined and stopped.

This is a **real VBA source defect**, not a PowerShell harness defect.

### Why Linux never saw it, and why scenario A did not either

- There is no VBA compiler on Linux. Every compile-time rule the tests do not
  encode is a Gate-B blocker waiting to happen — this is the second such rule to
  reach the target, after the `Optional` parameter caught at Gate A.
- **Importing a module is not compiling it.** Scenario A imported eight modules,
  saved, reopened in a fresh process and verified they persisted — all correctly,
  all green — while the project did not build at all. Excel compiles on the first
  `Application.Run`.

---

## The fix

All five declarations moved into the declaration section, immediately after
`MSG_TITLE` and **before** `CaptureAppState`, the first executable procedure. They
are declared **once** — the originals were removed from the Automation Hooks
section, which now carries only a pointer — and they remain **`Public`**, because
`modTimeline` and `modDrivers` reference `modAppState.gAutomationActive` by
qualified name. The automation interface itself is unchanged.

### The audit found a second offender

A full declaration-section sweep of all eight modules found one more, which the
independent review had not reached and which would have failed the **next** run
after this one was fixed:

```text
modWorkbook:312: Public Type TableSnapshot
```

`TableSnapshot` — the logical-rollback snapshot type — was declared beside the
procedures that use it, after `Sh()`. It has been moved into `modWorkbook`'s
declaration section on the same reasoning.

**Audit result: 8 modules, 0 remaining offenders.**

| Module | First executable procedure | Offenders |
|---|---|---|
| `modAppState` | `CaptureAppState` | fixed (5) |
| `modWorkbook` | `Sh` | fixed (1) |
| `modConstants` (generated) | none — declarations only | 0 |
| `modDrivers` | `CounterName` | 0 |
| `modInflation` | `RequiredFirstYear` | 0 |
| `modProfiling` | `ProfilingTable` | 0 |
| `modStructuralCheck` | `ValidateStructure` | 0 |
| `modTimeline` | `ReadTriple` | 0 |

---

## New regression

`test_45f` sweeps **every** module under `Option Explicit` for a module-level
variable, `Const`, `Type` or `Enum` declared after the first executable procedure.
It works on **logical** statements — line continuations joined first, via a shared
`logical_statements()` helper in `vba_source.py` — because a wrapped declaration is
still one declaration, and per-physical-line analysis both misses real offenders
and invents false ones.

It is emphatically **not** a grep for `Dim`:

| Test | Proves |
|---|---|
| `test_45f` | no module-level declaration follows the first executable procedure, in any module |
| `test_45g` | procedure-local `Dim`, `Const` and `Static` are **not** flagged — on real source with 16 such locals, on a constructed legal module, and on a constructed broken one where exactly the three offending forms are caught |
| `test_45h` | continuations are joined, so a wrapped declaration is read as one statement and located by its first line |
| `test_45i` | the five `gAutomation*` names are each declared **once**, in the declaration section, and stay `Public` because two other modules reference them |

`test_45f`, `test_45g` and `test_45i` fail against the exact package used for run
3, reporting all six offenders by module and line.

---

## `A1` — a named boundary for the first VBA call

The compile failure arrived between scenario A and scenario B, so it had no step
of its own and would otherwise read as a permanent-ID defect. `A1` now runs
immediately after the functional-runtime workbook is opened and before scenario B,
and calls the real `PCCM_AutomationBegin`, `PCCM_AutomationResult` and
`PCCM_AutomationEnd`, then re-arms automation for the scenarios below.

Its purpose is to give the first `Application.Run` a **name**, not to claim it is a
VBA compiler — three entry points are not a compiler, and the scenarios below are
what prove behaviour. Errors are deliberately not suppressed: no
`On Error Resume Next`, no `-ErrorAction SilentlyContinue`. A failure reports
Excel's own message as `[FAIL] A1` and then rethrows, so the scenarios never
compare results from a project that never compiled. `test_37b` enforces all of
that, including that `A1` really is the first `Application.Run` in the script.

---

## Status

Gate B remains **NOT ACCEPTED**. Runs 1 and 2 evidence is unchanged; run 3 adds
the first confirmation that the Stage-B artifact reaches Excel's VBA compiler at
all. No structural runtime scenario has passed. The corrected package is ready for
one rerun.
