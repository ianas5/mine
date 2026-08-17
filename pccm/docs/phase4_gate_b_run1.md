# PCCM — Phase 4 Gate B, run 1 (target Windows / Excel)

The first execution of the Phase-4 Gate-B package on the real target environment.

**Outcome: NOT ACCEPTED.** The structural runtime scenarios B onward were skipped
and have **not** run. Nothing in this document claims otherwise.

What the run did establish is real, and is recorded here because it was proved on
the target machine rather than argued from source.

---

## What the target environment demonstrated

Every substantive Excel operation in the Stage-B bootstrap completed before the
exception, and **both** Excel processes exited naturally, with no forced
termination:

| Step | Result |
|---|---|
| Read Stage-A build outputs | PASS |
| Excel COM instance creation (owned, identified) | PASS |
| Open the Stage-A workbook | PASS |
| `SaveAs` macro-enabled `.xlsm`, FileFormat 52 | PASS |
| Apply all **14** worksheet CodeNames | PASS |
| Import all **8** Phase-4 VBA modules | PASS |
| Create all **5** command buttons, `OnAction` assigned | PASS |
| Save the Stage-B workbook | PASS |
| Build instance closed **naturally** | PASS |
| Reopen the `.xlsm` in a **fresh** Excel instance | PASS |
| Verify 14 CodeNames, 8 modules and 5 buttons **persisted** | PASS |
| Verification instance closed **naturally** | PASS |

This is the first empirical confirmation of several things the Linux tests could
only assert about the source:

- the two-stage build produces a real macro-enabled workbook on the target;
- worksheet CodeName assignment through the VBProject works, and **persists**
  across a close and a reopen in a different process;
- all eight modules import and survive the round trip, including the generated
  `modConstants.bas`;
- button creation and `OnAction` binding persist;
- the COM ownership discipline proved at the Phase-1.6 readiness gate holds in
  the production build path: two separate Excel instances, both shut down
  naturally, with `Marshal.ReleaseComObject` only and no forced stop.

None of that is erased by the wrapper failing afterwards.

---

## The defect

Not a VBA or runtime-model failure. A PowerShell reporting bug, in the very last
block of the bootstrap, **after** every Excel operation had succeeded.

```
System.Management.Automation.PropertyNotFoundException:
The property 'Count' cannot be found on this object.
```

`com_lifecycle.ps1`:

```powershell
function Get-TransientFailures { return @($script:transientFailures) }
```

`build_stage_b.ps1`:

```powershell
$transient = Get-TransientFailures      # <- lands $null when there are none
if ($transient.Count -gt 0) {           # <- PropertyNotFoundException
```

A PowerShell function returning an **empty** collection emits **zero pipeline
objects**. The `@(...)` written inside the helper is evaluated before the return
value crosses the function boundary, so it does not survive: the assignment
receives nothing and `$transient` is `$null`. Under `Set-StrictMode -Version 2.0`,
reading `.Count` on `$null` throws.

It fires **specifically on the success path**, because zero transient release
failures is exactly the condition that produces an empty collection. A run with a
release failure would have reported normally.

The parent harness therefore recorded `[FAIL] A Stage-B build` and skipped the
structural runtime scenarios — which is why Gate B is not accepted.

---

## The fix

Materialisation belongs at the **caller**, where collection semantics are
required:

```powershell
$transient = @(Get-TransientFailures)
if ($transient.Count -gt 0) {
```

Applied in `build_stage_b.ps1` and in `phase4_functional_test.ps1`, which carried
the identical pattern and would have thrown there too — after every scenario had
already passed.

### The rule, applied uniformly

Fixing only the site that happened to fire would leave the same latent failure at
every other caller, waiting for the run where that particular collection comes
back empty. So the rule is applied everywhere:

> If zero results are a valid outcome and the caller needs collection semantics,
> materialise at the caller: `$items = @(Get-Something)`.

Scalar-returning helpers are **not** rewritten. Wrapping one would turn a string
into a one-element array and break every comparison against it; `test_46p` fails
if `Get-NamedValue`, `Get-TableRowCount` or `Get-TrustAccessGuidance` is ever
swept up that way.

### Audit result

| Helper | Returns | Callers | Action |
|---|---|---|---|
| `Get-TransientFailures` | `string[]`, often empty | 2 | materialised — **this is the defect that fired** |
| `Get-PreExistingExcelPids` | `int[]`, empty when no Excel is running | 3 | materialised |
| `Get-IdColumnValues` | `string[]`, empty before the first Add | 25 | materialised |
| `Get-TableBody` | jagged `string[][]` | 53 | materialised, **and see below** |
| `Get-TableColumnNames` | `string[]` | 10 | materialised |
| `Get-TrustAccessGuidance` | one joined string | 1 | left alone — scalar |
| `Get-NamedValue`, `Get-TableRowCount` | scalar | many | left alone — scalar |

The audit also found `$costBefore.Count` and `$costAfter.Count` on unmaterialised
`Get-TableBody` results — the same strict-mode exception, in the middle of the
timeline scenarios rather than at the end of the run.

### A second defect of the same class

`Get-TableBody` returns a **jagged** array, and caller-side `@(...)` alone cannot
repair that one:

- `return $rows` holding **one** row emits that row's inner `string[]` as the
  single pipeline object, and the caller's `@(...)` then enumerates **that** — so
  a one-row table arrives as N rows of one cell each;
- zero rows is the familiar `$null`.

Only the unary comma stops the outer array being enumerated on output:

```powershell
return ,$rows
```

after which the caller's `@(...)` unwraps exactly one level and gets `$rows` back
intact, for zero, one or many rows alike. The flat helpers need no comma — one
string emitted is one string, which `@(...)` makes a one-element array.

This one did not fire: no table in the current workbook is one row long. It is
recorded because it is the same class, found by auditing for it rather than by
waiting for it.

---

## Regressions

| Test | Proves |
|---|---|
| `test_46l` | both scripts materialise `Get-TransientFailures` with `@(...)`, and still test the count |
| `test_46m` | every call to every collection helper is materialised at the caller |
| `test_46n` | no `$x = Get-Collection` … `$x.Count` shape survives anywhere |
| `test_46o` | the jagged helper suppresses output enumeration; the flat ones deliberately do not |
| `test_46p` | scalar helpers were **not** swept up in the rewrite |

`test_46l`, `test_46m`, `test_46n` and `test_46o` were each verified to fail
against the exact source that ran on Windows.

---

## Status

**Gate B is not accepted.** Scenarios B onward have not run. The corrected package
is ready for one rerun on the target environment.
