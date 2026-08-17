# PCCM — Phase 4 Gate B, run 2 (target Windows / Excel)

**Outcome: NOT ACCEPTED.** No structural runtime scenario has run. The preflight
did not pass — it failed, and it failed on its own prerequisite rather than on
anything it was written to test.

---

## What actually happened on the target

```text
[FAIL] PRE  Collection shape preflight
System.Management.Automation.RuntimeException:
You cannot call a method on a null-valued expression.

PHASE-4 FUNCTIONAL TEST ABORTED before Excel was started.
```

| | |
|---|---|
| `PRE` began on the real target | yes |
| Excel started | **no** |
| Excel processes created | **none** |
| COM side effects | **none** |
| Structural runtime scenarios run | **none** |
| Result | harness prerequisite bug |

The abort behaved exactly as designed: the preflight ran first, failed, and
stopped the run before a single Excel process was created. That containment is the
one thing this run confirms. It is not a pass, and `PRE` is not reported as one.

**Run 1's evidence remains valid.** The target machine has already demonstrated
Excel COM instance creation, Stage-A open, `SaveAs` FileFormat 52, all 14
worksheet CodeNames, all 8 VBA modules imported, all 5 buttons created with
`OnAction`, Stage-B save, natural shutdown of the build instance, fresh-instance
reopen, persistence verification and natural shutdown of the verification
instance — see `phase4_gate_b_run1.md`. Nothing here changes that.

---

## Root cause

```powershell
function New-Checklist { return (New-Object System.Collections.ArrayList) }
```

`System.Collections.ArrayList` is **enumerable**. A newly created checklist is
**empty**. An empty enumerable emitted from a PowerShell function produces **zero
pipeline objects**, so:

```powershell
$list = New-Checklist        # $list is $null
```

and the first check inside `Add-Check` calls `$List.Add(...)` on `$null`, giving
the exact exception observed.

This is the **same class** as the run-1 defect — an empty collection crossing a
function boundary as nothing — but it is **not the same rule**, and applying the
run-1 rule here would have been wrong. Two contracts were conflated:

| | Contract | Caller |
|---|---|---|
| **Element producer** | emits zero, one or many **values** | materialises with `@(...)` |
| **Container factory** | returns **one object**, possibly empty at birth | keeps it and mutates it |

`New-Checklist` is a factory. `@(New-Checklist)` would have satisfied the
collection-materialisation rule and broken every caller, because `@(...)` yields an
`object[]` of the ArrayList's elements — which has no `.Add()`.

It affected the **whole harness**, not only `PRE`: there are 21 `$list =
New-Checklist` call sites, so every later scenario would have failed on its first
check even if `PRE` had been skipped.

---

## The corrected factory

```powershell
function New-Checklist {
    $list = New-Object System.Collections.ArrayList
    Write-Output -NoEnumerate $list
}
```

`-NoEnumerate` emits the ArrayList **itself**, so the caller receives the real
mutable object and `.Add()` keeps working. Converting the checklist to a plain
array was not an option: every caller mutates it in place.

Forbidden, and now enforced statically: `return (New-Object
System.Collections.ArrayList)` and `return $list` — both enumerate the empty list
and emit nothing.

---

## `PRE0` — a prerequisite that does not use the machinery it tests

`PRE` builds its findings in a checklist, so it could not also be what proved the
checklist factory works. When the factory returned `$null`, `Add-Check` threw
before a single row-shape check had run: the test infrastructure rested on an
untested prerequisite.

`PRE0` now runs first, uses **no checklist and no `Add-Check`**, and throws
directly:

```powershell
$probeChecklist = New-Checklist
if ($null -eq $probeChecklist) { throw ... }                              # non-null
if (-not ($probeChecklist -is [System.Collections.ArrayList])) { throw ... }  # type
$null = $probeChecklist.Add('sentinel')                                  # .Add works
if ($probeChecklist.Count -ne 1) { throw ... }                           # Count 1
if ($probeChecklist[0] -ne 'sentinel') { throw ... }                     # value survives
$probeChecklist.Clear()
```

A failure aborts before Excel is started, exactly as `PRE` does.

---

## Mutable-container audit

Every helper in the three Windows scripts that builds or hands back a collection,
classified explicitly:

| Helper | Contract | Mechanism | Verdict |
|---|---|---|---|
| `New-Checklist` | ONE mutable ArrayList | `Write-Output -NoEnumerate` | **fixed** |
| `New-ReleaseLedger` | ONE `PSCustomObject` | plain `return` | safe — a PSCustomObject is not enumerable; its ArrayLists are *properties*, reached as `$rel.Lines.Add(...)`, never crossing a boundary as a return value |
| `Get-TransientFailures` | zero/one/many values | caller `@(...)` | already fixed in run 1 |
| `Get-TableBody` | zero/one/many **row** objects | `Write-RowObject` per row | already fixed after run 1 |
| `Write-RowObject` | one row object | `Write-Output -NoEnumerate` | correct by construction |
| `Get-IdColumnValues`, `Get-TableColumnNames`, `Get-PreExistingExcelPids` | zero/one/many values | caller `@(...)` | already fixed in run 1 |
| `Format-ReleaseLedger`, `Get-TrustAccessGuidance` | one string (`-join`) | plain `return` | safe — scalar |
| `Get-TableRowCount` | one integer (`.Count`) | plain `return` | safe — scalar |
| `Release-ComObjectSafe`, `Get-ExcelIdentity` | one `PSCustomObject` | plain `return` | safe |

**`New-Checklist` was the only factory returning an enumerable empty container.**
Nothing else was rewritten.

---

## Regressions

| Test | Proves |
|---|---|
| `test_46q` | the factory emits the ArrayList non-enumerated; `return (New-Object …ArrayList)` and `return $list` are both forbidden; it is still an ArrayList, because callers need `.Add()` |
| `test_46q1` | all 21 call sites receive the mutable object, and none is wrapped in `@(...)` — the two rules must not cross |
| `test_46q2` | the general rule: no function anywhere returns an enumerable container that would emit nothing when empty; `New-ReleaseLedger` is classified as safe-because-PSCustomObject |
| `test_46q3` | `PRE0` exists, runs before the first `Add-Check` and before Excel, uses no `Add-Check` itself, proves all five properties, and aborts on failure |

`test_46q`, `test_46q2` and `test_46q3` were each verified to fail against the
exact package used in Gate-B run 2.

---

## Status

Gate B is **not accepted**. `PRE` did not pass. No structural runtime scenario has
run. The corrected package is ready for one rerun.
