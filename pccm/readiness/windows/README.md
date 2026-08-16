# PCCM - Windows Excel Build-Path Smoke Test

**This is not PCCM.** It is a small, disposable readiness test that checks whether your
Windows + Excel machine can perform every operation the future PCCM "Stage B" bootstrap
will need.

It takes under a minute to run and writes a single text report that you send back.

> **Status: run 1 completed, run 2 pending.** On the first Windows execution TEST 01
> passed - Excel COM automation, version/build/bitness detection and process-identity
> capture all worked - and TEST 02 failed with a PowerShell/COM type-marshalling error
> that has been diagnosed and patched. This is a script defect, not an Excel or
> environment problem. **No Excel setting needs to change.** This revision adds
> substep-level reporting so run 2 pinpoints the exact statement if anything still fails.

---

## What it checks

| Test | Capability |
|------|------------|
| 01 | Excel COM automation can be instantiated; version / build / bitness detected |
| 02 | Workbook creation with a `SmokeTest` worksheet (reported as substeps 02.1-02.7) |
| 02N | Numeric cell write/read round-trip - the chart's source data |
| 03 | Save in genuine macro-enabled `.xlsm` format |
| 04 | `Workbook.VBProject` / `VBComponents` access |
| 05 | Standard module injection (`modSmokeTest`) |
| 06 | Class module injection (`clsSmokeTest`) |
| 07 | `ThisWorkbook` document-module code injection |
| 08 | Worksheet CodeName assignment (`shSmokeTest`) **and** worksheet document-module code injection |
| 09 | Shape/button creation with `OnAction` wired to a macro |
| 10 | Native `ChartObject` creation (requires 02N) |
| 11 | Macro execution through automation |
| 12 | Explicit COM release, then save, close and quit - Excel must exit without being forced |
| 13 | Reopen the saved `.xlsm` in a fresh Excel instance |
| 14 | Everything above survived the round trip, verified by reading persisted VBA **source**, not just component names |
| 15 | Explicit COM release, final close and quit - again without force |

Every test reports **PASS**, **FAIL**, **BLOCKED** or **SKIPPED**, with the underlying COM
error message and HRESULT where one occurred.

Three details worth knowing:

- **TEST 02 reports substeps.** Each of acquiring the Workbooks collection, creating the
  workbook, acquiring Worksheets, reducing to one sheet, acquiring the sheet, renaming it
  and the text write/read round-trip is reported separately with its own error handling,
  so a failure names the exact operation rather than "workbook creation failed".
- **Numeric cell writes are their own test (02N)**, because the chart needs numeric data
  but workbook creation does not. A number-marshalling problem is now reported as exactly
  that, and only the chart test is skipped as a result.
- **TEST 08 has two sub-checks** and both must pass. PCCM needs worksheet document-module
  code for its output-sheet activation logic, so proving CodeName assignment alone is not
  enough. The report records which of the two legitimate CodeName mechanisms succeeded.
- **TEST 14 reads the persisted source of all four code modules** and matches each against
  its own unique marker (`MODULE_MARKER_OK`, `CLASS_MARKER_OK`, `THISWORKBOOK_MARKER_OK`,
  `WORKSHEET_MODULE_MARKER_OK`). A module that survives as an empty shell counts as a
  failure.

If a capability fails, the tests that genuinely depend on it are reported as
`SKIPPED - prerequisite test failed: TEST nn` rather than producing a cascade of secondary
COM errors. The final verdict names the **first root cause**, not a downstream symptom.

---

## Safety design

- **No registry writes.** One Office key is *read* to detect Excel bitness.
- **No security changes.** The script never alters macro security, Trusted Locations, or
  "Trust access to the VBA project object model". If that setting is off, it detects the
  condition, reports `BLOCKED - VBA PROJECT TRUST ACCESS`, and prints the manual Excel UI
  path. It will not enable it for you.
- **No persisted Excel preferences are modified.** In particular
  `Application.SheetsInNewWorkbook` is deliberately left alone.
- **Files only in `.\smoke_output\`** beside the script.
- **The embedded test VBA** writes one value into one cell of its own disposable workbook.
  No file, network, shell or registry access. There is **no** `Workbook_Open`, **no**
  `Worksheet_Change`, and no other auto-running event handler anywhere in it.
- **Only the Excel instances the script starts are automated.** Excel windows you already
  have open are never automated, never closed and never terminated. If you have Excel open
  when you start, the script says so and leaves it alone.

### About cleanup

The script attempts a clean shutdown: it releases every COM reference it owns, leaf before
parent, before calling `Application.Quit()`, and then waits for the process to exit on its
own. It also includes a guarded emergency cleanup path that runs even if a test fails.

**It will not force-terminate an Excel process unless it can verify that the process
belongs to this smoke test** - all three of: the process id was derived from the script's
own `Application.Hwnd`, the process is still named `EXCEL`, and its start time is unchanged
since the script created it. If identity cannot be proven, the process is left running and
the report asks you to close that window manually.

Forcing a process never turns a failed graceful shutdown into a PASS. If Excel had to be
forced, TEST 12 or TEST 15 stays FAIL and the report says so - that outcome would tell us
the COM lifecycle needs more work before Stage B, which is exactly what this test exists to
find out.

---

## Instructions

### 1. Confirm Microsoft Excel Desktop is installed

This must be the real desktop Excel (Microsoft 365 / Office). Excel for the web and the
mobile apps have no COM automation and cannot be used.

### 2. Enable "Trust access to the VBA project object model"

In Excel:

```
File
  -> Options
    -> Trust Center
      -> Trust Center Settings...
        -> Macro Settings
          -> [x] Trust access to the VBA project object model
```

**Why this is needed:** the PCCM build imports VBA modules, class modules, `ThisWorkbook`
code and worksheet module code into the workbook, and assigns worksheet CodeNames. All of
that goes through the VBA project object model, which Office blocks from automation by
default. This is a per-user Excel setting on your machine and is required for the **build**
process only.

**Do not lower the general macro security level.** Only this one checkbox is needed. The
script will not change it under any circumstances - if it is off, the script stops and
tells you.

If your organisation's policy locks this setting, say so and we will change the build
strategy rather than work around the policy.

### 3. Close Excel completely

Close every Excel window before running, so the test starts from a clean state.
(If you forget, the script still runs and will not disturb your open workbooks.)

### 4. Open Windows PowerShell in this folder

In File Explorer, navigate to the folder containing `excel_smoke_test.ps1`, then either:

- type `powershell` into the Explorer address bar and press Enter, or
- Shift + right-click in the folder -> "Open PowerShell window here".

Use **Windows PowerShell 5.1** (`powershell.exe`, the blue icon), not PowerShell 7
(`pwsh`). Windows PowerShell has the more reliable COM interop and ships with Windows -
there is nothing to install.

### 5. Run the script

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\excel_smoke_test.ps1
```

`-ExecutionPolicy Bypass` here applies **only to this one process**. It does not change
your machine's execution policy, and nothing persists after the script exits.

The script prints each test result as it goes, then prints the full report.

### 6. Send back the report

```
smoke_output\smoke_test_report.txt
```

**Please send that text file, not a screenshot.** The report is the authoritative
diagnostic - it carries the exact COM error messages, HRESULT codes, Excel version, build
and bitness, COM release counts and cleanup notes that a screenshot would lose. Send a
screenshot only if the script fails so early that no report file is produced.

---

## Output files

Everything is created in `smoke_output\` beside the script:

| File | Purpose |
|------|---------|
| `smoke_test_report.txt` | **Return this one.** Always the current run. |
| `PCCM_Excel_COM_Smoke_Test.xlsm` | Disposable test workbook from the current run |
| `previous_<timestamp>\` | Each earlier run's report and workbook, moved aside automatically. Run 1's evidence is preserved here. |

A previous run's report and workbook are **never overwritten**. If files from an earlier
run are present, they are moved into a `previous_<timestamp>\` folder before the new run
starts, so a failed diagnostic is never lost. The file you return is always the one at the
fixed path `smoke_output\smoke_test_report.txt` - never one from a `previous_*` folder.

Delete the whole `smoke_output` folder whenever you like; nothing depends on it.

---

## Final verdict

The report ends with exactly one of:

| Verdict | Meaning |
|---------|---------|
| `READY FOR PCCM STAGE B` | Everything passed. The two-stage build is confirmed viable on this machine. |
| `BLOCKED - VBA PROJECT TRUST ACCESS` | Step 2 above was not completed, or is locked by policy. Enable it and re-run. |
| `BLOCKED - MACRO EXECUTION POLICY` | VBA could be injected but not executed. A macro security policy is intercepting execution, separately from VBProject trust. |
| `BLOCKED - EXCEL COM UNAVAILABLE` | Excel COM could not be instantiated at all. Desktop Excel may not be installed, or the installation is damaged. |
| `FAILED - <test>` | A specific capability failed. The report names the first root-cause test and includes the COM error. |

(The verdict text uses an em dash, not the hyphen shown in this table.)

If the result is anything other than `READY FOR PCCM STAGE B`, send the report anyway.
A precise failure is more useful than a retry, and we will adjust the build strategy to
whatever your environment actually supports.
