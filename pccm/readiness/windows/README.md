# PCCM - Windows Excel Build-Path Smoke Test

**This is not PCCM.** It is a small, disposable readiness test that proves your Windows +
Excel machine can perform every operation the future PCCM "Stage B" bootstrap will need.

It takes about 30 seconds to run and writes a single text report that you send back.

---

## What it proves

| Test | Capability |
|------|------------|
| 01 | Excel COM automation can be instantiated; version / build / bitness detected |
| 02 | Workbook creation with a `SmokeTest` worksheet |
| 03 | Save in genuine macro-enabled `.xlsm` format |
| 04 | `Workbook.VBProject` / `VBComponents` access |
| 05 | Standard module injection (`modSmokeTest`) |
| 06 | Class module injection (`clsSmokeTest`) |
| 07 | `ThisWorkbook` document-module code injection |
| 08 | Worksheet CodeName assignment (`Sheet1` -> `shSmokeTest`) |
| 09 | Shape/button creation with `OnAction` wired to a macro |
| 10 | Native `ChartObject` creation |
| 11 | Macro execution through automation |
| 12 | Save, close workbook, quit Excel cleanly |
| 13 | Reopen the saved `.xlsm` in a fresh Excel instance |
| 14 | Everything above survived the save/reopen round trip |
| 15 | Final clean close |

Every test reports **PASS**, **FAIL**, **BLOCKED** or **SKIPPED**, with the underlying COM
error message and HRESULT where one occurred.

---

## Safety

This script is deliberately conservative:

- It **never** edits the registry. It only *reads* one Office key to detect Excel bitness.
- It **never** changes macro security, Trusted Locations, or the "Trust access to the VBA
  project object model" setting. If that setting is off, it detects the condition, reports
  `BLOCKED - VBA PROJECT TRUST ACCESS`, and tells you the manual path to enable it.
- It **never** changes a persisted Excel preference.
- It writes files only into `.\smoke_output\` beside the script.
- The embedded test VBA writes one value (`MACRO_EXECUTED`) into one cell of its own
  disposable workbook. It performs no file, network, shell or registry activity, and adds
  **no** auto-running `Workbook_Open` handler.
- It automates **only the Excel instance it starts itself**, tracked by process id. Excel
  windows you already have open are never touched, and never closed. If you have Excel
  open when you start, the script says so and leaves it alone.
- Cleanup runs even when a test fails, so no orphaned `EXCEL.EXE` is left behind from this
  script.

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

**Why this is needed:** the PCCM build process imports VBA modules, class modules and
`ThisWorkbook` code into the workbook and assigns worksheet CodeNames. All of that goes
through the VBA project object model, which Office blocks from automation by default. This
is a per-user Excel setting for *your* machine and is required for the **build** only.

**Do not lower the general macro security level.** Only this one checkbox is needed. The
script will not change it for you under any circumstances - if it is off, the script stops
and tells you.

If your organisation's policy locks this setting, say so and we will change the build
strategy rather than work around the policy.

### 3. Close Excel completely

Close every Excel window before running, so the test starts from a clean state.
(If you forget, the script still runs safely and will not disturb your open workbooks.)

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

The script prints each test result as it goes and prints the full report at the end.

### 6. Send back the report

```
smoke_output\smoke_test_report.txt
```

**Please send that text file, not a screenshot.** The report is the authoritative
diagnostic - it contains the exact COM error messages, HRESULT codes and environment
details that a screenshot would lose. Only send a screenshot if the script fails so early
that no report file is produced.

---

## Output files

Everything is created in `smoke_output\` beside the script:

| File | Purpose |
|------|---------|
| `smoke_test_report.txt` | **Return this.** Full environment + per-test results + verdict |
| `PCCM_Excel_COM_Smoke_Test.xlsm` | Disposable test workbook. Keep it only if a test failed and we need to inspect it |

Delete the whole `smoke_output` folder whenever you like - nothing depends on it.

---

## Final verdict

The report ends with exactly one of:

| Verdict | Meaning |
|---------|---------|
| `READY FOR PCCM STAGE B` | Everything passed. The two-stage build is confirmed viable. |
| `BLOCKED - VBA PROJECT TRUST ACCESS` | Step 2 above was not completed, or is locked by policy. Enable it and re-run. |
| `BLOCKED - MACRO EXECUTION POLICY` | VBA could be injected but not executed. A macro security policy is intercepting execution separately from VBProject trust. |
| `BLOCKED - EXCEL COM UNAVAILABLE` | Excel COM could not be instantiated at all. Desktop Excel may not be installed or the installation is damaged. |
| `FAILED - <test>` | A specific capability failed. The report names the test and includes the COM error. |

(The verdict text uses an em dash, not the hyphen shown in this table.)

If the result is anything other than `READY FOR PCCM STAGE B`, send the report anyway.
A precise failure is more useful than a retry, and we will adjust the build strategy to
whatever your environment actually supports.
