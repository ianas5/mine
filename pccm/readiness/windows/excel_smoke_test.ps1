<#
    PCCM - Excel COM Build-Path Smoke Test
    ======================================

    PURPOSE
        Proves that the Windows + Excel machine can perform every operation that the
        future PCCM "Stage B" bootstrap will require:
            Excel COM automation, genuine .xlsm creation, VBA project access,
            standard/class/ThisWorkbook module injection, worksheet CodeName assignment,
            macro execution, Shape + OnAction, ChartObject, save/close/reopen persistence.

        This is a DISPOSABLE TEST. It is not PCCM and contains no PCCM logic.

    SAFETY
        - Never edits the registry (registry is READ-ONLY, for bitness detection only).
        - Never changes macro security, Trusted Locations, or "Trust access to the VBA
          project object model". If that setting is off, the script DETECTS and REPORTS it.
        - Never changes persisted Excel application preferences.
        - Creates files only in .\smoke_output\ beside this script.
        - The embedded test VBA touches one cell in its own workbook. No files, no network,
          no shell, no registry, no auto-run events.
        - Automates ONLY the Excel instance it creates, tracked by process id, and never
          touches Excel instances belonging to the user.

    USAGE
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\excel_smoke_test.ps1

    OUTPUT
        .\smoke_output\smoke_test_report.txt          <- return this file
        .\smoke_output\PCCM_Excel_COM_Smoke_Test.xlsm <- disposable artifact
#>

# Windows PowerShell 5.1 compatible. Do not use PS7-only syntax.
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = (Get-Location).Path }

$outDir     = Join-Path $scriptDir 'smoke_output'
$wbPath     = Join-Path $outDir    'PCCM_Excel_COM_Smoke_Test.xlsm'
$reportPath = Join-Path $outDir    'smoke_test_report.txt'

if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }
if (Test-Path -LiteralPath $wbPath) { Remove-Item -LiteralPath $wbPath -Force }

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
$xlOpenXMLWorkbookMacroEnabled = 52
$vbext_ct_StdModule            = 1
$vbext_ct_ClassModule          = 2
$vbext_ct_Document             = 100
$msoShapeRoundedRectangle      = 5
$xlColumnClustered             = 51

$MARKER_MODULE       = 'MODULE_MARKER_OK'
$MARKER_CLASS        = 'CLASS_MARKER_OK'
$MARKER_THISWORKBOOK = 'THISWORKBOOK_MARKER_OK'
$MARKER_MACRO        = 'MACRO_EXECUTED'
$TARGET_CODENAME     = 'shSmokeTest'
$SHAPE_NAME          = 'btnSmokeTest'
$MACRO_NAME          = 'RunSmokeMacro'

# Em dash used in the required verdict strings. Built at runtime so this script
# file stays pure ASCII: Windows PowerShell 5.1 decodes BOM-less UTF-8 as ANSI,
# which would otherwise corrupt the verdict text.
$DASH = [string][char]0x2014

# ---------------------------------------------------------------------------
# Result collection
# ---------------------------------------------------------------------------
$results = New-Object System.Collections.ArrayList
$env_info = [ordered]@{}

function Add-Result {
    param(
        [string]$Id,
        [string]$Name,
        [ValidateSet('PASS','FAIL','BLOCKED','SKIPPED')]
        [string]$Status,
        [string]$Detail = ''
    )
    $null = $results.Add([pscustomobject]@{
        Id     = $Id
        Name   = $Name
        Status = $Status
        Detail = $Detail
    })
    $colour = 'Gray'
    switch ($Status) {
        'PASS'    { $colour = 'Green'  }
        'FAIL'    { $colour = 'Red'    }
        'BLOCKED' { $colour = 'Yellow' }
        'SKIPPED' { $colour = 'DarkGray' }
    }
    Write-Host ("  {0}  {1,-7}  {2}" -f $Id, $Status, $Name) -ForegroundColor $colour
    if ($Detail) { Write-Host ("           {0}" -f $Detail) -ForegroundColor DarkGray }
}

function Format-Err {
    param($ErrorRecord)
    $parts = New-Object System.Collections.ArrayList
    try {
        $ex = $ErrorRecord.Exception
        if ($ex) {
            $null = $parts.Add($ex.Message.Trim())
            $null = $parts.Add('[type=' + $ex.GetType().FullName + ']')
            try {
                $hr = $ex.HResult
                if ($hr -ne 0) { $null = $parts.Add('[HRESULT=0x{0:X8}]' -f $hr) }
            } catch { }
            try {
                if ($ex -is [System.Runtime.InteropServices.COMException]) {
                    $null = $parts.Add('[COM ErrorCode=0x{0:X8}]' -f $ex.ErrorCode)
                }
            } catch { }
            if ($ex.InnerException) {
                $null = $parts.Add('[inner=' + $ex.InnerException.Message.Trim() + ']')
            }
        } else {
            $null = $parts.Add([string]$ErrorRecord)
        }
    } catch {
        $null = $parts.Add('<error while formatting exception>')
    }
    return ($parts -join ' ')
}

function Test-TrustAccessError {
    param([string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $false }
    $m = $Message.ToLowerInvariant()
    return ($m -match 'not trusted') -or
           ($m -match 'programmatic access') -or
           ($m -match 'visual basic project')
}

function Test-MacroPolicyError {
    param([string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $false }
    $m = $Message.ToLowerInvariant()
    return ($m -match 'macro') -and (($m -match 'disabl') -or ($m -match 'cannot run') -or ($m -match 'security'))
}

# ---------------------------------------------------------------------------
# Native helper: map the Excel window handle we created to its process id,
# so cleanup can never touch an Excel instance belonging to the user.
# ---------------------------------------------------------------------------
$haveNative = $false
try {
    Add-Type -Namespace PccmSmoke -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction Stop
    $haveNative = $true
} catch {
    $haveNative = $false
}

function Get-ExcelPid {
    param($ExcelApp, [int[]]$PreExistingPids)
    # Preferred: resolve our own instance's window handle to a pid.
    if ($haveNative) {
        try {
            $procId = 0
            $hwnd = [System.IntPtr]::new([int]$ExcelApp.Hwnd)
            $null = [PccmSmoke.Native]::GetWindowThreadProcessId($hwnd, [ref]$procId)
            if ($procId -gt 0) { return $procId }
        } catch { }
    }
    # Fallback: the excel process that appeared after we started.
    try {
        $now = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        $new = @($now | Where-Object { $PreExistingPids -notcontains $_ })
        if ($new.Count -eq 1) { return [int]$new[0] }
    } catch { }
    return 0
}

function Release-Com {
    param($Obj)
    if ($null -eq $Obj) { return }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
}

function Stop-OurExcel {
    param($ExcelApp, [int]$ExcelPid)
    if ($null -ne $ExcelApp) {
        try { $ExcelApp.DisplayAlerts = $false } catch { }
        try { $ExcelApp.Quit() } catch { }
        Release-Com $ExcelApp
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    # Only ever act on the pid we created ourselves.
    if ($ExcelPid -gt 0) {
        for ($i = 0; $i -lt 20; $i++) {
            $p = Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue
            if ($null -eq $p) { return $true }
            Start-Sleep -Milliseconds 500
        }
        try {
            $p = Get-Process -Id $ExcelPid -ErrorAction SilentlyContinue
            if ($null -ne $p -and $p.ProcessName -eq 'EXCEL') {
                Stop-Process -Id $ExcelPid -Force -ErrorAction SilentlyContinue
                return $true
            }
        } catch { }
    }
    return $true
}

# ---------------------------------------------------------------------------
# Embedded test VBA (literal here-strings: no PowerShell interpolation)
# ---------------------------------------------------------------------------
$vbaStdModule = @'
Option Explicit

' Disposable smoke-test module. Touches one cell in its own workbook only.

Public Function SmokeMarker() As String
    SmokeMarker = "MODULE_MARKER_OK"
End Function

Public Sub RunSmokeMacro()
    ThisWorkbook.Worksheets("SmokeTest").Range("B2").Value = "MACRO_EXECUTED"
End Sub
'@

$vbaClassModule = @'
Option Explicit

' Disposable smoke-test class module.

Private mLabel As String

Public Property Get MarkerValue() As String
    MarkerValue = "CLASS_MARKER_OK"
End Property

Public Property Let Label(ByVal Value As String)
    mLabel = Value
End Property

Public Property Get Label() As String
    Label = mLabel
End Property
'@

$vbaThisWorkbook = @'

' Disposable smoke-test marker injected into the ThisWorkbook document module.
' Deliberately contains NO Workbook_Open or other auto-running event handler.

Public Function ThisWorkbookMarker() As String
    ThisWorkbookMarker = "THISWORKBOOK_MARKER_OK"
End Function
'@

# ---------------------------------------------------------------------------
# Environment facts
# ---------------------------------------------------------------------------
Write-Host ''
Write-Host 'PCCM - Excel COM Build-Path Smoke Test' -ForegroundColor Cyan
Write-Host '======================================' -ForegroundColor Cyan
Write-Host ''

$env_info['Timestamp']        = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
$env_info['Script folder']    = $scriptDir
$env_info['Workbook path']    = $wbPath

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $env_info['Windows']      = ('{0} (build {1})' -f $os.Caption, $os.BuildNumber)
    $env_info['Windows arch'] = $os.OSArchitecture
} catch {
    try { $env_info['Windows'] = [System.Environment]::OSVersion.VersionString }
    catch { $env_info['Windows'] = 'UNKNOWN' }
    $env_info['Windows arch'] = 'UNKNOWN'
}

$env_info['PowerShell version'] = $PSVersionTable.PSVersion.ToString()
try { $env_info['PowerShell edition'] = [string]$PSVersionTable.PSEdition } catch { $env_info['PowerShell edition'] = 'Desktop' }
$env_info['PowerShell 64-bit']  = [System.Environment]::Is64BitProcess.ToString()
$env_info['OS 64-bit']          = [System.Environment]::Is64BitOperatingSystem.ToString()

$env_info['Excel version']  = 'NOT DETECTED'
$env_info['Excel build']    = 'NOT DETECTED'
$env_info['Excel path']     = 'NOT DETECTED'
$env_info['Excel bitness']  = 'UNCONFIRMED'

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
$excel      = $null
$excel2     = $null
$wb         = $null
$wb2        = $null
$excelPid   = 0
$excel2Pid  = 0
$abort      = $false
$verdict    = $null

$preExistingPids = @()
try { $preExistingPids = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch { }
if ($preExistingPids.Count -gt 0) {
    Write-Host ("  NOTE: {0} Excel process(es) already running. They will not be touched." -f $preExistingPids.Count) -ForegroundColor Yellow
    Write-Host ''
}

try {

    # =======================================================================
    # TEST 01 - Excel COM
    # =======================================================================
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false
        $excel.ScreenUpdating = $false
        $excelPid = Get-ExcelPid -ExcelApp $excel -PreExistingPids $preExistingPids

        try { $env_info['Excel version'] = [string]$excel.Version } catch { }
        try { $env_info['Excel build']   = [string]$excel.Build }   catch { }
        try { $env_info['Excel path']    = [string]$excel.Path }    catch { }

        # Bitness - READ-ONLY probes. Nothing is written to the registry.
        $bitness = 'UNCONFIRMED'
        try {
            $cfg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
            if ($cfg.Platform) {
                if ($cfg.Platform -eq 'x64') { $bitness = '64-bit (ClickToRun Platform=x64)' }
                elseif ($cfg.Platform -eq 'x86') { $bitness = '32-bit (ClickToRun Platform=x86)' }
            }
        } catch { }
        if ($bitness -eq 'UNCONFIRMED') {
            try {
                $p = [string]$excel.Path
                if ($p -match '(?i)Program Files \(x86\)') { $bitness = '32-bit (inferred from install path)' }
                elseif ($p -match '(?i)Program Files' -and [System.Environment]::Is64BitOperatingSystem) { $bitness = '64-bit (inferred from install path)' }
            } catch { }
        }
        $env_info['Excel bitness'] = $bitness

        Add-Result '01' 'Excel COM instantiation' 'PASS' ("Excel {0} build {1}; pid {2}; bitness {3}" -f $env_info['Excel version'], $env_info['Excel build'], $excelPid, $bitness)
    } catch {
        Add-Result '01' 'Excel COM instantiation' 'BLOCKED' (Format-Err $_)
        $verdict = "BLOCKED $DASH EXCEL COM UNAVAILABLE"
        $abort = $true
    }

    # =======================================================================
    # TEST 02 - Workbook creation
    # =======================================================================
    if (-not $abort) {
        try {
            $wb = $excel.Workbooks.Add()
            # Do NOT touch Application.SheetsInNewWorkbook - it is a persisted user preference.
            while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = 'SmokeTest'

            $ws.Range('A1').Value2 = 'PCCM disposable Excel COM smoke test'
            $ws.Range('A2').Value2 = 'Macro marker cell ->'
            $ws.Range('B2').Value2 = 'PENDING'
            # Harmless sample data for the chart (kept clear of the marker cell).
            $ws.Range('D1').Value2 = 'Year'
            $ws.Range('E1').Value2 = 'Value'
            for ($i = 1; $i -le 5; $i++) {
                $ws.Cells.Item($i + 1, 4).Value2 = 2028 + $i - 1
                $ws.Cells.Item($i + 1, 5).Value2 = $i * 10
            }
            Add-Result '02' 'Workbook creation + SmokeTest sheet' 'PASS' ("Sheets={0}; name='{1}'" -f $wb.Worksheets.Count, $ws.Name)
        } catch {
            Add-Result '02' 'Workbook creation + SmokeTest sheet' 'FAIL' (Format-Err $_)
            $abort = $true
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 02 workbook creation" }
        }
    }

    # =======================================================================
    # TEST 03 - XLSM save
    # =======================================================================
    if (-not $abort) {
        try {
            $wb.SaveAs($wbPath, $xlOpenXMLWorkbookMacroEnabled)
            if (-not (Test-Path -LiteralPath $wbPath)) { throw "SaveAs reported success but the file does not exist: $wbPath" }
            $sz = (Get-Item -LiteralPath $wbPath).Length
            Add-Result '03' 'Save as macro-enabled .xlsm' 'PASS' ("FileFormat={0}; {1} bytes" -f $wb.FileFormat, $sz)
        } catch {
            Add-Result '03' 'Save as macro-enabled .xlsm' 'FAIL' (Format-Err $_)
            $abort = $true
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 03 xlsm save" }
        }
    }

    # =======================================================================
    # TEST 04 - VBA project access
    # =======================================================================
    $vbproj = $null
    if (-not $abort) {
        try {
            $vbproj = $wb.VBProject
            if ($null -eq $vbproj) { throw 'Workbook.VBProject returned null without raising an error.' }
            $componentCount = $vbproj.VBComponents.Count
            Add-Result '04' 'VBProject / VBComponents access' 'PASS' ("VBComponents.Count={0}" -f $componentCount)
        } catch {
            $msg = Format-Err $_
            if (Test-TrustAccessError $msg) {
                Add-Result '04' 'VBProject / VBComponents access' 'BLOCKED' ("Trust access to the VBA project object model is DISABLED. " + $msg)
                $verdict = "BLOCKED $DASH VBA PROJECT TRUST ACCESS"
            } else {
                Add-Result '04' 'VBProject / VBComponents access' 'FAIL' $msg
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 04 VBProject access" }
            }
            $abort = $true
        }
    }

    # =======================================================================
    # TEST 05 - Standard module
    # =======================================================================
    if (-not $abort) {
        try {
            $c = $vbproj.VBComponents.Add($vbext_ct_StdModule)
            $c.Name = 'modSmokeTest'
            $c.CodeModule.AddFromString($vbaStdModule)
            $lines = $c.CodeModule.CountOfLines
            Add-Result '05' 'Standard module injection (modSmokeTest)' 'PASS' ("CodeModule.CountOfLines={0}" -f $lines)
        } catch {
            Add-Result '05' 'Standard module injection (modSmokeTest)' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 05 standard module injection" }
        }
    }

    # =======================================================================
    # TEST 06 - Class module
    # =======================================================================
    if (-not $abort) {
        try {
            $c = $vbproj.VBComponents.Add($vbext_ct_ClassModule)
            $c.Name = 'clsSmokeTest'
            $c.CodeModule.AddFromString($vbaClassModule)
            $lines = $c.CodeModule.CountOfLines
            Add-Result '06' 'Class module injection (clsSmokeTest)' 'PASS' ("CodeModule.CountOfLines={0}" -f $lines)
        } catch {
            Add-Result '06' 'Class module injection (clsSmokeTest)' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 06 class module injection" }
        }
    }

    # =======================================================================
    # TEST 07 - ThisWorkbook document module
    # =======================================================================
    if (-not $abort) {
        try {
            # Locate the workbook document module without relying on the English
            # component name: it is the Type=100 component that is not a worksheet.
            $sheetCodeNames = @()
            foreach ($sh in $wb.Worksheets) { $sheetCodeNames += [string]$sh.CodeName }

            $twComp = $null
            foreach ($comp in $vbproj.VBComponents) {
                if ($comp.Type -eq $vbext_ct_Document -and ($sheetCodeNames -notcontains [string]$comp.Name)) {
                    $twComp = $comp
                    break
                }
            }
            if ($null -eq $twComp) { throw 'Could not locate the ThisWorkbook document module among the VBComponents.' }

            $twComp.CodeModule.AddFromString($vbaThisWorkbook)
            $code = $twComp.CodeModule.Lines(1, $twComp.CodeModule.CountOfLines)
            if ($code -notmatch [regex]::Escape($MARKER_THISWORKBOOK)) { throw 'Marker text not found in ThisWorkbook code module after injection.' }
            Add-Result '07' 'ThisWorkbook document-module injection' 'PASS' ("component='{0}'; marker present" -f $twComp.Name)
        } catch {
            Add-Result '07' 'ThisWorkbook document-module injection' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 07 ThisWorkbook injection" }
        }
    }

    # =======================================================================
    # TEST 08 - Worksheet CodeName
    # =======================================================================
    if (-not $abort) {
        try {
            $ws = $wb.Worksheets.Item('SmokeTest')
            $oldCodeName = [string]$ws.CodeName
            # Worksheet.CodeName is read-only at runtime. The supported mechanism is the
            # document component's "_CodeName" design-time property.
            $comp = $vbproj.VBComponents.Item($oldCodeName)
            $prop = $null
            try { $prop = $comp.Properties.Item('_CodeName') } catch { }
            if ($null -eq $prop) {
                foreach ($p in $comp.Properties) { if ([string]$p.Name -eq '_CodeName') { $prop = $p; break } }
            }
            if ($null -eq $prop) { throw 'Document component does not expose the _CodeName property.' }
            $prop.Value = $TARGET_CODENAME

            $newCodeName = [string]$ws.CodeName
            if ($newCodeName -ne $TARGET_CODENAME) { throw ("CodeName is '{0}' after assignment, expected '{1}'." -f $newCodeName, $TARGET_CODENAME) }
            Add-Result '08' 'Worksheet CodeName assignment' 'PASS' ("'{0}' -> '{1}' (verified via Worksheet.CodeName)" -f $oldCodeName, $newCodeName)
        } catch {
            Add-Result '08' 'Worksheet CodeName assignment' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 08 worksheet CodeName assignment" }
        }
    }

    # =======================================================================
    # TEST 09 - Shape / button with OnAction
    # =======================================================================
    if (-not $abort) {
        try {
            $ws = $wb.Worksheets.Item('SmokeTest')
            $shp = $ws.Shapes.AddShape($msoShapeRoundedRectangle, 20, 70, 150, 34)
            $shp.Name = $SHAPE_NAME
            try { $shp.TextFrame2.TextRange.Text = 'Run Smoke Macro' } catch { }
            $shp.OnAction = $MACRO_NAME
            $readBack = [string]$shp.OnAction
            if ($readBack -notlike ("*{0}*" -f $MACRO_NAME)) { throw ("OnAction read back as '{0}'." -f $readBack) }
            Add-Result '09' 'Shape creation + OnAction assignment' 'PASS' ("name='{0}'; OnAction='{1}'" -f $shp.Name, $readBack)
        } catch {
            Add-Result '09' 'Shape creation + OnAction assignment' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 09 shape / OnAction" }
        }
    }

    # =======================================================================
    # TEST 10 - Chart
    # =======================================================================
    if (-not $abort) {
        try {
            $ws = $wb.Worksheets.Item('SmokeTest')
            $co = $ws.ChartObjects().Add(280, 20, 320, 200)
            $co.Name = 'chtSmokeTest'
            $co.Chart.SetSourceData($ws.Range('D1:E6'))
            $co.Chart.ChartType = $xlColumnClustered
            $count = $ws.ChartObjects().Count
            Add-Result '10' 'ChartObject creation' 'PASS' ("ChartObjects.Count={0}; name='{1}'; type={2}" -f $count, $co.Name, $co.Chart.ChartType)
        } catch {
            Add-Result '10' 'ChartObject creation' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 10 chart creation" }
        }
    }

    # =======================================================================
    # TEST 11 - Macro execution
    # =======================================================================
    if (-not $abort) {
        $runErr1 = ''
        $runErr2 = ''
        $ran = $false
        try {
            $excel.Run($MACRO_NAME)
            $ran = $true
        } catch {
            $runErr1 = Format-Err $_
            # Retry with a workbook-qualified macro name in case of name resolution.
            try {
                $qualified = "'" + $wb.Name + "'!" + $MACRO_NAME
                $excel.Run($qualified)
                $ran = $true
            } catch {
                $runErr2 = Format-Err $_
            }
        }

        if ($ran) {
            try {
                $ws = $wb.Worksheets.Item('SmokeTest')
                $val = [string]$ws.Range('B2').Value2
                if ($val -eq $MARKER_MACRO) {
                    Add-Result '11' 'Macro execution (RunSmokeMacro)' 'PASS' ("SmokeTest!B2='{0}'" -f $val)
                } else {
                    Add-Result '11' 'Macro execution (RunSmokeMacro)' 'FAIL' ("Macro ran without error but SmokeTest!B2='{0}', expected '{1}'." -f $val, $MARKER_MACRO)
                    if (-not $verdict) { $verdict = "FAILED $DASH TEST 11 macro marker not written" }
                }
            } catch {
                Add-Result '11' 'Macro execution (RunSmokeMacro)' 'FAIL' (Format-Err $_)
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 11 macro verification" }
            }
        } else {
            $combined = ($runErr1 + ' || retry: ' + $runErr2)
            if ((Test-MacroPolicyError $runErr1) -or (Test-MacroPolicyError $runErr2)) {
                Add-Result '11' 'Macro execution (RunSmokeMacro)' 'BLOCKED' ('Macro execution appears blocked by Excel macro security policy (distinct from VBProject trust, which passed at TEST 04). ' + $combined)
                if (-not $verdict) { $verdict = "BLOCKED $DASH MACRO EXECUTION POLICY" }
            } else {
                Add-Result '11' 'Macro execution (RunSmokeMacro)' 'FAIL' $combined
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 11 macro execution" }
            }
        }
    }

    # =======================================================================
    # TEST 12 - Save, close, quit
    # =======================================================================
    if (-not $abort) {
        try {
            $wb.Save()
            $wb.Close($false)
            Release-Com $wb
            $wb = $null
            $excel.Quit()
            Release-Com $excel
            $excel = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            $gone = $true
            if ($excelPid -gt 0) {
                for ($i = 0; $i -lt 20; $i++) {
                    if ($null -eq (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { break }
                    Start-Sleep -Milliseconds 500
                }
                if ($null -ne (Get-Process -Id $excelPid -ErrorAction SilentlyContinue)) { $gone = $false }
            }
            if ($gone) {
                Add-Result '12' 'Save, close workbook, quit Excel' 'PASS' ("pid {0} exited cleanly" -f $excelPid)
            } else {
                Add-Result '12' 'Save, close workbook, quit Excel' 'FAIL' ("Excel pid {0} did not exit after Quit(); it will be force-stopped during cleanup." -f $excelPid)
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 12 Excel did not exit cleanly" }
            }
        } catch {
            Add-Result '12' 'Save, close workbook, quit Excel' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 12 save/close/quit" }
        }
    }

    # =======================================================================
    # TEST 13 - Reopen in a fresh Excel instance
    # =======================================================================
    if (-not $abort) {
        try {
            $preExisting2 = @()
            try { $preExisting2 = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch { }
            $excel2 = New-Object -ComObject Excel.Application
            $excel2.Visible = $false
            $excel2.DisplayAlerts = $false
            $excel2Pid = Get-ExcelPid -ExcelApp $excel2 -PreExistingPids $preExisting2
            $wb2 = $excel2.Workbooks.Open($wbPath)
            Add-Result '13' 'Reopen saved .xlsm in a fresh instance' 'PASS' ("pid {0}; opened '{1}'" -f $excel2Pid, $wb2.Name)
        } catch {
            Add-Result '13' 'Reopen saved .xlsm in a fresh instance' 'FAIL' (Format-Err $_)
            $abort = $true
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 13 reopen" }
        }
    }

    # =======================================================================
    # TEST 14 - Persistence verification
    # =======================================================================
    if (-not $abort) {
        $checks = New-Object System.Collections.ArrayList
        $allOk = $true

        function Add-Check {
            param([string]$Label, [bool]$Ok, [string]$Info)
            $null = $checks.Add(("      [{0}] {1}: {2}" -f $(if ($Ok) {'PASS'} else {'FAIL'}), $Label, $Info))
            return $Ok
        }

        try {
            # file format
            $ff = -1
            try { $ff = [int]$wb2.FileFormat } catch { }
            if (-not (Add-Check 'File format is macro-enabled' ($ff -eq $xlOpenXMLWorkbookMacroEnabled) ("FileFormat={0} (expected {1})" -f $ff, $xlOpenXMLWorkbookMacroEnabled))) { $allOk = $false }

            # VBA project
            $vb2 = $null
            $vbOk = $false
            try { $vb2 = $wb2.VBProject; $vbOk = ($null -ne $vb2) } catch { }
            if (-not (Add-Check 'VBProject accessible after reopen' $vbOk '')) { $allOk = $false }

            $names = @()
            if ($vbOk) { foreach ($c in $vb2.VBComponents) { $names += [string]$c.Name } }

            if (-not (Add-Check 'modSmokeTest present' ($names -contains 'modSmokeTest') ("components: " + ($names -join ', ')))) { $allOk = $false }
            if (-not (Add-Check 'clsSmokeTest present' ($names -contains 'clsSmokeTest') '')) { $allOk = $false }

            # ThisWorkbook marker
            $twOk = $false
            $twInfo = 'not located'
            if ($vbOk) {
                $sheetCodeNames2 = @()
                foreach ($sh in $wb2.Worksheets) { $sheetCodeNames2 += [string]$sh.CodeName }
                foreach ($comp in $vb2.VBComponents) {
                    if ($comp.Type -eq $vbext_ct_Document -and ($sheetCodeNames2 -notcontains [string]$comp.Name)) {
                        $n = $comp.CodeModule.CountOfLines
                        if ($n -gt 0) {
                            $code = $comp.CodeModule.Lines(1, $n)
                            $twOk = ($code -match [regex]::Escape($MARKER_THISWORKBOOK))
                            $twInfo = ("component='{0}'; lines={1}" -f $comp.Name, $n)
                        }
                        break
                    }
                }
            }
            if (-not (Add-Check 'ThisWorkbook marker persisted' $twOk $twInfo)) { $allOk = $false }

            # CodeName
            $ws2 = $null
            $cnOk = $false
            $cn = ''
            try {
                $ws2 = $wb2.Worksheets.Item('SmokeTest')
                $cn = [string]$ws2.CodeName
                $cnOk = ($cn -eq $TARGET_CODENAME)
            } catch { }
            if (-not (Add-Check 'Worksheet CodeName persisted' $cnOk ("CodeName='{0}' (expected '{1}')" -f $cn, $TARGET_CODENAME))) { $allOk = $false }

            # Shape + OnAction
            $shpOk = $false
            $shpInfo = 'shape not found'
            try {
                $s = $ws2.Shapes.Item($SHAPE_NAME)
                $oa = [string]$s.OnAction
                $shpOk = ($oa -like ("*{0}*" -f $MACRO_NAME))
                $shpInfo = ("OnAction='{0}'" -f $oa)
            } catch { $shpInfo = (Format-Err $_) }
            if (-not (Add-Check 'Shape present with OnAction intact' $shpOk $shpInfo)) { $allOk = $false }

            # Chart
            $chOk = $false
            $chInfo = ''
            try {
                $n = [int]$ws2.ChartObjects().Count
                $chOk = ($n -ge 1)
                $chInfo = ("ChartObjects.Count={0}" -f $n)
            } catch { $chInfo = (Format-Err $_) }
            if (-not (Add-Check 'ChartObject persisted' $chOk $chInfo)) { $allOk = $false }

            # Macro marker
            $mkOk = $false
            $mkVal = ''
            try {
                $mkVal = [string]$ws2.Range('B2').Value2
                $mkOk = ($mkVal -eq $MARKER_MACRO)
            } catch { }
            if (-not (Add-Check 'Macro-written marker persisted in B2' $mkOk ("B2='{0}' (expected '{1}')" -f $mkVal, $MARKER_MACRO))) { $allOk = $false }

            $detail = "`r`n" + ($checks -join "`r`n")
            if ($allOk) {
                Add-Result '14' 'Persistence verification after reopen' 'PASS' $detail
            } else {
                Add-Result '14' 'Persistence verification after reopen' 'FAIL' $detail
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 14 persistence verification" }
            }
        } catch {
            Add-Result '14' 'Persistence verification after reopen' 'FAIL' ((Format-Err $_) + "`r`n" + ($checks -join "`r`n"))
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 14 persistence verification" }
        }
    }

    # =======================================================================
    # TEST 15 - Final close
    # =======================================================================
    if (-not $abort) {
        try {
            if ($null -ne $wb2) { $wb2.Close($false); Release-Com $wb2; $wb2 = $null }
            if ($null -ne $excel2) { $excel2.Quit(); Release-Com $excel2; $excel2 = $null }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            $gone = $true
            if ($excel2Pid -gt 0) {
                for ($i = 0; $i -lt 20; $i++) {
                    if ($null -eq (Get-Process -Id $excel2Pid -ErrorAction SilentlyContinue)) { break }
                    Start-Sleep -Milliseconds 500
                }
                if ($null -ne (Get-Process -Id $excel2Pid -ErrorAction SilentlyContinue)) { $gone = $false }
            }
            if ($gone) {
                Add-Result '15' 'Final close and quit' 'PASS' ("pid {0} exited cleanly" -f $excel2Pid)
            } else {
                Add-Result '15' 'Final close and quit' 'FAIL' ("Excel pid {0} did not exit after Quit(); it will be force-stopped during cleanup." -f $excel2Pid)
                if (-not $verdict) { $verdict = "FAILED $DASH TEST 15 Excel did not exit cleanly" }
            }
        } catch {
            Add-Result '15' 'Final close and quit' 'FAIL' (Format-Err $_)
            if (-not $verdict) { $verdict = "FAILED $DASH TEST 15 final close" }
        }
    }

    # Mark anything never reached as SKIPPED so the report is complete.
    $allTests = @(
        @('01','Excel COM instantiation'),
        @('02','Workbook creation + SmokeTest sheet'),
        @('03','Save as macro-enabled .xlsm'),
        @('04','VBProject / VBComponents access'),
        @('05','Standard module injection (modSmokeTest)'),
        @('06','Class module injection (clsSmokeTest)'),
        @('07','ThisWorkbook document-module injection'),
        @('08','Worksheet CodeName assignment'),
        @('09','Shape creation + OnAction assignment'),
        @('10','ChartObject creation'),
        @('11','Macro execution (RunSmokeMacro)'),
        @('12','Save, close workbook, quit Excel'),
        @('13','Reopen saved .xlsm in a fresh instance'),
        @('14','Persistence verification after reopen'),
        @('15','Final close and quit')
    )
    $seen = @($results | Select-Object -ExpandProperty Id)
    foreach ($t in $allTests) {
        if ($seen -notcontains $t[0]) {
            Add-Result $t[0] $t[1] 'SKIPPED' 'Not reached because an earlier test blocked or failed.'
        }
    }

}
catch {
    # Any unexpected top-level failure.
    Add-Result 'XX' 'Unhandled script error' 'FAIL' ((Format-Err $_) + " || stack: " + $_.ScriptStackTrace)
    if (-not $verdict) { $verdict = "FAILED $DASH unhandled script error" }
}
finally {
    # -----------------------------------------------------------------------
    # Cleanup. Runs whatever happened above. Only touches instances we created.
    # -----------------------------------------------------------------------
    Write-Host ''
    Write-Host '  Cleaning up...' -ForegroundColor DarkGray
    try { if ($null -ne $wb)  { $wb.Close($false)  } } catch { }
    try { if ($null -ne $wb2) { $wb2.Close($false) } } catch { }
    Release-Com $wb
    Release-Com $wb2
    Stop-OurExcel -ExcelApp $excel  -ExcelPid $excelPid  | Out-Null
    Stop-OurExcel -ExcelApp $excel2 -ExcelPid $excel2Pid | Out-Null

    # -----------------------------------------------------------------------
    # Verdict
    # -----------------------------------------------------------------------
    if (-not $verdict) {
        $bad = @($results | Where-Object { $_.Status -eq 'FAIL' })
        $blk = @($results | Where-Object { $_.Status -eq 'BLOCKED' })
        if ($blk.Count -gt 0) {
            $verdict = ("FAILED $DASH TEST {0} {1}" -f $blk[0].Id, $blk[0].Name)
        } elseif ($bad.Count -gt 0) {
            $verdict = ("FAILED $DASH TEST {0} {1}" -f $bad[0].Id, $bad[0].Name)
        } else {
            $verdict = 'READY FOR PCCM STAGE B'
        }
    }

    # -----------------------------------------------------------------------
    # Report
    # -----------------------------------------------------------------------
    $sb = New-Object System.Text.StringBuilder
    $null = $sb.AppendLine('===============================================================================')
    $null = $sb.AppendLine(' PCCM - EXCEL COM BUILD-PATH SMOKE TEST REPORT')
    $null = $sb.AppendLine(' Disposable readiness test for the PCCM Stage B bootstrap. Not a PCCM build.')
    $null = $sb.AppendLine('===============================================================================')
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('ENVIRONMENT')
    $null = $sb.AppendLine('-----------')
    foreach ($k in $env_info.Keys) {
        $null = $sb.AppendLine(('  {0,-22}: {1}' -f $k, $env_info[$k]))
    }
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('TEST RESULTS')
    $null = $sb.AppendLine('------------')
    foreach ($r in $results) {
        $null = $sb.AppendLine(('  TEST {0}  {1,-8} {2}' -f $r.Id, $r.Status, $r.Name))
        if ($r.Detail) {
            foreach ($line in ($r.Detail -split "`r?`n")) {
                if ($line.Trim().Length -gt 0) { $null = $sb.AppendLine(('           {0}' -f $line.Trim())) }
            }
        }
    }
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('SUMMARY')
    $null = $sb.AppendLine('-------')
    $null = $sb.AppendLine(('  PASS    : {0}' -f @($results | Where-Object { $_.Status -eq 'PASS' }).Count))
    $null = $sb.AppendLine(('  FAIL    : {0}' -f @($results | Where-Object { $_.Status -eq 'FAIL' }).Count))
    $null = $sb.AppendLine(('  BLOCKED : {0}' -f @($results | Where-Object { $_.Status -eq 'BLOCKED' }).Count))
    $null = $sb.AppendLine(('  SKIPPED : {0}' -f @($results | Where-Object { $_.Status -eq 'SKIPPED' }).Count))
    $null = $sb.AppendLine('')

    if ($verdict -eq "BLOCKED $DASH VBA PROJECT TRUST ACCESS") {
        $null = $sb.AppendLine('REQUIRED MANUAL ACTION')
        $null = $sb.AppendLine('----------------------')
        $null = $sb.AppendLine('  Programmatic access to the VBA project is disabled. Enable it manually in Excel:')
        $null = $sb.AppendLine('')
        $null = $sb.AppendLine('      File')
        $null = $sb.AppendLine('        -> Options')
        $null = $sb.AppendLine('          -> Trust Center')
        $null = $sb.AppendLine('            -> Trust Center Settings...')
        $null = $sb.AppendLine('              -> Macro Settings')
        $null = $sb.AppendLine('                -> [x] Trust access to the VBA project object model')
        $null = $sb.AppendLine('')
        $null = $sb.AppendLine('  Then close Excel completely and run this script again.')
        $null = $sb.AppendLine('  Do NOT lower the general macro security level. Only this one checkbox is needed,')
        $null = $sb.AppendLine('  and only because the PCCM build imports VBA modules into the workbook.')
        $null = $sb.AppendLine('  This script did not and will not change the setting for you.')
        $null = $sb.AppendLine('')
    }

    $null = $sb.AppendLine('===============================================================================')
    $null = $sb.AppendLine((' FINAL VERDICT: {0}' -f $verdict))
    $null = $sb.AppendLine('===============================================================================')

    $reportText = $sb.ToString()
    try {
        Set-Content -LiteralPath $reportPath -Value $reportText -Encoding UTF8
    } catch {
        Write-Host ('  Could not write report file: ' + $_.Exception.Message) -ForegroundColor Red
    }

    Write-Host ''
    Write-Host $reportText
    Write-Host ''
    Write-Host ('  Report written to: {0}' -f $reportPath) -ForegroundColor Cyan
    Write-Host '  Please return that file.' -ForegroundColor Cyan
    Write-Host ''
}
