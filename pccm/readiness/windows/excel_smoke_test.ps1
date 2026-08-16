<#
    PCCM - Excel COM Build-Path Smoke Test  (hardened revision)
    ===========================================================

    PURPOSE
        Proves that this Windows + Excel machine can perform every operation the future
        PCCM "Stage B" bootstrap will require:
            Excel COM automation, genuine .xlsm creation, VBA project access,
            standard module / class module / ThisWorkbook / worksheet document-module
            code injection, worksheet CodeName assignment, macro execution,
            Shape + OnAction, ChartObject, and save/close/reopen persistence.

        This is a DISPOSABLE TEST. It is not PCCM and contains no PCCM logic.

    SAFETY DESIGN
        - Does not write to the registry. One Office key is READ for bitness detection.
        - Does not change macro security, Trusted Locations, or "Trust access to the VBA
          project object model". If that setting is off, the script detects and reports it.
        - Does not change any persisted Excel application preference.
        - Writes files only inside .\smoke_output\ beside this script.
        - The embedded test VBA writes one value into one cell of its own disposable
          workbook. No file, network, shell or registry access. No auto-running events,
          no Workbook_Open, no Worksheet_Change.
        - Automates only the Excel instance it starts itself.
        - Emergency force-stop is permitted ONLY for a process whose identity is verified
          three ways: pid derived from our own Application.Hwnd, process name still EXCEL,
          and process StartTime unchanged. If identity cannot be proven, the process is
          left alone and the user is asked to close it manually.

    RUN 2 INSTRUMENTATION
        TEST 02 is split into reported substeps, numeric cell writes are a separate test
        (02N), and the COM release path reports registered/released counts, completion
        state and any exception instead of swallowing it.

    COM LIFECYCLE
        Every COM object this script obtains is either released immediately after last use
        or registered on a per-instance stack released LIFO (leaf before parent) before
        Application.Quit(). Garbage collection is a final aid only, never the ownership
        strategy. COM collections are walked by index, never with foreach, because COM
        enumerators retain RCWs of their own.

    USAGE
        powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\excel_smoke_test.ps1

    OUTPUT
        .\smoke_output\smoke_test_report.txt            <- return this file
        .\smoke_output\PCCM_Excel_COM_Smoke_Test.xlsm   <- disposable artifact
        .\smoke_output\previous_<timestamp>\            <- prior run's files, if any
#>

# Windows PowerShell 5.1 compatible. Do not use PS7-only syntax.
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Paths. A previous run's artifacts are preserved, never silently destroyed.
# ---------------------------------------------------------------------------
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrEmpty($scriptDir)) { $scriptDir = (Get-Location).Path }

$outDir     = Join-Path $scriptDir 'smoke_output'
$wbPath     = Join-Path $outDir    'PCCM_Excel_COM_Smoke_Test.xlsm'
$reportPath = Join-Path $outDir    'smoke_test_report.txt'

if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir | Out-Null }

$archivedTo = ''
$priorFiles = @()
if (Test-Path -LiteralPath $wbPath)     { $priorFiles += $wbPath }
if (Test-Path -LiteralPath $reportPath) { $priorFiles += $reportPath }
if ($priorFiles.Count -gt 0) {
    $stamp   = (Get-Date).ToString('yyyyMMdd_HHmmss')
    $prevDir = Join-Path $outDir ('previous_' + $stamp)
    New-Item -ItemType Directory -Path $prevDir | Out-Null
    foreach ($f in $priorFiles) {
        try { Move-Item -LiteralPath $f -Destination $prevDir -Force } catch { }
    }
    $archivedTo = $prevDir
}

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
$MARKER_WORKSHEET    = 'WORKSHEET_MODULE_MARKER_OK'
$MARKER_MACRO        = 'MACRO_EXECUTED'
$TARGET_CODENAME     = 'shSmokeTest'
$SHAPE_NAME          = 'btnSmokeTest'
$MACRO_NAME          = 'RunSmokeMacro'
$STD_MODULE_NAME     = 'modSmokeTest'
$CLS_MODULE_NAME     = 'clsSmokeTest'

# Em dash used in the required verdict strings. Built at runtime so this script file
# stays pure ASCII: Windows PowerShell 5.1 decodes BOM-less UTF-8 as ANSI, which would
# otherwise corrupt the verdict text.
$DASH = [string][char]0x2014

# ---------------------------------------------------------------------------
# Result collection and capability flags
# ---------------------------------------------------------------------------
$results      = New-Object System.Collections.ArrayList
$cleanupNotes = New-Object System.Collections.ArrayList
$env_info     = [ordered]@{}
$cap          = @{}

function Add-Result {
    param(
        [string]$Id,
        [string]$Name,
        [ValidateSet('PASS','FAIL','BLOCKED','SKIPPED')]
        [string]$Status,
        [string]$Detail = ''
    )
    $null = $script:results.Add([pscustomobject]@{
        Id = $Id; Name = $Name; Status = $Status; Detail = $Detail
    })
    $colour = 'Gray'
    switch ($Status) {
        'PASS'    { $colour = 'Green'    }
        'FAIL'    { $colour = 'Red'      }
        'BLOCKED' { $colour = 'Yellow'   }
        'SKIPPED' { $colour = 'DarkGray' }
    }
    Write-Host ("  TEST {0}  {1,-8} {2}" -f $Id, $Status, $Name) -ForegroundColor $colour
    if ($Detail) {
        foreach ($ln in ($Detail -split "`r?`n")) {
            if ($ln.Trim().Length -gt 0) { Write-Host ("            {0}" -f $ln.Trim()) -ForegroundColor DarkGray }
        }
    }
}

function Set-Cap { param([string]$Key, [bool]$Value) $script:cap[$Key] = $Value }
function Get-Cap { param([string]$Key) if ($script:cap.ContainsKey($Key)) { return [bool]$script:cap[$Key] } return $false }

# Returns $true when the test may run. Otherwise records SKIPPED naming the root cause,
# so a failed prerequisite does not generate downstream COM noise.
function Test-Prereq {
    param([string]$Id, [string]$Name, [string[]]$Requires)
    foreach ($r in $Requires) {
        if (-not (Get-Cap $r)) {
            Add-Result $Id $Name 'SKIPPED' ("prerequisite test failed: TEST {0}" -f $r)
            return $false
        }
    }
    return $true
}

function Add-CleanupNote { param([string]$Text) $null = $script:cleanupNotes.Add($Text) }

function Format-Err {
    param($ErrorRecord)
    $parts = New-Object System.Collections.ArrayList
    try {
        $ex = $ErrorRecord.Exception
        if ($ex) {
            $null = $parts.Add($ex.Message.Trim())
            $null = $parts.Add('[type=' + $ex.GetType().FullName + ']')
            try { $hr = $ex.HResult; if ($hr -ne 0) { $null = $parts.Add('[HRESULT=0x{0:X8}]' -f $hr) } } catch { }
            try {
                if ($ex -is [System.Runtime.InteropServices.COMException]) {
                    $null = $parts.Add('[COM ErrorCode=0x{0:X8}]' -f $ex.ErrorCode)
                }
            } catch { }
            if ($ex.InnerException) { $null = $parts.Add('[inner=' + $ex.InnerException.Message.Trim() + ']') }
        } else {
            $null = $parts.Add([string]$ErrorRecord)
        }
    } catch { $null = $parts.Add('<error while formatting exception>') }
    return ($parts -join ' ')
}

function Test-TrustAccessError {
    param([string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $false }
    $m = $Message.ToLowerInvariant()
    return ($m -match 'not trusted') -or ($m -match 'programmatic access') -or ($m -match 'visual basic project')
}

function Test-MacroPolicyError {
    param([string]$Message)
    if ([string]::IsNullOrEmpty($Message)) { return $false }
    $m = $Message.ToLowerInvariant()
    return ($m -match 'macro') -and (($m -match 'disabl') -or ($m -match 'cannot run') -or ($m -match 'security'))
}

# ===========================================================================
# COM lifecycle
# ===========================================================================
# Single release helper. FinalReleaseComObject is used deliberately: every RCW here is
# created and owned exclusively by this short-lived script, held by exactly one variable,
# and released once. Under those conditions final release is the reliable way to drive the
# RCW count to zero. The bulk pass de-duplicates by reference identity so a shared RCW can
# never be final-released twice, and no object is released while another live variable
# still points at it.
function Remove-ComObjectRef {
    param($Obj)
    if ($null -eq $Obj) { return $false }
    try { if (-not [System.Runtime.InteropServices.Marshal]::IsComObject($Obj)) { return $false } } catch { return $false }
    try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Obj); return $true } catch { return $false }
}

$comStack1        = New-Object System.Collections.ArrayList
$comStack2        = New-Object System.Collections.ArrayList
$comStackActive   = $comStack1
$comTrackWarnings = New-Object System.Collections.ArrayList
$rel1             = $null
$rel2             = $null

# Diagnostic readers. These read ArrayList.Count only and never touch a COM object, so
# they cannot alter COM ownership.
function Get-ComStackCount {
    param([int]$Which)
    try {
        if ($Which -eq 2) { return [int]$script:comStack2.Count }
        return [int]$script:comStack1.Count
    } catch { return -1 }
}
function Get-ComDiag {
    return ("COM stack 1 count: {0}; COM stack 2 count: {1}" -f (Get-ComStackCount 1), (Get-ComStackCount 2))
}

# Registers a long-lived COM reference for LIFO release. Returns nothing: returning a COM
# object from a PowerShell function can cause the pipeline to enumerate COM collections.
function Register-ComRef {
    param($Obj, [string]$Label)
    if ($null -eq $Obj) {
        $null = $script:comTrackWarnings.Add(("Register-ComRef received null for '{0}'" -f $Label))
        return
    }
    if ($null -eq $script:comStackActive) {
        $null = $script:comTrackWarnings.Add(("no active COM stack when registering '{0}'" -f $Label))
        return
    }
    $before = [int]$script:comStackActive.Count
    $null = $script:comStackActive.Add([pscustomobject]@{ Obj = $Obj; Label = $Label })
    $after = [int]$script:comStackActive.Count
    if ($after -ne ($before + 1)) {
        $null = $script:comTrackWarnings.Add(("stack did not grow registering '{0}': before={1} after={2}" -f $Label, $before, $after))
    }
}

# Releases the stack LIFO. Acquisition order is parent -> child, so reverse order is
# leaf -> parent as required.
# ---------------------------------------------------------------------------
# Two-phase COM release.
#
# INVARIANT: no algorithmic operation may touch an RCW after FinalReleaseComObject
# has been called on it. FinalReleaseComObject separates the RCW from its underlying
# COM object; any later use - including reading it back out of a collection to compare
# references - raises InvalidComObjectException. The previous single-pass algorithm
# added each object to a $released list and then final-released it, so the very next
# iteration's ReferenceEquals comparison touched a separated RCW and aborted the whole
# traversal after the first object. That is corrected here by completing every
# de-duplication decision BEFORE any release occurs.
# ---------------------------------------------------------------------------

# PHASE A - build the release plan while every RCW is still valid.
# Walks the stack LIFO, so the resulting plan is already ordered leaf-before-parent.
# ReferenceEquals is safe here precisely because nothing has been released yet.
# Releases nothing.
function Build-ComReleasePlan {
    param($Stack)
    $plan       = New-Object System.Collections.ArrayList
    $dupCount   = 0
    $skipped    = New-Object System.Collections.ArrayList
    for ($i = $Stack.Count - 1; $i -ge 0; $i--) {
        $entry = $null; $o = $null; $lbl = '<unknown>'
        try {
            $entry = $Stack[$i]
            $lbl   = [string]$entry.Label
            $o     = $entry.Obj
        } catch {
            $null = $skipped.Add(("<stack entry {0} unreadable: {1}>" -f $i, $_.Exception.Message))
            continue
        }
        if ($null -eq $o) { continue }
        $isDup = $false
        for ($j = 0; $j -lt $plan.Count; $j++) {
            if ([Object]::ReferenceEquals($plan[$j].Obj, $o)) { $isDup = $true; break }
        }
        if ($isDup) { $dupCount++ }
        else { $null = $plan.Add([pscustomobject]@{ Label = $lbl; Obj = $o }) }
    }
    return ([pscustomobject]@{
        Plan           = $plan
        DuplicateCount = $dupCount
        Skipped        = @($skipped.ToArray())
    })
}

# PHASE B lives inside this wrapper. It always returns a result object normally, even
# when an individual release fails, so a partial ReleasedCount can never be discarded.
function Invoke-ComStackRelease {
    param($Stack, [string]$Label)
    $res = [pscustomobject]@{
        Label               = $Label
        RegisteredCount     = -1
        UniqueCount         = 0
        DuplicateCount      = 0
        ReleasedCount       = 0
        FailedReleaseLabels = @()
        Completed           = $false
        Error               = ''
    }
    $failed = New-Object System.Collections.ArrayList

    # ---- PHASE A: plan only, nothing released ----------------------------
    $plan = $null
    try {
        $res.RegisteredCount = [int]$Stack.Count
        $built = Build-ComReleasePlan $Stack
        $plan  = $built.Plan
        $res.UniqueCount    = [int]$plan.Count
        $res.DuplicateCount = [int]$built.DuplicateCount
        foreach ($sk in @($built.Skipped)) { $null = $failed.Add($sk) }
    } catch {
        $res.Error = 'release plan build failed: ' + (Format-Err $_)
        $null = $failed.Add('<plan build failed - nothing was released>')
        $res.FailedReleaseLabels = @($failed.ToArray())
        return $res
    }

    # ---- PHASE B: release only, never inspect a released RCW again -------
    try {
        for ($k = 0; $k -lt $plan.Count; $k++) {
            $lbl = '<unknown>'
            $obj = $null
            try {
                $item = $plan[$k]
                $lbl  = [string]$item.Label
                $obj  = $item.Obj
                $item.Obj = $null      # the plan drops its reference BEFORE release
            } catch {
                $null = $failed.Add(("<plan entry {0} unreadable>" -f $k))
                continue
            }
            $okRel = $false
            try { $okRel = [bool](Remove-ComObjectRef $obj) } catch { $okRel = $false }
            $obj = $null               # sole remaining reference dropped; never touched again
            if ($okRel) { $res.ReleasedCount = $res.ReleasedCount + 1 }
            else { $null = $failed.Add($lbl) }
        }
        $res.Completed = $true
    } catch {
        # Completed=$false means the traversal itself did not finish. ReleasedCount
        # still holds the true partial count because it is mutated in place.
        $res.Completed = $false
        $res.Error = 'release traversal failed: ' + (Format-Err $_)
    }

    $res.FailedReleaseLabels = @($failed.ToArray())

    # The stack is emptied without reading any entry, so no released RCW is touched.
    try { $Stack.Clear() } catch { }
    return $res
}

function Format-ReleaseResult {
    param($Res, [string]$Name)
    if ($null -eq $Res) { return ("  {0}: release never attempted (instance was not created)." -f $Name) }
    $clean = ($Res.Completed -and @($Res.FailedReleaseLabels).Count -eq 0 -and -not $Res.Error)
    if ($clean) {
        return ("  {0}: registered {1}, unique {2}, duplicates {3}, released {4}, traversal completed OK" -f `
                $Name, $Res.RegisteredCount, $Res.UniqueCount, $Res.DuplicateCount, $Res.ReleasedCount)
    }
    $out = New-Object System.Collections.ArrayList
    $null = $out.Add(("  {0}: registered {1}, unique {2}, duplicates {3}, released {4}, traversal completed={5}" -f `
                     $Name, $Res.RegisteredCount, $Res.UniqueCount, $Res.DuplicateCount, $Res.ReleasedCount, $Res.Completed))
    if ($Res.Error) { $null = $out.Add(("      exception: {0}" -f $Res.Error)) }
    if (@($Res.FailedReleaseLabels).Count -gt 0) {
        $null = $out.Add(("      release FAILED for: {0}" -f (@($Res.FailedReleaseLabels) -join ', ')))
    }
    return ($out -join "`r`n")
}

# Substep reporter, so one large try/catch can never hide the failing statement.
function Add-SubStep {
    param($List, [string]$Id, [string]$Text, [string]$Status, [string]$Info = '')
    $line = ("      {0,-7} [{1,-7}] {2}" -f $Id, $Status, $Text)
    if ($Info) { $line = $line + ' :: ' + $Info }
    $null = $List.Add($line)
    $c = 'DarkGray'
    if ($Status -eq 'FAIL') { $c = 'Red' } elseif ($Status -eq 'PASS') { $c = 'DarkGreen' }
    Write-Host $line -ForegroundColor $c
}

# Cell access helpers, so Range RCWs never accumulate through chained property access.
# Text and numeric writes are DELIBERATELY separate and strongly typed. The previous
# helper took an untyped $Value, which handed a boxed Int32 to the COM binder for the
# numeric writes; Excel stores every number as a Double, so [double] is both the correct
# and the safest type to marshal.
function Set-CellText {
    param($Sheet, [string]$Address, [string]$Value)
    $rng = $Sheet.Range($Address)
    try { $rng.Value2 = $Value } finally { [void](Remove-ComObjectRef $rng) }
}

function Get-CellText {
    param($Sheet, [string]$Address)
    $rng = $Sheet.Range($Address)
    try { return [string]$rng.Value2 } finally { [void](Remove-ComObjectRef $rng) }
}

# Three legitimate numeric write mechanisms, tried in order. Returns the name of the one
# that succeeded so the report records it, mirroring the TEST 08 CodeName pattern.
function Set-CellNumber {
    param($Sheet, [string]$Address, [double]$Value)
    $rng = $Sheet.Range($Address)
    $attempts = New-Object System.Collections.ArrayList
    try {
        try { $rng.Value2 = $Value; return 'Range.Value2 <- [double]' }
        catch { $null = $attempts.Add('Range.Value2 <- [double] :: ' + (Format-Err $_)) }

        try { $rng.Value = $Value; return 'Range.Value <- [double]' }
        catch { $null = $attempts.Add('Range.Value <- [double] :: ' + (Format-Err $_)) }

        $txt = $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
        try { $rng.Formula = $txt; return 'Range.Formula <- invariant [string]' }
        catch { $null = $attempts.Add('Range.Formula <- invariant [string] :: ' + (Format-Err $_)) }

        throw ("all numeric write mechanisms failed for " + $Address + " -> " + ($attempts -join ' || '))
    } finally {
        [void](Remove-ComObjectRef $rng)
    }
}

function Get-CellNumber {
    param($Sheet, [string]$Address)
    $rng = $Sheet.Range($Address)
    try {
        $v = $rng.Value2
        if ($null -eq $v) { return [double]::NaN }
        return [double]$v
    } catch {
        return [double]::NaN
    } finally {
        [void](Remove-ComObjectRef $rng)
    }
}

# Reads a VBComponent's source. Returns a plain string; the CodeModule RCW is released here.
function Get-ComponentSource {
    param($Component)
    $cm = $null
    try {
        $cm = $Component.CodeModule
        $n = [int]$cm.CountOfLines
        if ($n -le 0) { return '' }
        return [string]$cm.Lines(1, $n)
    } catch {
        return ''
    } finally {
        if ($null -ne $cm) { [void](Remove-ComObjectRef $cm) }
    }
}

# ===========================================================================
# Process identity
# ===========================================================================
$haveNative = $false
try {
    Add-Type -Namespace PccmSmoke -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction Stop
    $haveNative = $true
} catch { $haveNative = $false }

# Captures a strong identity for the Excel instance WE created.
#   Source = 'HWND'     -> pid derived from our own Application.Hwnd. Force-stop permitted.
#   Source = 'FALLBACK' -> diagnostic only. Force-stop NOT permitted.
function Get-ExcelIdentity {
    param($ExcelApp, [int[]]$PreExistingPids)
    $id = [pscustomobject]@{
        ProcessId      = 0
        Source         = 'UNKNOWN'
        HasStartTime   = $false
        StartTimeTicks = [long]0
        StartTimeText  = 'unavailable'
        ProcessName    = ''
    }
    if ($haveNative) {
        try {
            $procId = 0
            $hwnd = New-Object System.IntPtr ([int]$ExcelApp.Hwnd)
            $null = [PccmSmoke.Native]::GetWindowThreadProcessId($hwnd, [ref]$procId)
            if ($procId -gt 0) { $id.ProcessId = $procId; $id.Source = 'HWND' }
        } catch { }
    }
    if ($id.ProcessId -le 0) {
        try {
            $now = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $new = @($now | Where-Object { $PreExistingPids -notcontains $_ })
            if ($new.Count -eq 1) { $id.ProcessId = [int]$new[0]; $id.Source = 'FALLBACK' }
        } catch { }
    }
    if ($id.ProcessId -gt 0) {
        try {
            $p = Get-Process -Id $id.ProcessId -ErrorAction Stop
            $id.ProcessName = [string]$p.ProcessName
            try {
                $st = $p.StartTime
                $id.HasStartTime   = $true
                $id.StartTimeTicks = [long]$st.Ticks
                $id.StartTimeText  = $st.ToString('yyyy-MM-dd HH:mm:ss.fff')
            } catch { }
        } catch { }
    }
    return $id
}

# Force-stop is allowed only when all three identity facts still hold.
function Test-IsOurExcelProcess {
    param($Identity)
    if ($null -eq $Identity) { return $false }
    if ($Identity.ProcessId -le 0) { return $false }
    if ($Identity.Source -ne 'HWND') { return $false }
    if (-not $Identity.HasStartTime) { return $false }
    $p = Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue
    if ($null -eq $p) { return $false }
    if ($p.ProcessName -notmatch '^(?i)excel$') { return $false }
    try { if ([long]$p.StartTime.Ticks -ne [long]$Identity.StartTimeTicks) { return $false } } catch { return $false }
    return $true
}

# Returns $true only when the process is confirmed gone. Never force-stops.
function Wait-ExcelExit {
    param($Identity, [int]$TimeoutSeconds = 20)
    if ($null -eq $Identity -or $Identity.ProcessId -le 0) { return $false }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 400
    }
    return ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue))
}

# Emergency path only. Never called from a test; never converts a failed graceful close
# into a PASS.
function Invoke-EmergencyExcelCleanup {
    param($Identity, [string]$Label)
    if ($null -eq $Identity -or $Identity.ProcessId -le 0) {
        return ("{0}: no process identity was captured; nothing to clean up." -f $Label)
    }
    if ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue)) {
        return ("{0}: pid {1} already exited." -f $Label, $Identity.ProcessId)
    }
    if (Test-IsOurExcelProcess $Identity) {
        try {
            Stop-Process -Id $Identity.ProcessId -Force -ErrorAction Stop
            return ("{0}: EMERGENCY - force-stopped smoke-test Excel pid {1} (identity verified: HWND-derived, name EXCEL, StartTime {2})." -f $Label, $Identity.ProcessId, $Identity.StartTimeText)
        } catch {
            return ("{0}: force-stop of pid {1} FAILED: {2}. Please close that Excel window manually." -f $Label, $Identity.ProcessId, $_.Exception.Message)
        }
    }
    return ("{0}: pid {1} is still running but its identity could NOT be verified (source={2}, startTimeCaptured={3}). NOT terminated. Please close the leftover Excel window manually." -f $Label, $Identity.ProcessId, $Identity.Source, $Identity.HasStartTime)
}

# ===========================================================================
# Embedded test VBA (literal here-strings: no PowerShell interpolation).
# No events, no auto-run handlers, no external access.
# ===========================================================================
$vbaStdModule = @'
Option Explicit

' Disposable smoke-test standard module.
' Touches one cell in its own workbook. Nothing else.

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
' Deliberately contains NO Workbook_Open or any other auto-running event handler.

Public Function ThisWorkbookMarker() As String
    ThisWorkbookMarker = "THISWORKBOOK_MARKER_OK"
End Function
'@

$vbaWorksheetModule = @'

' Disposable smoke-test marker injected into the SmokeTest worksheet document module.
' This proves worksheet document-module code injection, which PCCM needs for output-sheet
' activation logic. Deliberately contains NO Worksheet_Change, NO Worksheet_Activate and
' no other event handler.

Public Function WorksheetMarker() As String
    WorksheetMarker = "WORKSHEET_MODULE_MARKER_OK"
End Function
'@

# ===========================================================================
# Environment
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Excel COM Build-Path Smoke Test' -ForegroundColor Cyan
Write-Host '=====================================' -ForegroundColor Cyan
Write-Host ''

$env_info['Timestamp']     = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss zzz')
$env_info['Script folder'] = $scriptDir
$env_info['Workbook path'] = $wbPath
if ($archivedTo) { $env_info['Previous run archived to'] = $archivedTo }

try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $env_info['Windows']      = ('{0} (build {1})' -f $os.Caption, $os.BuildNumber)
    $env_info['Windows arch'] = [string]$os.OSArchitecture
} catch {
    try { $env_info['Windows'] = [System.Environment]::OSVersion.VersionString } catch { $env_info['Windows'] = 'UNKNOWN' }
    $env_info['Windows arch'] = 'UNKNOWN'
}

$env_info['PowerShell version'] = $PSVersionTable.PSVersion.ToString()
try { $env_info['PowerShell edition'] = [string]$PSVersionTable.PSEdition } catch { $env_info['PowerShell edition'] = 'Desktop' }
$env_info['PowerShell 64-bit'] = [System.Environment]::Is64BitProcess.ToString()
$env_info['OS 64-bit']         = [System.Environment]::Is64BitOperatingSystem.ToString()
$env_info['Excel version']     = 'NOT DETECTED'
$env_info['Excel build']       = 'NOT DETECTED'
$env_info['Excel path']        = 'NOT DETECTED'
$env_info['Excel bitness']     = 'UNCONFIRMED'
$env_info['Excel pid source']  = 'n/a'

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------
$excel = $null; $workbooks = $null; $wb = $null; $worksheets = $null; $ws = $null
$vbproj = $null; $vbcomps = $null; $shapes = $null; $shp = $null
$chartObjects = $null; $co = $null; $cht = $null
$excel2 = $null; $workbooks2 = $null; $wb2 = $null; $worksheets2 = $null; $ws2 = $null
$vbproj2 = $null; $vbcomps2 = $null

$id1 = $null
$id2 = $null
$inst1Finished = $false
$inst2Finished = $false
$codeNameMechanism = 'none'

$preExistingPids = @()
try { $preExistingPids = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch { }
if ($preExistingPids.Count -gt 0) {
    Write-Host ("  NOTE: {0} Excel process(es) are already running. They are never automated, closed or terminated by this script." -f $preExistingPids.Count) -ForegroundColor Yellow
    Write-Host ''
}

try {

    # =======================================================================
    # TEST 01 - Excel COM
    # =======================================================================
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        $excel.ScreenUpdating = $false

        $id1 = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExistingPids
        $env_info['Excel pid source'] = ('{0} (pid {1}, StartTime {2})' -f $id1.Source, $id1.ProcessId, $id1.StartTimeText)

        try { $env_info['Excel version'] = [string]$excel.Version } catch { }
        try { $env_info['Excel build']   = [string]$excel.Build }   catch { }
        try { $env_info['Excel path']    = [string]$excel.Path }    catch { }

        # Bitness: READ-ONLY probes. Nothing is written to the registry.
        $bitness = 'UNCONFIRMED'
        try {
            $cfg = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction Stop
            if ($cfg.Platform -eq 'x64')      { $bitness = '64-bit (ClickToRun Platform=x64)' }
            elseif ($cfg.Platform -eq 'x86')  { $bitness = '32-bit (ClickToRun Platform=x86)' }
        } catch { }
        if ($bitness -eq 'UNCONFIRMED') {
            try {
                $p = [string]$excel.Path
                if ($p -match '(?i)Program Files \(x86\)') { $bitness = '32-bit (inferred from install path)' }
                elseif ($p -match '(?i)Program Files' -and [System.Environment]::Is64BitOperatingSystem) { $bitness = '64-bit (inferred from install path)' }
            } catch { }
        }
        $env_info['Excel bitness'] = $bitness

        Add-Result '01' 'Excel COM instantiation' 'PASS' ("Excel {0} build {1}; bitness {2}; pid {3} via {4}" -f $env_info['Excel version'], $env_info['Excel build'], $bitness, $id1.ProcessId, $id1.Source)
        Set-Cap '01' $true
    } catch {
        Add-Result '01' 'Excel COM instantiation' 'BLOCKED' (Format-Err $_)
        Set-Cap '01' $false
    }

    # =======================================================================
    # TEST 02 - Workbook and worksheet creation
    #   Split into reported substeps. Each substep has its own try/catch so the exact
    #   failing statement is named instead of being hidden by one large handler.
    #   Numeric cell writes are deliberately NOT here - see TEST 02N.
    # =======================================================================
    if (Test-Prereq '02' 'Workbook and worksheet creation' @('01')) {
        $steps = New-Object System.Collections.ArrayList
        $failedAt = ''
        $failErr  = ''

        # --- 02.1 Workbooks collection acquired -----------------------------
        if ($failedAt -eq '') {
            try {
                $workbooks = $excel.Workbooks
                Register-ComRef $workbooks 'Application.Workbooks'
                Add-SubStep $steps '02.1' 'Workbooks collection acquired' 'PASS' ("stack1={0}" -f (Get-ComStackCount 1))
            } catch {
                $failedAt = '02.1'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.1' 'Workbooks collection acquired' 'FAIL' $failErr
            }
        }

        # --- 02.2 Workbook created ------------------------------------------
        if ($failedAt -eq '') {
            try {
                $wb = $workbooks.Add()
                Register-ComRef $wb 'Workbook'
                Add-SubStep $steps '02.2' 'Workbook created' 'PASS' ("stack1={0}" -f (Get-ComStackCount 1))
            } catch {
                $failedAt = '02.2'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.2' 'Workbook created' 'FAIL' $failErr
            }
        }

        # --- 02.3 Worksheets collection acquired ----------------------------
        if ($failedAt -eq '') {
            try {
                $worksheets = $wb.Worksheets
                Register-ComRef $worksheets 'Workbook.Worksheets'
                Add-SubStep $steps '02.3' 'Worksheets collection acquired' 'PASS' ("stack1={0}" -f (Get-ComStackCount 1))
            } catch {
                $failedAt = '02.3'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.3' 'Worksheets collection acquired' 'FAIL' $failErr
            }
        }

        # --- 02.4 Reduced to a single worksheet -----------------------------
        # Application.SheetsInNewWorkbook is a persisted user preference and is NOT touched.
        if ($failedAt -eq '') {
            $initialSheets = -1
            try {
                $initialSheets = [int]$worksheets.Count
                while ([int]$worksheets.Count -gt 1) {
                    $extra = $worksheets.Item([int]$worksheets.Count)
                    $extra.Delete()
                    [void](Remove-ComObjectRef $extra)
                    $extra = $null
                }
                Add-SubStep $steps '02.4' 'Reduced to a single worksheet' 'PASS' ("initial={0}; final={1}" -f $initialSheets, [int]$worksheets.Count)
            } catch {
                $failedAt = '02.4'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.4' 'Reduced to a single worksheet' 'FAIL' (("initial={0}; " -f $initialSheets) + $failErr)
            }
        }

        # --- 02.5 Worksheet acquired ----------------------------------------
        if ($failedAt -eq '') {
            try {
                $ws = $worksheets.Item(1)
                Register-ComRef $ws 'Worksheet(SmokeTest)'
                Add-SubStep $steps '02.5' 'Worksheet acquired via Worksheets.Item(1)' 'PASS' ("stack1={0}" -f (Get-ComStackCount 1))
            } catch {
                $failedAt = '02.5'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.5' 'Worksheet acquired via Worksheets.Item(1)' 'FAIL' $failErr
            }
        }

        # --- 02.6 Worksheet renamed -----------------------------------------
        if ($failedAt -eq '') {
            try {
                $ws.Name = 'SmokeTest'
                $nameBack = [string]$ws.Name
                if ($nameBack -ne 'SmokeTest') { throw ("rename read back as '{0}'" -f $nameBack) }
                Add-SubStep $steps '02.6' 'Worksheet renamed to SmokeTest' 'PASS' ("Worksheet.Name='{0}'" -f $nameBack)
            } catch {
                $failedAt = '02.6'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.6' 'Worksheet renamed to SmokeTest' 'FAIL' $failErr
            }
        }

        # --- 02.7 Text cell write/read round-trip ---------------------------
        if ($failedAt -eq '') {
            $cell = ''
            try {
                $cell = 'A1'; Set-CellText $ws 'A1' 'PCCM disposable Excel COM smoke test'
                $cell = 'A2'; Set-CellText $ws 'A2' 'Macro marker cell ->'
                $cell = 'B2'; Set-CellText $ws 'B2' 'PENDING'
                $cell = 'B2 read-back'
                $rb = Get-CellText $ws 'B2'
                if ($rb -ne 'PENDING') { throw ("text round-trip mismatch: B2='{0}', expected 'PENDING'" -f $rb) }
                Add-SubStep $steps '02.7' 'Text cell write/read round-trip' 'PASS' ("A1, A2, B2 written; B2 read back as '{0}'" -f $rb)
            } catch {
                $failedAt = '02.7'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02.7' 'Text cell write/read round-trip' 'FAIL' (("at {0}: " -f $cell) + $failErr)
            }
        }

        $detail = "`r`n" + ($steps -join "`r`n")
        if ($failedAt -eq '') {
            Add-Result '02' 'Workbook and worksheet creation' 'PASS' $detail
            Set-Cap '02' $true
        } else {
            Add-Result '02' 'Workbook and worksheet creation' 'FAIL' ($detail + "`r`n      failing substep: " + $failedAt + "`r`n      " + (Get-ComDiag))
            Set-Cap '02' $false
        }
    }

    # =======================================================================
    # TEST 02N - Numeric cell write/read round-trip (chart source data)
    #   Separated from TEST 02 because the chart needs numeric data but workbook
    #   creation does not. A numeric-marshalling problem must not be reported as a
    #   failure of the workbook creation capability.
    #   Id '02N' sorts between '02' and '03', so report order is preserved.
    # =======================================================================
    if (Test-Prereq '02N' 'Numeric cell write/read round-trip (chart source data)' @('02')) {
        $steps = New-Object System.Collections.ArrayList
        $failedAt = ''
        $failErr  = ''
        $numMech  = 'none'

        # --- 02N.1 single numeric probe -------------------------------------
        if ($failedAt -eq '') {
            try {
                $numMech = Set-CellNumber $ws 'H1' 12345
                Add-SubStep $steps '02N.1' 'Numeric probe written to H1' 'PASS' ("mechanism: {0}" -f $numMech)
            } catch {
                $failedAt = '02N.1'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02N.1' 'Numeric probe written to H1' 'FAIL' $failErr
            }
        }

        # --- 02N.2 numeric read-back ----------------------------------------
        if ($failedAt -eq '') {
            try {
                $v = Get-CellNumber $ws 'H1'
                if ([double]::IsNaN($v)) { throw 'H1 read back as non-numeric / empty' }
                if ([math]::Abs($v - 12345) -gt 0.000000001) { throw ("H1 read back as {0}, expected 12345" -f $v) }
                $rngc = $ws.Range('H1')
                try { $null = $rngc.ClearContents() } finally { [void](Remove-ComObjectRef $rngc) }
                Add-SubStep $steps '02N.2' 'Numeric probe read back and cleared' 'PASS' ("value={0}" -f $v)
            } catch {
                $failedAt = '02N.2'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02N.2' 'Numeric probe read back and cleared' 'FAIL' $failErr
            }
        }

        # --- 02N.3 chart source block D1:E6 ---------------------------------
        if ($failedAt -eq '') {
            $cell = ''
            try {
                $cell = 'D1/E1'
                Set-CellText $ws 'D1' 'Year'
                Set-CellText $ws 'E1' 'Value'
                for ($i = 1; $i -le 5; $i++) {
                    $cell = ('D{0}' -f ($i + 1)); [void](Set-CellNumber $ws $cell (2027 + $i))
                    $cell = ('E{0}' -f ($i + 1)); [void](Set-CellNumber $ws $cell ($i * 10))
                }
                $cell = 'D6 read-back'
                $chk = Get-CellNumber $ws 'D6'
                if ([double]::IsNaN($chk) -or [math]::Abs($chk - 2032) -gt 0.000000001) { throw ("D6 read back as {0}, expected 2032" -f $chk) }
                Add-SubStep $steps '02N.3' 'Chart source block D1:E6 populated' 'PASS' ("D6 read back as {0}" -f $chk)
            } catch {
                $failedAt = '02N.3'; $failErr = (Format-Err $_)
                Add-SubStep $steps '02N.3' 'Chart source block D1:E6 populated' 'FAIL' (("at {0}: " -f $cell) + $failErr)
            }
        }

        $detail = "`r`n" + ($steps -join "`r`n")
        if ($failedAt -eq '') {
            Add-Result '02N' 'Numeric cell write/read round-trip (chart source data)' 'PASS' $detail
            Set-Cap '02N' $true
        } else {
            Add-Result '02N' 'Numeric cell write/read round-trip (chart source data)' 'FAIL' ($detail + "`r`n      failing substep: " + $failedAt + "`r`n      " + (Get-ComDiag))
            Set-Cap '02N' $false
        }
    }

    # =======================================================================
    # TEST 03 - XLSM save
    # =======================================================================
    if (Test-Prereq '03' 'Save as macro-enabled .xlsm' @('02')) {
        try {
            $wb.SaveAs($wbPath, $xlOpenXMLWorkbookMacroEnabled)
            if (-not (Test-Path -LiteralPath $wbPath)) { throw "SaveAs reported success but the file does not exist: $wbPath" }
            $sz = (Get-Item -LiteralPath $wbPath).Length
            Add-Result '03' 'Save as macro-enabled .xlsm' 'PASS' ("FileFormat={0}; {1} bytes" -f [int]$wb.FileFormat, $sz)
            Set-Cap '03' $true
        } catch {
            Add-Result '03' 'Save as macro-enabled .xlsm' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '03' $false
        }
    }

    # =======================================================================
    # TEST 04 - VBA project access
    # =======================================================================
    if (Test-Prereq '04' 'VBProject / VBComponents access' @('03')) {
        try {
            $vbproj = $wb.VBProject
            if ($null -eq $vbproj) { throw 'Workbook.VBProject returned null without raising an error.' }
            Register-ComRef $vbproj 'Workbook.VBProject'

            $vbcomps = $vbproj.VBComponents
            if ($null -eq $vbcomps) { throw 'VBProject.VBComponents returned null.' }
            Register-ComRef $vbcomps 'VBProject.VBComponents'

            $n = [int]$vbcomps.Count
            Add-Result '04' 'VBProject / VBComponents access' 'PASS' ("VBComponents.Count={0}" -f $n)
            Set-Cap '04' $true
        } catch {
            $msg = Format-Err $_
            if (Test-TrustAccessError $msg) {
                Add-Result '04' 'VBProject / VBComponents access' 'BLOCKED' ('Trust access to the VBA project object model is DISABLED. ' + $msg)
            } else {
                Add-Result '04' 'VBProject / VBComponents access' 'FAIL' ($msg + ' | ' + (Get-ComDiag))
            }
            Set-Cap '04' $false
        }
    }

    # =======================================================================
    # TEST 05 - Standard module
    # =======================================================================
    if (Test-Prereq '05' ('Standard module injection (' + $STD_MODULE_NAME + ')') @('04')) {
        $comp = $null; $cm = $null
        try {
            $comp = $vbcomps.Add($vbext_ct_StdModule)
            $comp.Name = $STD_MODULE_NAME
            $cm = $comp.CodeModule
            $cm.AddFromString($vbaStdModule)
            $lines = [int]$cm.CountOfLines
            [void](Remove-ComObjectRef $cm);   $cm = $null
            [void](Remove-ComObjectRef $comp); $comp = $null
            Add-Result '05' ('Standard module injection (' + $STD_MODULE_NAME + ')') 'PASS' ("CodeModule.CountOfLines={0}" -f $lines)
            Set-Cap '05' $true
        } catch {
            if ($null -ne $cm)   { [void](Remove-ComObjectRef $cm) }
            if ($null -ne $comp) { [void](Remove-ComObjectRef $comp) }
            Add-Result '05' ('Standard module injection (' + $STD_MODULE_NAME + ')') 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '05' $false
        }
    }

    # =======================================================================
    # TEST 06 - Class module
    # =======================================================================
    if (Test-Prereq '06' ('Class module injection (' + $CLS_MODULE_NAME + ')') @('04')) {
        $comp = $null; $cm = $null
        try {
            $comp = $vbcomps.Add($vbext_ct_ClassModule)
            $comp.Name = $CLS_MODULE_NAME
            $cm = $comp.CodeModule
            $cm.AddFromString($vbaClassModule)
            $lines = [int]$cm.CountOfLines
            [void](Remove-ComObjectRef $cm);   $cm = $null
            [void](Remove-ComObjectRef $comp); $comp = $null
            Add-Result '06' ('Class module injection (' + $CLS_MODULE_NAME + ')') 'PASS' ("CodeModule.CountOfLines={0}" -f $lines)
            Set-Cap '06' $true
        } catch {
            if ($null -ne $cm)   { [void](Remove-ComObjectRef $cm) }
            if ($null -ne $comp) { [void](Remove-ComObjectRef $comp) }
            Add-Result '06' ('Class module injection (' + $CLS_MODULE_NAME + ')') 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '06' $false
        }
    }

    # =======================================================================
    # TEST 07 - ThisWorkbook document module
    # =======================================================================
    if (Test-Prereq '07' 'ThisWorkbook document-module injection' @('04')) {
        $twComp = $null; $twCm = $null
        try {
            # The workbook has exactly one worksheet, so the sheet CodeName set is a single
            # value. ThisWorkbook is the Type=Document component that is not that sheet.
            # Localisation-proof: does not depend on the English name "ThisWorkbook".
            $sheetCodeName = [string]$ws.CodeName
            $n = [int]$vbcomps.Count
            for ($i = 1; $i -le $n; $i++) {
                $c = $vbcomps.Item($i)
                $isDoc = ([int]$c.Type -eq $vbext_ct_Document)
                $nm    = [string]$c.Name
                if ($isDoc -and ($nm -ne $sheetCodeName)) { $twComp = $c; break }
                [void](Remove-ComObjectRef $c)
                $c = $null
            }
            if ($null -eq $twComp) { throw 'Could not locate the ThisWorkbook document module among the VBComponents.' }

            $twName = [string]$twComp.Name
            $twCm = $twComp.CodeModule
            $twCm.AddFromString($vbaThisWorkbook)
            $cnt = [int]$twCm.CountOfLines
            $src = [string]$twCm.Lines(1, $cnt)
            [void](Remove-ComObjectRef $twCm);   $twCm = $null
            [void](Remove-ComObjectRef $twComp); $twComp = $null

            if ($src -notmatch [regex]::Escape($MARKER_THISWORKBOOK)) { throw 'Marker text not found in ThisWorkbook code module after injection.' }
            Add-Result '07' 'ThisWorkbook document-module injection' 'PASS' ("component='{0}'; lines={1}; marker present" -f $twName, $cnt)
            Set-Cap '07' $true
        } catch {
            if ($null -ne $twCm)   { [void](Remove-ComObjectRef $twCm) }
            if ($null -ne $twComp) { [void](Remove-ComObjectRef $twComp) }
            Add-Result '07' 'ThisWorkbook document-module injection' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '07' $false
        }
    }

    # =======================================================================
    # TEST 08 - Worksheet CodeName assignment AND worksheet document-module code
    #           08.1 CodeName -> shSmokeTest, proven by reading Worksheet.CodeName
    #           08.2 worksheet document-module injection (PCCM needs this for the
    #                output-sheet activation logic)
    #           Both sub-checks must pass for TEST 08 to pass.
    # =======================================================================
    if (Test-Prereq '08' 'Worksheet CodeName + worksheet document-module injection' @('04')) {
        $comp = $null; $props = $null; $prop = $null; $wsCm = $null
        $sub = New-Object System.Collections.ArrayList
        $ok1 = $false; $ok2 = $false
        try {
            $oldCodeName = [string]$ws.CodeName
            $comp = $vbcomps.Item($oldCodeName)

            # --- 08.1 mechanism 1: the _CodeName design-time property -------------
            try {
                $props = $comp.Properties
                try { $prop = $props.Item('_CodeName') } catch { $prop = $null }
                if ($null -eq $prop) {
                    $pn = [int]$props.Count
                    for ($i = 1; $i -le $pn; $i++) {
                        $p = $props.Item($i)
                        if ([string]$p.Name -eq '_CodeName') { $prop = $p; break }
                        [void](Remove-ComObjectRef $p)
                        $p = $null
                    }
                }
                if ($null -ne $prop) {
                    $prop.Value = $TARGET_CODENAME
                    [void](Remove-ComObjectRef $prop); $prop = $null
                }
                [void](Remove-ComObjectRef $props); $props = $null
            } catch {
                $null = $sub.Add(('08.1 mechanism "_CodeName property" raised: ' + (Format-Err $_)))
                if ($null -ne $prop)  { [void](Remove-ComObjectRef $prop);  $prop = $null }
                if ($null -ne $props) { [void](Remove-ComObjectRef $props); $props = $null }
            }

            if ([string]$ws.CodeName -eq $TARGET_CODENAME) {
                $script:codeNameMechanism = '_CodeName property'
                $ok1 = $true
            } else {
                # --- 08.1 mechanism 2: rename the document component --------------
                try { $comp.Name = $TARGET_CODENAME }
                catch { $null = $sub.Add(('08.1 mechanism "VBComponent.Name" raised: ' + (Format-Err $_))) }
                if ([string]$ws.CodeName -eq $TARGET_CODENAME) {
                    $script:codeNameMechanism = 'VBComponent.Name'
                    $ok1 = $true
                }
            }

            $null = $sub.Add(("08.1 [{0}] CodeName: '{1}' -> Worksheet.CodeName='{2}' (expected '{3}'); mechanism: {4}" -f $(if ($ok1) {'PASS'} else {'FAIL'}), $oldCodeName, [string]$ws.CodeName, $TARGET_CODENAME, $script:codeNameMechanism))

            # --- 08.2 worksheet document-module injection -------------------------
            # $comp remains the same component after the rename, so it is reused rather
            # than re-acquired (re-acquiring would risk a duplicate RCW).
            try {
                $wsCm = $comp.CodeModule
                $wsCm.AddFromString($vbaWorksheetModule)
                $cnt = [int]$wsCm.CountOfLines
                $src = [string]$wsCm.Lines(1, $cnt)
                [void](Remove-ComObjectRef $wsCm); $wsCm = $null
                $ok2 = ($src -match [regex]::Escape($MARKER_WORKSHEET))
                $null = $sub.Add(("08.2 [{0}] worksheet document-module code: lines={1}; marker '{2}' {3}" -f $(if ($ok2) {'PASS'} else {'FAIL'}), $cnt, $MARKER_WORKSHEET, $(if ($ok2) {'present'} else {'NOT FOUND'})))
            } catch {
                if ($null -ne $wsCm) { [void](Remove-ComObjectRef $wsCm); $wsCm = $null }
                $null = $sub.Add(('08.2 [FAIL] worksheet document-module code: ' + (Format-Err $_)))
            }

            [void](Remove-ComObjectRef $comp); $comp = $null

            $detail = "`r`n" + ($sub -join "`r`n")
            if ($ok1 -and $ok2) {
                Add-Result '08' 'Worksheet CodeName + worksheet document-module injection' 'PASS' $detail
                Set-Cap '08' $true
            } else {
                Add-Result '08' 'Worksheet CodeName + worksheet document-module injection' 'FAIL' ($detail + "`r`n      " + (Get-ComDiag))
                Set-Cap '08' $false
            }
        } catch {
            if ($null -ne $wsCm)  { [void](Remove-ComObjectRef $wsCm) }
            if ($null -ne $prop)  { [void](Remove-ComObjectRef $prop) }
            if ($null -ne $props) { [void](Remove-ComObjectRef $props) }
            if ($null -ne $comp)  { [void](Remove-ComObjectRef $comp) }
            Add-Result '08' 'Worksheet CodeName + worksheet document-module injection' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag) + "`r`n" + ($sub -join "`r`n"))
            Set-Cap '08' $false
        }
    }

    # =======================================================================
    # TEST 09 - Shape / button with OnAction  (independent of the VBA project)
    # =======================================================================
    if (Test-Prereq '09' 'Shape creation + OnAction assignment' @('02')) {
        try {
            $shapes = $ws.Shapes
            Register-ComRef $shapes 'Worksheet.Shapes'

            $shp = $shapes.AddShape($msoShapeRoundedRectangle, 20, 70, 150, 34)
            Register-ComRef $shp 'Shape(btnSmokeTest)'
            $shp.Name = $SHAPE_NAME

            # Transient chained objects: acquired explicitly, released immediately.
            try {
                $tf2 = $shp.TextFrame2
                $tr  = $tf2.TextRange
                $tr.Text = 'Run Smoke Macro'
                [void](Remove-ComObjectRef $tr);  $tr  = $null
                [void](Remove-ComObjectRef $tf2); $tf2 = $null
            } catch { }

            $shp.OnAction = $MACRO_NAME
            $readBack = [string]$shp.OnAction
            if ($readBack -notlike ('*' + $MACRO_NAME + '*')) { throw ("OnAction read back as '{0}'." -f $readBack) }
            Add-Result '09' 'Shape creation + OnAction assignment' 'PASS' ("name='{0}'; OnAction='{1}'" -f $SHAPE_NAME, $readBack)
            Set-Cap '09' $true
        } catch {
            Add-Result '09' 'Shape creation + OnAction assignment' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '09' $false
        }
    }

    # =======================================================================
    # TEST 10 - Chart  (independent of the VBA project)
    # =======================================================================
    if (Test-Prereq '10' 'ChartObject creation' @('02','02N')) {
        try {
            $chartObjects = $ws.ChartObjects()
            Register-ComRef $chartObjects 'Worksheet.ChartObjects'

            $co = $chartObjects.Add(280, 20, 320, 200)
            Register-ComRef $co 'ChartObject'
            $co.Name = 'chtSmokeTest'

            $cht = $co.Chart
            Register-ComRef $cht 'Chart'

            $srcRange = $ws.Range('D1:E6')
            $cht.SetSourceData($srcRange)
            [void](Remove-ComObjectRef $srcRange); $srcRange = $null

            $cht.ChartType = $xlColumnClustered
            $cnt = [int]$chartObjects.Count
            Add-Result '10' 'ChartObject creation' 'PASS' ("ChartObjects.Count={0}; name='chtSmokeTest'; ChartType={1}" -f $cnt, [int]$cht.ChartType)
            Set-Cap '10' $true
        } catch {
            Add-Result '10' 'ChartObject creation' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '10' $false
        }
    }

    # =======================================================================
    # TEST 11 - Macro execution  (requires the standard module to exist)
    # =======================================================================
    if (Test-Prereq '11' ('Macro execution (' + $MACRO_NAME + ')') @('05')) {
        $runErr1 = ''; $runErr2 = ''; $ran = $false
        try {
            $null = $excel.Run($MACRO_NAME)
            $ran = $true
        } catch {
            $runErr1 = Format-Err $_
            try {
                $qualified = "'" + [string]$wb.Name + "'!" + $MACRO_NAME
                $null = $excel.Run($qualified)
                $ran = $true
            } catch { $runErr2 = Format-Err $_ }
        }

        if ($ran) {
            try {
                $val = Get-CellText $ws 'B2'
                if ($val -eq $MARKER_MACRO) {
                    Add-Result '11' ('Macro execution (' + $MACRO_NAME + ')') 'PASS' ("SmokeTest!B2='{0}'" -f $val)
                    Set-Cap '11' $true
                } else {
                    Add-Result '11' ('Macro execution (' + $MACRO_NAME + ')') 'FAIL' (("Macro ran without error but SmokeTest!B2='{0}', expected '{1}'." -f $val, $MARKER_MACRO) + ' | ' + (Get-ComDiag))
                    Set-Cap '11' $false
                }
            } catch {
                Add-Result '11' ('Macro execution (' + $MACRO_NAME + ')') 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
                Set-Cap '11' $false
            }
        } else {
            $combined = ($runErr1 + ' || qualified-name retry: ' + $runErr2)
            if ((Test-MacroPolicyError $runErr1) -or (Test-MacroPolicyError $runErr2)) {
                Add-Result '11' ('Macro execution (' + $MACRO_NAME + ')') 'BLOCKED' ('Macro execution appears blocked by an Excel macro security policy. This is distinct from VBProject trust access, which passed at TEST 04. ' + $combined)
            } else {
                Add-Result '11' ('Macro execution (' + $MACRO_NAME + ')') 'FAIL' ($combined + ' | ' + (Get-ComDiag))
            }
            Set-Cap '11' $false
        }
    }

    # =======================================================================
    # TEST 12 - Save, release all first-instance COM children, close, quit
    #           Clean exit must be achieved by proper release, not by force.
    # =======================================================================
    if (Test-Prereq '12' 'Save, release COM, close workbook, quit Excel' @('02')) {
        try {
            # Save only if TEST 03 established a file on disk. Calling Save() on a workbook
            # that was never saved would try to resolve a default path rather than ours.
            if (Get-Cap '03') { $wb.Save() }
            $wb.Close($false)

            # Drop every first-instance local alias BEFORE the release pass, so that no
            # accidental RCW use can occur after release. The COM stack still holds its
            # own references and remains the owner used to release them; it is NOT
            # cleared here, because the release plan is built from it.
            $cht = $null; $co = $null; $chartObjects = $null
            $shp = $null; $shapes = $null
            $vbcomps = $null; $vbproj = $null
            $ws = $null; $worksheets = $null; $wb = $null; $workbooks = $null

            # Release EVERY first-instance child reference, leaf before parent,
            # BEFORE Application.Quit().
            $script:rel1 = Invoke-ComStackRelease -Stack $comStack1 -Label 'instance 1'

            $excel.Quit()
            [void](Remove-ComObjectRef $excel)
            $excel = $null

            # GC is a final aid only, never the ownership strategy.
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            $inst1Finished = $true
            $exited = Wait-ExcelExit -Identity $id1 -TimeoutSeconds 20
            $relTxt = Format-ReleaseResult $script:rel1 'instance 1'
            $relClean = ($script:rel1.Completed -and @($script:rel1.FailedReleaseLabels).Count -eq 0 -and $script:rel1.ReleasedCount -eq $script:rel1.UniqueCount)

            if (-not $relClean) {
                Add-Result '12' 'Save, release COM, close workbook, quit Excel' 'FAIL' ("The COM release pass did not complete cleanly. That is a lifecycle defect regardless of whether the process exited (natural exit observed: " + $exited + ").`r`n" + $relTxt)
                Set-Cap '12' $false
            } elseif ($exited) {
                Add-Result '12' 'Save, release COM, close workbook, quit Excel' 'PASS' ((("released {0} of {1} unique COM references ({2} registered, {3} duplicates) before Quit; pid {4} exited naturally, no force-stop required." -f $script:rel1.ReleasedCount, $script:rel1.UniqueCount, $script:rel1.RegisteredCount, $script:rel1.DuplicateCount, $id1.ProcessId)))
                Set-Cap '12' $true
            } elseif ($null -eq $id1 -or $id1.ProcessId -le 0) {
                Add-Result '12' 'Save, release COM, close workbook, quit Excel' 'FAIL' ("Quit() returned but no process identity was captured, so natural exit could not be verified.`r`n" + $relTxt)
                Set-Cap '12' $false
            } else {
                Add-Result '12' 'Save, release COM, close workbook, quit Excel' 'FAIL' ((("Every planned COM reference was released, but Excel pid {0} is still running after 20s, so something still holds a reference. The emergency cleanup path will handle the process; this test remains FAIL.`r`n" -f $id1.ProcessId) + $relTxt))
                Set-Cap '12' $false
            }
        } catch {
            Add-Result '12' 'Save, release COM, close workbook, quit Excel' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '12' $false
        }
    }

    # =======================================================================
    # TEST 13 - Reopen in a fresh Excel instance
    # =======================================================================
    if (Test-Prereq '13' 'Reopen saved .xlsm in a fresh instance' @('03','12')) {
        try {
            $preExisting2 = @()
            try { $preExisting2 = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id) } catch { }

            $comStackActive = $comStack2

            $excel2 = New-Object -ComObject Excel.Application
            $excel2.Visible       = $false
            $excel2.DisplayAlerts = $false
            $id2 = Get-ExcelIdentity -ExcelApp $excel2 -PreExistingPids $preExisting2

            $workbooks2 = $excel2.Workbooks
            Register-ComRef $workbooks2 'Application2.Workbooks'

            $wb2 = $workbooks2.Open($wbPath)
            Register-ComRef $wb2 'Workbook2'

            Add-Result '13' 'Reopen saved .xlsm in a fresh instance' 'PASS' ("pid {0} via {1}; opened '{2}'" -f $id2.ProcessId, $id2.Source, [string]$wb2.Name)
            Set-Cap '13' $true
        } catch {
            Add-Result '13' 'Reopen saved .xlsm in a fresh instance' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '13' $false
        }
    }

    # =======================================================================
    # TEST 14 - Persistence verification (component presence AND source markers)
    # =======================================================================
    if (Test-Prereq '14' 'Persistence verification after reopen' @('13')) {
        $sub = New-Object System.Collections.ArrayList
        $allOk = $true

        function Add-Check {
            param([string]$Label, [bool]$Ok, [string]$Info)
            $null = $sub.Add(("      [{0}] {1}{2}" -f $(if ($Ok) {'PASS'} else {'FAIL'}), $Label, $(if ($Info) { ': ' + $Info } else { '' })))
            return $Ok
        }

        try {
            # --- file format ---------------------------------------------------
            $ff = -1
            try { $ff = [int]$wb2.FileFormat } catch { }
            if (-not (Add-Check 'File format is macro-enabled' ($ff -eq $xlOpenXMLWorkbookMacroEnabled) ("FileFormat={0}, expected {1}" -f $ff, $xlOpenXMLWorkbookMacroEnabled))) { $allOk = $false }

            # --- worksheet + CodeName -------------------------------------------
            $worksheets2 = $wb2.Worksheets
            Register-ComRef $worksheets2 'Workbook2.Worksheets'
            $ws2 = $worksheets2.Item('SmokeTest')
            Register-ComRef $ws2 'Worksheet2(SmokeTest)'

            $cn = [string]$ws2.CodeName
            if (-not (Add-Check 'Worksheet CodeName persisted' ($cn -eq $TARGET_CODENAME) ("Worksheet.CodeName='{0}', expected '{1}'" -f $cn, $TARGET_CODENAME))) { $allOk = $false }

            # --- macro-written marker -------------------------------------------
            $mkVal = ''
            try { $mkVal = Get-CellText $ws2 'B2' } catch { }
            if (-not (Add-Check 'Macro-written marker persisted in B2' ($mkVal -eq $MARKER_MACRO) ("B2='{0}', expected '{1}'" -f $mkVal, $MARKER_MACRO))) { $allOk = $false }

            # --- shape + OnAction -----------------------------------------------
            $shpOk = $false; $shpInfo = 'shape not found'
            $shapes2 = $null; $shp2 = $null
            try {
                $shapes2 = $ws2.Shapes
                $shp2 = $shapes2.Item($SHAPE_NAME)
                $oa = [string]$shp2.OnAction
                $shpOk = ($oa -like ('*' + $MACRO_NAME + '*'))
                $shpInfo = ("OnAction='{0}'" -f $oa)
            } catch { $shpInfo = (Format-Err $_) }
            if ($null -ne $shp2)    { [void](Remove-ComObjectRef $shp2);    $shp2 = $null }
            if ($null -ne $shapes2) { [void](Remove-ComObjectRef $shapes2); $shapes2 = $null }
            if (-not (Add-Check 'Shape present with OnAction intact' $shpOk $shpInfo)) { $allOk = $false }

            # --- chart ------------------------------------------------------------
            $chOk = $false; $chInfo = ''
            $chartObjects2 = $null
            try {
                $chartObjects2 = $ws2.ChartObjects()
                $n = [int]$chartObjects2.Count
                $chOk = ($n -ge 1)
                $chInfo = ("ChartObjects.Count={0}" -f $n)
            } catch { $chInfo = (Format-Err $_) }
            if ($null -ne $chartObjects2) { [void](Remove-ComObjectRef $chartObjects2); $chartObjects2 = $null }
            if (-not (Add-Check 'ChartObject persisted' $chOk $chInfo)) { $allOk = $false }

            # --- VBA project: component presence AND source-marker persistence ----
            # Each component's own source is read and matched against its own unique
            # marker. Note MODULE_MARKER_OK is a substring of WORKSHEET_MODULE_MARKER_OK,
            # so markers are only ever matched against the component they belong to.
            $vbOk = $false
            try {
                $vbproj2 = $wb2.VBProject
                if ($null -ne $vbproj2) {
                    Register-ComRef $vbproj2 'Workbook2.VBProject'
                    $vbcomps2 = $vbproj2.VBComponents
                    if ($null -ne $vbcomps2) { Register-ComRef $vbcomps2 'VBProject2.VBComponents'; $vbOk = $true }
                }
            } catch { }
            if (-not (Add-Check 'VBProject accessible after reopen' $vbOk '')) { $allOk = $false }

            $names        = @()
            $modOk = $false; $clsOk = $false; $twOk = $false; $wsModOk = $false
            $modInfo = 'component not found'; $clsInfo = 'component not found'
            $twInfo  = 'component not found'; $wsModInfo = 'component not found'

            if ($vbOk) {
                $n = [int]$vbcomps2.Count
                for ($i = 1; $i -le $n; $i++) {
                    $c = $vbcomps2.Item($i)
                    $nm = [string]$c.Name
                    $tp = [int]$c.Type
                    $names += $nm
                    $src = Get-ComponentSource $c

                    if ($nm -eq $STD_MODULE_NAME) {
                        $modOk = ($src -match [regex]::Escape($MARKER_MODULE))
                        $modInfo = ("source {0} chars; marker '{1}' {2}" -f $src.Length, $MARKER_MODULE, $(if ($modOk) {'present'} else {'NOT FOUND'}))
                    } elseif ($nm -eq $CLS_MODULE_NAME) {
                        $clsOk = ($src -match [regex]::Escape($MARKER_CLASS))
                        $clsInfo = ("source {0} chars; marker '{1}' {2}" -f $src.Length, $MARKER_CLASS, $(if ($clsOk) {'present'} else {'NOT FOUND'}))
                    } elseif ($tp -eq $vbext_ct_Document -and $nm -eq $cn) {
                        $wsModOk = ($src -match [regex]::Escape($MARKER_WORKSHEET))
                        $wsModInfo = ("component='{0}'; source {1} chars; marker '{2}' {3}" -f $nm, $src.Length, $MARKER_WORKSHEET, $(if ($wsModOk) {'present'} else {'NOT FOUND'}))
                    } elseif ($tp -eq $vbext_ct_Document) {
                        $twOk = ($src -match [regex]::Escape($MARKER_THISWORKBOOK))
                        $twInfo = ("component='{0}'; source {1} chars; marker '{2}' {3}" -f $nm, $src.Length, $MARKER_THISWORKBOOK, $(if ($twOk) {'present'} else {'NOT FOUND'}))
                    }

                    [void](Remove-ComObjectRef $c)
                    $c = $null
                }
            }

            $null = $sub.Add(("      components after reopen: {0}" -f ($names -join ', ')))
            if (-not (Add-Check ($STD_MODULE_NAME + ' source marker persisted') $modOk   $modInfo))   { $allOk = $false }
            if (-not (Add-Check ($CLS_MODULE_NAME + ' source marker persisted') $clsOk   $clsInfo))   { $allOk = $false }
            if (-not (Add-Check 'ThisWorkbook source marker persisted'          $twOk    $twInfo))    { $allOk = $false }
            if (-not (Add-Check 'Worksheet document-module source marker persisted' $wsModOk $wsModInfo)) { $allOk = $false }

            $detail = "`r`n" + ($sub -join "`r`n")
            if ($allOk) {
                Add-Result '14' 'Persistence verification after reopen' 'PASS' $detail
                Set-Cap '14' $true
            } else {
                Add-Result '14' 'Persistence verification after reopen' 'FAIL' ($detail + "`r`n      " + (Get-ComDiag))
                Set-Cap '14' $false
            }
        } catch {
            Add-Result '14' 'Persistence verification after reopen' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag) + "`r`n" + ($sub -join "`r`n"))
            Set-Cap '14' $false
        }
    }

    # =======================================================================
    # TEST 15 - Release second-instance COM children, close, quit
    # =======================================================================
    if (Test-Prereq '15' 'Release COM, final close and quit' @('13')) {
        try {
            $wb2.Close($false)

            # Drop every second-instance local alias BEFORE the release pass. The COM
            # stack still owns the references and is not cleared here.
            $vbcomps2 = $null; $vbproj2 = $null; $ws2 = $null; $worksheets2 = $null
            $wb2 = $null; $workbooks2 = $null

            # Release EVERY second-instance child reference before Quit().
            $script:rel2 = Invoke-ComStackRelease -Stack $comStack2 -Label 'instance 2'

            $excel2.Quit()
            [void](Remove-ComObjectRef $excel2)
            $excel2 = $null

            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()

            $inst2Finished = $true
            $exited = Wait-ExcelExit -Identity $id2 -TimeoutSeconds 20
            $relTxt = Format-ReleaseResult $script:rel2 'instance 2'
            $relClean = ($script:rel2.Completed -and @($script:rel2.FailedReleaseLabels).Count -eq 0 -and $script:rel2.ReleasedCount -eq $script:rel2.UniqueCount)

            if (-not $relClean) {
                Add-Result '15' 'Release COM, final close and quit' 'FAIL' ("The COM release pass did not complete cleanly. That is a lifecycle defect regardless of whether the process exited (natural exit observed: " + $exited + ").`r`n" + $relTxt)
                Set-Cap '15' $false
            } elseif ($exited) {
                Add-Result '15' 'Release COM, final close and quit' 'PASS' ((("released {0} of {1} unique COM references ({2} registered, {3} duplicates) before Quit; pid {4} exited naturally, no force-stop required." -f $script:rel2.ReleasedCount, $script:rel2.UniqueCount, $script:rel2.RegisteredCount, $script:rel2.DuplicateCount, $id2.ProcessId)))
                Set-Cap '15' $true
            } elseif ($null -eq $id2 -or $id2.ProcessId -le 0) {
                Add-Result '15' 'Release COM, final close and quit' 'FAIL' ("Quit() returned but no process identity was captured, so natural exit could not be verified.`r`n" + $relTxt)
                Set-Cap '15' $false
            } else {
                Add-Result '15' 'Release COM, final close and quit' 'FAIL' ((("Every planned COM reference was released, but Excel pid {0} is still running after 20s. The emergency cleanup path will handle the process; this test remains FAIL.`r`n" -f $id2.ProcessId) + $relTxt))
                Set-Cap '15' $false
            }
        } catch {
            Add-Result '15' 'Release COM, final close and quit' 'FAIL' ((Format-Err $_) + ' | ' + (Get-ComDiag))
            Set-Cap '15' $false
        }
    }

}
catch {
    Add-Result 'XX' 'Unhandled script error' 'FAIL' ((Format-Err $_) + ' || stack: ' + $_.ScriptStackTrace)
}
finally {
    # =======================================================================
    # Cleanup. Runs whatever happened above, and only ever touches the Excel
    # instances this script created.
    # =======================================================================
    Write-Host ''
    Write-Host '  Cleaning up...' -ForegroundColor DarkGray

    if (-not $inst1Finished) {
        try { if ($null -ne $wb) { $wb.Close($false) } } catch { }
        $cht = $null; $co = $null; $chartObjects = $null
        $shp = $null; $shapes = $null
        $vbcomps = $null; $vbproj = $null
        $ws = $null; $worksheets = $null; $wb = $null; $workbooks = $null
        $script:rel1 = Invoke-ComStackRelease -Stack $comStack1 -Label 'instance 1'
        if ($null -ne $excel) {
            try { $excel.Quit() } catch { }
            [void](Remove-ComObjectRef $excel)
            $excel = $null
        }
    }
    if (-not $inst2Finished) {
        try { if ($null -ne $wb2) { $wb2.Close($false) } } catch { }
        $vbcomps2 = $null; $vbproj2 = $null; $ws2 = $null; $worksheets2 = $null
        $wb2 = $null; $workbooks2 = $null
        $script:rel2 = Invoke-ComStackRelease -Stack $comStack2 -Label 'instance 2'
        if ($null -ne $excel2) {
            try { $excel2.Quit() } catch { }
            [void](Remove-ComObjectRef $excel2)
            $excel2 = $null
        }
    }

    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if ($null -ne $id1) { Add-CleanupNote (Invoke-EmergencyExcelCleanup -Identity $id1 -Label 'instance 1') }
    if ($null -ne $id2) { Add-CleanupNote (Invoke-EmergencyExcelCleanup -Identity $id2 -Label 'instance 2') }

    # -----------------------------------------------------------------------
    # Verdict: the FIRST root cause in test order, never a downstream symptom.
    # SKIPPED never becomes the verdict.
    # -----------------------------------------------------------------------
    $ordered = @($results | Sort-Object Id)
    $root = $null
    foreach ($r in $ordered) {
        if ($r.Status -eq 'FAIL' -or $r.Status -eq 'BLOCKED') { $root = $r; break }
    }

    if ($null -eq $root) {
        $verdict = 'READY FOR PCCM STAGE B'
    } elseif ($root.Id -eq '01' -and $root.Status -eq 'BLOCKED') {
        $verdict = "BLOCKED $DASH EXCEL COM UNAVAILABLE"
    } elseif ($root.Id -eq '04' -and $root.Status -eq 'BLOCKED') {
        $verdict = "BLOCKED $DASH VBA PROJECT TRUST ACCESS"
    } elseif ($root.Id -eq '11' -and $root.Status -eq 'BLOCKED') {
        $verdict = "BLOCKED $DASH MACRO EXECUTION POLICY"
    } else {
        $verdict = ("FAILED $DASH TEST {0} {1}" -f $root.Id, $root.Name)
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
    foreach ($k in $env_info.Keys) { $null = $sb.AppendLine(('  {0,-26}: {1}' -f $k, $env_info[$k])) }
    $null = $sb.AppendLine(('  {0,-26}: {1}' -f 'CodeName mechanism used', $codeNameMechanism))
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('TEST RESULTS')
    $null = $sb.AppendLine('------------')
    foreach ($r in $ordered) {
        $null = $sb.AppendLine(('  TEST {0}  {1,-8} {2}' -f $r.Id, $r.Status, $r.Name))
        if ($r.Detail) {
            foreach ($line in ($r.Detail -split "`r?`n")) {
                if ($line.Trim().Length -gt 0) { $null = $sb.AppendLine(('           {0}' -f $line.TrimEnd())) }
            }
        }
    }
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('COM LIFECYCLE')
    $null = $sb.AppendLine('-------------')
    $null = $sb.AppendLine((Format-ReleaseResult $rel1 'Instance 1'))
    $null = $sb.AppendLine((Format-ReleaseResult $rel2 'Instance 2'))
    if ($comTrackWarnings.Count -gt 0) {
        $null = $sb.AppendLine('  COM tracking warnings:')
        foreach ($w in $comTrackWarnings) { $null = $sb.AppendLine(('      ' + $w)) }
    }
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('CLEANUP')
    $null = $sb.AppendLine('-------')
    if ($cleanupNotes.Count -eq 0) {
        $null = $sb.AppendLine('  No Excel instance was created, so there was nothing to clean up.')
    } else {
        foreach ($n in $cleanupNotes) { $null = $sb.AppendLine(('  ' + $n)) }
    }
    $null = $sb.AppendLine('')
    $null = $sb.AppendLine('SUMMARY')
    $null = $sb.AppendLine('-------')
    $null = $sb.AppendLine(('  PASS    : {0}' -f @($results | Where-Object { $_.Status -eq 'PASS' }).Count))
    $null = $sb.AppendLine(('  FAIL    : {0}' -f @($results | Where-Object { $_.Status -eq 'FAIL' }).Count))
    $null = $sb.AppendLine(('  BLOCKED : {0}' -f @($results | Where-Object { $_.Status -eq 'BLOCKED' }).Count))
    $null = $sb.AppendLine(('  SKIPPED : {0}' -f @($results | Where-Object { $_.Status -eq 'SKIPPED' }).Count))
    if ($null -ne $root) {
        $null = $sb.AppendLine(('  Root cause: TEST {0} ({1}) {2}' -f $root.Id, $root.Status, $root.Name))
    }
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
    try { Set-Content -LiteralPath $reportPath -Value $reportText -Encoding UTF8 }
    catch { Write-Host ('  Could not write report file: ' + $_.Exception.Message) -ForegroundColor Red }

    Write-Host ''
    Write-Host $reportText
    Write-Host ''
    Write-Host ('  Report written to: {0}' -f $reportPath) -ForegroundColor Cyan
    Write-Host '  Please return that file.' -ForegroundColor Cyan
    Write-Host ''
}
