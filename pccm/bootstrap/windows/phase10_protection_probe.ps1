<#
.SYNOPSIS
    PCCM Phase 10 - the PROTECTION / TABLE-STRUCTURE PROBE.

.DESCRIPTION
    ONE QUESTION, ASKED OF REAL EXCEL, IN ABOUT A MINUTE.

    Windows Run 3 of the benchmark died on `ListRow.Delete()` with "Table
    features aren't available because the sheet is protected." That call came
    from PowerShell. Static reading of `pccm/src/vba` says the same class of
    call - `ListColumns.Add`, `ListRows.Add`, `ListRows(...).Delete` - sits
    inside SIX of the seven user commands, so the same wall should stop a user
    pressing a button. But "should" is not evidence, and the correction it
    implies edits accepted production modules and an accepted contract.

    So this probe asks Excel directly, and separates the two capabilities that
    the benchmark run conflated:

      1. CONTROL - can code write a VALUE to a locked cell on a protected sheet?
         Run 3 already suggests yes (the fixture set four Setup scalars before it
         died). This confirms it deliberately rather than as a side effect.

      2. THE QUESTION - can the real production commands perform their
         CONTRACTED STRUCTURAL operation on a protected sheet? PCCM_ApplyTimeline
         adds year ListColumns to three grids and PCCM_Calculate resizes the
         _Calc tables; neither is optional and neither is capacity-dependent.

    IT MEASURES NOTHING AND DECIDES NOTHING. It runs each endpoint once, reports
    what the workbook announced, reports whether protection was still in force
    before and after, and stops. No timing, no baseline, no Gate-B result.

    IT CHANGES NO PRODUCTION CODE. It opens a disposable Stage-B build, presses
    real buttons through their real entry points, and never saves.

.PARAMETER BuildDir
    The Stage-A build directory. Defaults to <repo>/pccm/build.

.PARAMETER WorkDir
    Where the disposable copy is made. Defaults to the system temp directory.

.NOTES
    SAFETY. No security setting is altered, no registry key is touched, and no
    Excel process this script did not create is ever terminated. Shutdown is the
    accepted `com_lifecycle.ps1` path. The workbook is never saved.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [string]$WorkDir,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'com_lifecycle.ps1')

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $scriptDir))
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }
if ([string]::IsNullOrWhiteSpace($WorkDir))  { $WorkDir  = [System.IO.Path]::GetTempPath() }

$script:ProbeLines = New-Object System.Collections.ArrayList
$script:ProbePath = ''

function Write-ProbeLine {
    param([string]$Text = '')
    $null = $script:ProbeLines.Add($Text)
    Write-Host $Text
    if (-not [string]::IsNullOrWhiteSpace($script:ProbePath)) {
        try {
            Set-Content -LiteralPath $script:ProbePath -Value ($script:ProbeLines -join "`r`n") -Encoding UTF8
        } catch {
            $script:ProbePath = ''
            Write-Host '  (the probe log could not be written through)' -ForegroundColor DarkYellow
        }
    }
}

# NORMALISE, VALIDATE, THEN USE - the W1 lesson, applied here from the start.
function Get-ProbeProperty {
    param($InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    return $property.Value
}

function Get-ProbeNamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $range = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $range = $nm.RefersToRange
        return $range.Value2
    } finally {
        if ($null -ne $range) { Release-Transient $range 'Range'; $range = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';  $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names'; $names = $null }
    }
}

function Set-ProbeNamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $names = $null; $nm = $null; $range = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $range = $nm.RefersToRange
        $range.Value2 = $Value
    } finally {
        if ($null -ne $range) { Release-Transient $range 'Range'; $range = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';  $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names'; $names = $null }
    }
}

# IS THE SHEET ACTUALLY PROTECTED RIGHT NOW? Asked of the sheet, not of a
# constant, and asked at the moment it matters.
function Get-ProbeProtectionState {
    param($Workbook)
    $sheets = $null
    $protectedCount = 0; $total = 0; $unprotected = @()
    try {
        $sheets = $Workbook.Worksheets
        for ($index = 1; $index -le $sheets.Count; $index++) {
            $sheet = $null
            try {
                $sheet = $sheets.Item($index)
                $total++
                if ($sheet.ProtectContents) { $protectedCount++ } else { $unprotected += [string]$sheet.Name }
            } finally {
                if ($null -ne $sheet) { Release-Transient $sheet 'Worksheet'; $sheet = $null }
            }
        }
    } finally {
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
    }
    return [pscustomobject]@{
        Total = $total; Protected = $protectedCount; Unprotected = $unprotected
        Structure = [bool]$Workbook.ProtectStructure
    }
}

# ONE ENDPOINT, ONCE. The announcement is production's own; nothing is inferred
# from the absence of an exception, because a REFUSAL is also an announcement.
function Invoke-ProbeEndpoint {
    param($Excel, [string]$Endpoint, [string]$Argument = '', [switch]$WithArgument)
    $raised = ''
    $result = ''
    try {
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        if ($WithArgument) { $Excel.Run($Endpoint, $Argument) | Out-Null }
        else               { $Excel.Run($Endpoint) | Out-Null }
        $result = [string]$Excel.Run('PCCM_AutomationResult')
    } catch {
        $raised = [string]$_.Exception.Message
    }
    return [pscustomobject]@{
        Endpoint = $Endpoint
        Raised   = $raised
        Result   = $result
        Ok       = ([bool](($raised -eq '') -and ($result -like 'OK|*')))
    }
}

function Write-ProbeOutcome {
    param($Outcome, [string]$Expectation)
    Write-ProbeLine ('  ' + $Outcome.Endpoint)
    if ($Outcome.Raised -ne '') {
        Write-ProbeLine ('    RAISED   : ' + $Outcome.Raised)
    } else {
        Write-ProbeLine ('    announced: ' + $Outcome.Result)
    }
    Write-ProbeLine ('    verdict  : ' + $(if ($Outcome.Ok) { 'SUCCEEDED' } else { 'DID NOT SUCCEED' }))
    Write-ProbeLine ('    what this settles: ' + $Expectation)
    Write-ProbeLine ''
}

# ===========================================================================
# PREFLIGHT AND THE DISPOSABLE BUILD
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 10 protection / table-structure probe' -ForegroundColor Cyan
Write-Host '=================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host 'This asks Excel one question. It measures nothing and decides' -ForegroundColor Yellow
Write-Host 'nothing: no timing, no baseline, no Gate-B result.' -ForegroundColor Yellow
Write-Host ''

$manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath  = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
foreach ($required in @($manifestPath, $inspectPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ("$required not found. Run the Stage-A build first: " +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest   = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath $inspectPath  -Raw | ConvertFrom-Json

$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
$tempRoot = Join-Path $WorkDir ('pccm-phase10-protection-probe-' + $stamp)
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
Copy-Item -LiteralPath $manifestPath -Destination $tempRoot
Copy-Item -LiteralPath $inspectPath  -Destination $tempRoot
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:ProbePath = Join-Path $tempRoot ('phase10_protection_probe_' + $stamp + '.log')
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

& (Join-Path $scriptDir 'build_stage_b.ps1') -BuildDir $tempRoot -Force
if (($LASTEXITCODE -ne 0) -or (-not (Test-Path -LiteralPath $stageBPath))) {
    Write-Host ''
    Write-Host 'The Stage-B bootstrap did not complete. Nothing was probed.' -ForegroundColor Red
    exit 1
}

Write-ProbeLine 'PCCM - PHASE 10 PROTECTION / TABLE-STRUCTURE PROBE'
Write-ProbeLine '================================================='
Write-ProbeLine ''
Write-ProbeLine ('run started : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-ProbeLine ('workbook    : ' + $stageBPath)
Write-ProbeLine ''
Write-ProbeLine 'THE QUESTION'
Write-ProbeLine '------------'
Write-ProbeLine 'Windows Run 3 of the benchmark died on a ListRow.Delete() with'
Write-ProbeLine '"Table features aren''t available because the sheet is protected."'
Write-ProbeLine 'That call came from PowerShell. Six of the seven user commands make'
Write-ProbeLine 'the SAME CLASS of call from VBA. This asks whether the real commands'
Write-ProbeLine 'can do their contracted work on the protected workbook.'
Write-ProbeLine ''

# ===========================================================================
# THE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null; $excelIdentity = $null; $rel = $null
$verdict = 'INCONCLUSIVE'

try {
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)

    # --- is protection actually in force? ---------------------------------
    $before = Get-ProbeProtectionState -Workbook $wb
    Write-ProbeLine 'PROTECTION STATE AS THE WORKBOOK OPENED'
    Write-ProbeLine '---------------------------------------'
    Write-ProbeLine ('  sheets protected : ' + [string]$before.Protected + ' of ' + [string]$before.Total)
    Write-ProbeLine ('  structure        : ' + [string]$before.Structure)
    if (@($before.Unprotected).Count -gt 0) {
        Write-ProbeLine ('  NOT protected    : ' + ((@($before.Unprotected)) -join ', '))
    }
    # THE OWNER'S OWN ANSWER, asked by its real name. `ProtectionIsApplied` carries
    # no PCCM_ prefix because it is not a user command - it is the one fact the
    # protection owner reports about itself.
    Write-ProbeLine ('  applied by       : ' + [string]$excel.Run('ProtectionIsApplied'))
    Write-ProbeLine ''
    if ($before.Protected -lt $before.Total) {
        Write-ProbeLine 'THE PROBE IS INCONCLUSIVE: the workbook did not come back protected,'
        Write-ProbeLine 'so nothing below is a statement about protected behaviour. That is'
        Write-ProbeLine 'itself a Workbook_Open protection-lifecycle finding.'
        Write-ProbeLine ''
        $verdict = 'INCONCLUSIVE - the workbook was not protected'
    }

    # --- CONTROL: a value written to a locked cell -------------------------
    Write-ProbeLine 'CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?'
    Write-ProbeLine '--------------------------------------------------'
    $discountName = [string]$inspection.inputs.discount_rate.defined_name
    $original = Get-ProbeNamedValue -Workbook $wb -DefinedName $discountName
    $writeRaised = ''
    try {
        Set-ProbeNamedValue -Workbook $wb -DefinedName $discountName -Value ([double]0.05)
    } catch {
        $writeRaised = [string]$_.Exception.Message
    }
    if ($writeRaised -eq '') {
        Write-ProbeLine '  a cell VALUE write on the protected sheet: SUCCEEDED'
        Write-ProbeLine '  so UserInterfaceOnly is honoured for code that writes VALUES.'
        try { Set-ProbeNamedValue -Workbook $wb -DefinedName $discountName -Value $original }
        catch { Write-ProbeLine '  (the original value could not be put back; the copy is disposable)' }
    } else {
        Write-ProbeLine ('  a cell VALUE write on the protected sheet: REFUSED - ' + $writeRaised)
        Write-ProbeLine '  so protection is blocking code writes as well, which is a DIFFERENT'
        Write-ProbeLine '  and larger finding than the table-structure one.'
    }
    Write-ProbeLine ''

    # --- THE QUESTION: the real commands -----------------------------------
    Write-ProbeLine 'THE QUESTION - CAN THE REAL COMMANDS DO THEIR STRUCTURAL WORK?'
    Write-ProbeLine '--------------------------------------------------------------'
    Write-ProbeLine 'Each is the accepted public entry point, invoked once, exactly as its'
    Write-ProbeLine 'button invokes it. What is reported is what the workbook announced.'
    Write-ProbeLine ''

    # Apply Timeline is FIRST because it is unconditionally structural: a fresh
    # workbook has no year columns, so any timeline adds ListColumns to three
    # grids. It is also the first button a real user presses.
    Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.base_year.defined_name) -Value ([double]2026)
    Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.project_start_year.defined_name) -Value ([double]2027)
    Set-ProbeNamedValue -Workbook $wb -DefinedName ([string]$inspection.inputs.duration_years.defined_name) -Value ([double]3)

    $timeline = Invoke-ProbeEndpoint -Excel $excel -Endpoint 'PCCM_ApplyTimeline'
    Write-ProbeOutcome -Outcome $timeline -Expectation `
        ('ListColumns.Add on three grids. A fresh workbook has no year columns, so ' +
         'this fires for ANY timeline and is not capacity-dependent.')

    $addCost = Invoke-ProbeEndpoint -Excel $excel -Endpoint 'PCCM_AddCostLine'
    Write-ProbeOutcome -Outcome $addCost -Expectation `
        ('ListRows.Add only when the reserved capacity is exhausted, but SyncRows ' +
         'runs on every add and can add a profiling row.')

    $addRisk = Invoke-ProbeEndpoint -Excel $excel -Endpoint 'PCCM_AddRisk'
    Write-ProbeOutcome -Outcome $addRisk -Expectation 'the same path on the Risk Register.'

    $calculate = Invoke-ProbeEndpoint -Excel $excel -Endpoint 'PCCM_Calculate'
    Write-ProbeOutcome -Outcome $calculate -Expectation `
        ('the _Calc tables are resized to the model on every Calculate, so ListRows.Add ' +
         'or ListRows(...).Delete fires for any model with more than one driver.')

    # --- protection state afterwards ---------------------------------------
    $after = Get-ProbeProtectionState -Workbook $wb
    Write-ProbeLine 'PROTECTION STATE AFTER THE COMMANDS'
    Write-ProbeLine '-----------------------------------'
    Write-ProbeLine ('  sheets protected : ' + [string]$after.Protected + ' of ' + [string]$after.Total)
    Write-ProbeLine ('  structure        : ' + [string]$after.Structure)
    if (@($after.Unprotected).Count -gt 0) {
        Write-ProbeLine ('  NOT protected    : ' + ((@($after.Unprotected)) -join ', '))
    }
    Write-ProbeLine ''

    if ($verdict -eq 'INCONCLUSIVE') {
        $structural = @($timeline, $addCost, $addRisk, $calculate)
        $failed = @($structural | Where-Object { -not $_.Ok })
        if ($failed.Count -eq 0) {
            $verdict = ('PRODUCTION IS FINE UNDER PROTECTION - every structural command ' +
                        'succeeded, so Windows Run 3 was a HARNESS defect only')
        } else {
            $verdict = ('PRODUCTION IS BLOCKED BY PROTECTION - ' + [string]$failed.Count +
                        ' of ' + [string]$structural.Count +
                        ' structural commands did not succeed on the protected workbook')
        }
    }
    $excel.Run('PCCM_AutomationEnd') | Out-Null
} catch {
    Write-ProbeLine ''
    Write-ProbeLine ('THE PROBE RAISED: ' + (Format-Err $_))
    $verdict = 'INCONCLUSIVE - the probe session raised before it finished'
} finally {
    $rel = New-ReleaseLedger 'phase-10 protection probe'
    try {
        if ($null -ne $wb) {
            # NEVER SAVED. The disposable copy is discarded.
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-NamedRelease $rel $wb        'Workbook';  $wb        = $null
        Invoke-NamedRelease $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-NamedRelease $rel $excel 'Application'; $excel = $null
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        $rel.NaturalExit = Wait-ExcelExit -Identity $excelIdentity
        if (-not $rel.NaturalExit) {
            $rel.EmergencyRequired = $true
            Write-ProbeLine (Invoke-EmergencyExcelCleanup -Identity $excelIdentity `
                -Label 'phase-10 protection probe')
        }
    } catch {
        Write-ProbeLine ('Shutdown raised: ' + (Format-Err $_))
    }
    Write-ProbeLine 'SHUTDOWN'
    Write-ProbeLine (Format-ReleaseLedger $rel)
}

Write-ProbeLine ''
Write-ProbeLine 'VERDICT'
Write-ProbeLine '-------'
Write-ProbeLine ('  ' + $verdict)
Write-ProbeLine ''
Write-ProbeLine 'THIS IS NOT A BASELINE, NOT A GATE-B RESULT AND NOT AN ACCEPTANCE RUN.'
Write-ProbeLine 'It is one question, asked once, so a correction is made against evidence'
Write-ProbeLine 'rather than against a reading of the documentation.'
Write-ProbeLine ''
Write-ProbeLine ('  log : ' + $script:ProbePath)

$kept = $script:ProbePath
if (-not $KeepArtifacts) {
    try {
        $kept = Join-Path $WorkDir ('pccm-phase10-protection-probe-' + $stamp + '.log')
        Copy-Item -LiteralPath $script:ProbePath -Destination $kept -Force
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    } catch { $kept = $script:ProbePath }
}
Write-Host ''
Write-Host ('Probe log kept at ' + $kept) -ForegroundColor Yellow
