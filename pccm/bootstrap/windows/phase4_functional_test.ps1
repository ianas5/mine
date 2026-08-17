<#
.SYNOPSIS
    PCCM Phase-4 Windows functional test harness. Exercises the structural runtime
    against a real Excel instance and produces a human-readable report.

.DESCRIPTION
    DO NOT RUN THIS UNTIL PHASE-4 GATE-A SOURCE REVIEW HAS BEEN APPROVED.

    The harness never touches the real build output. It copies the Stage-A build
    into a temporary directory, runs the Stage-B bootstrap there, and drives that
    disposable .xlsm.

    It contains no hand-written expected timeline values. Every expected shape --
    project-year headers, required inflation years, column counts, empty-span
    handling -- is read from build/phase4_scenarios.json, which the Stage-A build
    emits from the same structural oracle the Linux tests exercise. Value
    preservation ("unchanged bit for bit") is checked by comparing the workbook
    against its OWN captured previous values, because that is what the claim
    actually means.

    Test matrix:
      A  Stage-B build: .xlsm, FileFormat 52, 14 CodeNames, modules, five buttons,
         natural COM shutdown
      B  Permanent Cost Line IDs, including non-reuse after deletion and stability
         under reordering
      C  Permanent Risk IDs, independently sequenced
      D  First timeline application
      E  Duration increase
      F  Start-year shift with unchanged duration
      G  Duration decrease: destructive confirmation, cancelled then accepted
      H  Base-Year movement, earlier and later
      I  Combined three-way change
      J  Degenerate inflation span
      K  Profiling synchronisation by permanent ID
      L  Runtime failure containment and logical restore

    Safety, unchanged from the readiness gate: no security setting is altered, no
    registry key is touched, no Trusted Location is added, and no Excel process
    this script did not create is ever terminated. A forced stop is never reported
    as a pass.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.PARAMETER KeepArtifacts
    Leave the temporary workbook on disk for inspection.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'com_lifecycle.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

$results  = New-Object System.Collections.ArrayList
$notes    = New-Object System.Collections.ArrayList

function Add-Result {
    param([string]$Id, [string]$Name, [string]$Status, [string]$Detail = '')
    $null = $results.Add([pscustomobject]@{ Id = $Id; Name = $Name; Status = $Status; Detail = $Detail })
    $colour = 'Green'
    if ($Status -eq 'FAIL') { $colour = 'Red' } elseif ($Status -eq 'SKIP') { $colour = 'Yellow' }
    Write-Host ("  [{0}] {1}  {2}" -f $Status, $Id, $Name) -ForegroundColor $colour
    if ($Detail) {
        foreach ($line in ($Detail -split "`r?`n")) { Write-Host ("        " + $line) -ForegroundColor DarkGray }
    }
}

function Add-Note { param([string]$Text) $null = $notes.Add($Text) }

function New-Checklist { return (New-Object System.Collections.ArrayList) }

function Add-Check {
    param($List, [string]$Text, [bool]$Condition, [string]$Detail = '')
    if ($Condition) {
        $null = $List.Add(("  ok   " + $Text))
    } else {
        $null = $List.Add(("  FAIL " + $Text + $(if ($Detail) { " -- " + $Detail } else { '' })))
    }
    return $Condition
}

function Test-ChecklistOk { param($List) return -not (@($List | Where-Object { $_ -like '  FAIL*' }).Count -gt 0) }
function Format-Checklist { param($List) return ($List -join "`r`n") }

Write-Host ''
Write-Host 'PCCM - Phase 4 Windows functional test' -ForegroundColor Cyan
Write-Host '======================================' -ForegroundColor Cyan
Write-Host ''

# ===========================================================================
# Prepare a disposable copy of the build
# ===========================================================================
$tempRoot   = $null
$manifest   = $null
$scenarios  = $null
$stageBPath = $null

try {
    $manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
    $scenarioPath = Join-Path $BuildDir 'phase4_scenarios.json'
    foreach ($required in @($manifestPath, $scenarioPath)) {
        if (-not (Test-Path -LiteralPath $required)) {
            throw "$required not found. Run the Stage-A build first: python3 pccm/builder/build_stage_a.py"
        }
    }
    $manifest  = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    $scenarios = Get-Content -LiteralPath $scenarioPath -Raw | ConvertFrom-Json

    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("pccm-phase4-" + (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $null = New-Item -ItemType Directory -Path $tempRoot -Force
    Copy-Item -LiteralPath (Join-Path $BuildDir $manifest.stage_a_filename) -Destination $tempRoot
    Copy-Item -LiteralPath $manifestPath -Destination $tempRoot
    Copy-Item -LiteralPath $scenarioPath -Destination $tempRoot
    Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

    $stageBPath = Join-Path $tempRoot $manifest.stage_b_filename
    Add-Note ("Working copy: " + $tempRoot)
} catch {
    Add-Result '00' 'Prepare a disposable copy of the build' 'FAIL' (Format-Err $_)
    Write-Host ''
    Write-Host 'PHASE-4 FUNCTIONAL TEST ABORTED before Excel was started.' -ForegroundColor Red
    exit 1
}

# ===========================================================================
# A. Stage-B build
# ===========================================================================
$buildOk = $false
try {
    $bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
    & $bootstrap -BuildDir $tempRoot -Force
    $exitCode = $LASTEXITCODE
    $list = New-Checklist
    $null = Add-Check $list 'the Stage-B bootstrap exited cleanly' ($exitCode -eq 0) ("exit code $exitCode")
    $null = Add-Check $list 'the .xlsm exists' (Test-Path -LiteralPath $stageBPath) $stageBPath
    $buildOk = Test-ChecklistOk $list
    Add-Result 'A' 'Stage-B build (bootstrap, CodeNames, modules, buttons, natural shutdown)' `
        $(if ($buildOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
} catch {
    Add-Result 'A' 'Stage-B build' 'FAIL' (Format-Err $_)
}

# ===========================================================================
# Drive the workbook
# ===========================================================================
$excel = $null; $workbooks = $null; $wb = $null; $worksheets = $null
$id = $null
$rel = $null

# --- helpers over the driven workbook --------------------------------------
function Get-NamedValue {
    param($Workbook, [string]$DefinedName)
    $nm = $null; $rng = $null
    try {
        $nm = $Workbook.Names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $v = $rng.Value2
        if ($null -eq $v) { return '' }
        return [string]$v
    } finally {
        if ($null -ne $rng) { Release-Transient $rng 'Range(name)'; $rng = $null }
        if ($null -ne $nm)  { Release-Transient $nm  'Name';        $nm  = $null }
    }
}

function Set-NamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $nm = $null; $rng = $null
    try {
        $nm = $Workbook.Names.Item($DefinedName)
        $rng = $nm.RefersToRange
        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }
    } finally {
        if ($null -ne $rng) { Release-Transient $rng 'Range(name)'; $rng = $null }
        if ($null -ne $nm)  { Release-Transient $nm  'Name';        $nm  = $null }
    }
}

function Get-TableColumnNames {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $ws = $null; $los = $null; $lo = $null; $cols = $null
    $out = @()
    try {
        $ws = $Workbook.Worksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        for ($i = 1; $i -le $cols.Count; $i++) {
            $c = $null
            try { $c = $cols.Item($i); $out += [string]$c.Name }
            finally { if ($null -ne $c) { Release-Transient $c 'ListColumn'; $c = $null } }
        }
    } finally {
        if ($null -ne $cols) { Release-Transient $cols 'ListColumns'; $cols = $null }
        if ($null -ne $lo)   { Release-Transient $lo   'ListObject';  $lo   = $null }
        if ($null -ne $los)  { Release-Transient $los  'ListObjects'; $los  = $null }
        if ($null -ne $ws)   { Release-Transient $ws   'Worksheet';   $ws   = $null }
    }
    return $out
}

# Reads the whole data body of a table as a jagged array of strings, so a later
# comparison is a plain value comparison and never a live COM read.
function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $ws = $null; $los = $null; $lo = $null; $body = $null
    $rows = @()
    try {
        $ws = $Workbook.Worksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        if ($null -eq $body) { return $rows }
        for ($r = 1; $r -le $body.Rows.Count; $r++) {
            $line = @()
            for ($c = 1; $c -le $body.Columns.Count; $c++) {
                $cell = $null
                try {
                    $cell = $body.Cells($r, $c)
                    $v = $cell.Value2
                    if ($null -eq $v) { $line += '' } else { $line += [string]$v }
                } finally {
                    if ($null -ne $cell) { Release-Transient $cell 'Range(cell)'; $cell = $null }
                }
            }
            $rows += ,$line
        }
    } finally {
        if ($null -ne $body) { Release-Transient $body 'Range(body)'; $body = $null }
        if ($null -ne $lo)   { Release-Transient $lo   'ListObject';  $lo   = $null }
        if ($null -ne $los)  { Release-Transient $los  'ListObjects'; $los  = $null }
        if ($null -ne $ws)   { Release-Transient $ws   'Worksheet';   $ws   = $null }
    }
    return $rows
}

function Set-TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value)
    $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $ws = $Workbook.Worksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        if ($Value -is [string]) { $cell.Value2 = [string]$Value } else { $cell.Value2 = [double]$Value }
    } finally {
        if ($null -ne $cell) { Release-Transient $cell 'Range(cell)'; $cell = $null }
        if ($null -ne $body) { Release-Transient $body 'Range(body)'; $body = $null }
        if ($null -ne $lo)   { Release-Transient $lo   'ListObject';  $lo   = $null }
        if ($null -ne $los)  { Release-Transient $los  'ListObjects'; $los  = $null }
        if ($null -ne $ws)   { Release-Transient $ws   'Worksheet';   $ws   = $null }
    }
}

function Get-GridInfo {
    param([string]$Key)
    foreach ($g in $manifest.grids) { if ($g.key -eq $Key) { return $g } }
    throw "no grid '$Key' in the manifest"
}

function Get-RegisterInfo {
    param([string]$Key)
    foreach ($r in $manifest.registers) { if ($r.key -eq $Key) { return $r } }
    throw "no register '$Key' in the manifest"
}

function Get-IdColumnValues {
    param($Workbook, $Info)
    $out = @()
    foreach ($row in (Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}

if ($buildOk) {
  $preExisting = Get-PreExistingExcelPids
  try {
    $excel = New-Object -ComObject Excel.Application
    $id = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)
    $worksheets = $wb.Worksheets

    $costGrid   = Get-GridInfo 'cost_profiling'
    $riskGrid   = Get-GridInfo 'risk_profiling'
    $inflGrid   = Get-GridInfo 'inflation'
    $costReg    = Get-RegisterInfo 'cost_lines'
    $riskReg    = Get-RegisterInfo 'risk_register'
    $fixedCost  = $costGrid.fixed_columns.Count
    $fixedInfl  = $inflGrid.fixed_columns.Count

    # Automation on: confirmations are answered by the harness, not by a human.
    # No failure stage is armed, so every operation runs its real path.
    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null

    # -------------------------------------------------------------------
    # B. Permanent Cost Line IDs
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $expected = $scenarios.identity.cost_sequence

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $afterThree = Get-IdColumnValues -Workbook $wb -Info $costReg
        $null = Add-Check $list 'three adds issue CL-001, CL-002, CL-003' `
            (($afterThree -join ',') -eq (($expected[0..2]) -join ',')) ("got " + ($afterThree -join ','))

        $excel.Run('PCCM_DeleteCostLineById', $expected[1]) | Out-Null
        $afterDelete = Get-IdColumnValues -Workbook $wb -Info $costReg
        $null = Add-Check $list 'deleting CL-002 leaves CL-001 and CL-003' `
            (($afterDelete -join ',') -eq (($expected[0], $expected[2]) -join ',')) ("got " + ($afterDelete -join ','))

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $afterFourth = Get-IdColumnValues -Workbook $wb -Info $costReg
        $null = Add-Check $list 'the next add issues CL-004, NOT the deleted CL-002' `
            (($afterFourth -join ',') -eq (($expected[0], $expected[2], $expected[3]) -join ',')) ("got " + ($afterFourth -join ','))
        $null = Add-Check $list 'CL-002 was not reused' ($afterFourth -notcontains $expected[1])

        # Reorder by rewriting a user field, then confirm the ID still travels with
        # its own row data rather than with a row position.
        Set-TableCell -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name -RowIndex 1 -ColumnIndex 3 -Value 'FIRST ROW MARKER'
        $body = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $null = Add-Check $list 'the marker sits on the row that carries CL-001' `
            (($body[0][0] -eq $expected[0]) -and ($body[0][2] -eq 'FIRST ROW MARKER'))

        Add-Result 'B' 'Permanent Cost Line IDs: sequence, non-reuse, stability' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'B' 'Permanent Cost Line IDs' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # C. Permanent Risk IDs, independently sequenced
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $expected = $scenarios.identity.risk_sequence

        $excel.Run('PCCM_AddRisk') | Out-Null
        $excel.Run('PCCM_AddRisk') | Out-Null
        $afterTwo = Get-IdColumnValues -Workbook $wb -Info $riskReg
        $null = Add-Check $list 'two adds issue R-001 and R-002' `
            (($afterTwo -join ',') -eq (($expected[0], $expected[1]) -join ',')) ("got " + ($afterTwo -join ','))
        $null = Add-Check $list 'the risk sequence is independent of the cost sequence' `
            ($afterTwo[0] -eq 'R-001') 'four cost lines have been added; the risk sequence still starts at 1'

        $excel.Run('PCCM_DeleteRiskById', $expected[1]) | Out-Null
        $excel.Run('PCCM_AddRisk') | Out-Null
        $afterReadd = Get-IdColumnValues -Workbook $wb -Info $riskReg
        $null = Add-Check $list 'R-002 is not reused after deletion' `
            (($afterReadd -join ',') -eq (($expected[0], $expected[2]) -join ',')) ("got " + ($afterReadd -join ','))

        Add-Result 'C' 'Permanent Risk IDs: independent sequence, non-reuse' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'C' 'Permanent Risk IDs' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # D - J. Timeline scenarios, driven from the oracle-derived fixture
    # -------------------------------------------------------------------
    $stepIndex = 0
    foreach ($step in $scenarios.steps) {
        $stepIndex++
        try {
            $list = New-Checklist

            # Capture the state this step starts from, so preservation claims are
            # checked against the workbook's own previous values.
            $costBefore = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
            $inflBefore = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
            $appliedBefore = @(
                (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
                (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
                (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
            )

            # Seed non-zero percentages once a grid exists, so a later shrink has
            # real data to threaten and the destructive path is genuinely exercised.
            if ($costBefore.Count -gt 0 -and $costBefore[0].Count -gt $fixedCost) {
                for ($c = 1; $c -le ($costBefore[0].Count - $fixedCost); $c++) {
                    Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                        -RowIndex 1 -ColumnIndex ($fixedCost + $c) -Value (0.1 * $c)
                }
                $costBefore = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
            }
            if ($inflBefore.Count -gt 0 -and $inflBefore[0].Count -gt $fixedInfl) {
                Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                    -RowIndex 1 -ColumnIndex ($fixedInfl + 1) -Value 0.035
                $inflBefore = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
            }

            Set-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Entered'       -Value $step.entered.base_year
            Set-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Entered'      -Value $step.entered.start_year
            Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered'       -Value $step.entered.duration

            # STRUCTURE CHANGE PENDING must appear the moment an entered value
            # differs from its applied counterpart, with no macro involved.
            $pendingState = Get-NamedValue -Workbook $wb -DefinedName 'nmStructuralState'
            $expectPending = -not (
                ($appliedBefore[0] -eq [string]$step.entered.base_year) -and
                ($appliedBefore[1] -eq [string]$step.entered.start_year) -and
                ($appliedBefore[2] -eq [string]$step.entered.duration)
            )
            if ($expectPending) {
                $null = Add-Check $list 'the structural state indicator reads STRUCTURE CHANGE PENDING before Apply' `
                    ($pendingState -eq $manifest.state_labels.pending) ("read '$pendingState'")
            }

            $excel.Run('PCCM_AutomationBegin', [bool]$step.confirm, '') | Out-Null
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $prompt = [string]$excel.Run('PCCM_AutomationPrompt')
            $outcome = [string]$excel.Run('PCCM_AutomationResult')

            if ($step.expect_rejected) {
                $null = Add-Check $list 'prevalidation rejected the entered timeline' ($outcome -like 'FAIL|*') ("outcome '$outcome'")
            }
            if ($step.expect_destructive_prompt) {
                $null = Add-Check $list 'the confirmation identified data that would be permanently deleted' `
                    ($prompt -match 'PERMANENTLY DELETED') 'the prompt must name the loss before anything moves'
            }

            $appliedAfter = @(
                (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
                (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
                (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
            )
            $expectApplied = @(
                [string]$step.expect.applied.base_year,
                [string]$step.expect.applied.start_year,
                [string]$step.expect.applied.duration
            )
            $null = Add-Check $list 'the applied triple matches the expected structural state' `
                (($appliedAfter -join '/') -eq ($expectApplied -join '/')) `
                ("expected " + ($expectApplied -join '/') + ", got " + ($appliedAfter -join '/'))

            $costHeaders = @(Get-TableColumnNames -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
            $riskHeaders = @(Get-TableColumnNames -Workbook $wb -SheetName $riskGrid.sheet -TableName $riskGrid.table_name)
            $inflHeaders = @(Get-TableColumnNames -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
            $costYears = @($costHeaders | Select-Object -Skip $fixedCost)
            $riskYears = @($riskHeaders | Select-Object -Skip $riskGrid.fixed_columns.Count)
            $inflYears = @($inflHeaders | Select-Object -Skip $fixedInfl)

            $null = Add-Check $list 'cost profiling project-year headers match the oracle' `
                (($costYears -join ',') -eq (@($step.expect.profiling_headers) -join ',')) `
                ("expected " + (@($step.expect.profiling_headers) -join ',') + ", got " + ($costYears -join ','))
            $null = Add-Check $list 'risk profiling project-year headers match the oracle' `
                (($riskYears -join ',') -eq (@($step.expect.profiling_headers) -join ','))
            $null = Add-Check $list 'inflation calendar-year headers match the required span' `
                (($inflYears -join ',') -eq (@($step.expect.inflation_years) -join ',')) `
                ("expected " + (@($step.expect.inflation_years) -join ',') + ", got " + ($inflYears -join ','))

            if ($step.expect.inflation_span_is_empty) {
                $null = Add-Check $list 'the empty inflation span left a well-formed table with its fixed column only' `
                    ($inflHeaders.Count -eq $fixedInfl) ("column count " + $inflHeaders.Count)
            }

            $costAfter = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
            $inflAfter = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name

            if ($step.expect_rejected -or (-not $step.confirm)) {
                # Nothing may have moved.
                $sameCost = (($costBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($costAfter | ForEach-Object { $_ -join '|' }) -join ';')
                $sameInfl = (($inflBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($inflAfter | ForEach-Object { $_ -join '|' }) -join ';')
                $null = Add-Check $list 'the cost profiling grid is logically unchanged' $sameCost
                $null = Add-Check $list 'the inflation grid is logically unchanged' $sameInfl
            } elseif ($step.transition.duration_delta -gt 0) {
                # Growth: every pre-existing project-year value must be identical, and
                # the appended cells must be 0%.
                $preserved = $true
                for ($r = 0; $r -lt $costBefore.Count; $r++) {
                    for ($c = $fixedCost; $c -lt $costBefore[$r].Count; $c++) {
                        if ($costBefore[$r][$c] -ne $costAfter[$r][$c]) { $preserved = $false }
                    }
                }
                $null = Add-Check $list 'existing profiling percentages are unchanged bit for bit' $preserved
                $appendedZero = $true
                for ($r = 0; $r -lt $costAfter.Count; $r++) {
                    if ($costAfter[$r][0] -ne '') {
                        for ($c = ($fixedCost + $costBefore[0].Count - $fixedCost); $c -lt $costAfter[$r].Count; $c++) {
                            if ([double]($costAfter[$r][$c]) -ne 0) { $appendedZero = $false }
                        }
                    }
                }
                $null = Add-Check $list 'each appended project-year cell is 0%' $appendedZero
            } elseif ($step.transition.headers_relabelled_only) {
                $preserved = $true
                for ($r = 0; $r -lt $costBefore.Count; $r++) {
                    for ($c = $fixedCost; $c -lt $costBefore[$r].Count; $c++) {
                        if ($costBefore[$r][$c] -ne $costAfter[$r][$c]) { $preserved = $false }
                    }
                }
                $null = Add-Check $list 'a start-year shift relabels headers and moves no profiling value' $preserved
            }

            # Inflation rates survive by calendar year, not by column index.
            if ((-not $step.expect_rejected) -and $step.confirm -and $inflBefore.Count -gt 0) {
                $survivedOk = $true
                foreach ($year in @($step.transition.added_inflation_years)) {
                    $idx = [array]::IndexOf($inflYears, $year)
                    if ($idx -ge 0) {
                        if ($inflAfter[0][$fixedInfl + $idx] -ne '') { $survivedOk = $false }
                    }
                }
                $null = Add-Check $list 'newly required inflation years arrive BLANK, never zero' $survivedOk
            }

            $report = [string]$excel.Run('PCCM_StructuralReport')
            $null = Add-Check $list 'structural revalidation reports no fault' ([string]::IsNullOrWhiteSpace($report)) $report

            $currentState = Get-NamedValue -Workbook $wb -DefinedName 'nmStructuralState'
            if ((-not $step.expect_rejected) -and $step.confirm) {
                $null = Add-Check $list 'the structural state indicator returns to current after a successful apply' `
                    ($currentState -eq $manifest.state_labels.current) ("read '$currentState'")
            }

            Add-Result ('D-J.' + $stepIndex) $step.title `
                $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) `
                ($step.note + "`r`n" + (Format-Checklist $list))
        } catch {
            Add-Result ('D-J.' + $stepIndex) $step.title 'FAIL' (Format-Err $_)
        }
    }

    # -------------------------------------------------------------------
    # K. Profiling synchronisation
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $costIds = Get-IdColumnValues -Workbook $wb -Info $costReg
        $profileIds = @()
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'every identified cost line has exactly one profiling row' `
            (($costIds -join ',') -eq ($profileIds -join ',')) `
            ("register " + ($costIds -join ',') + " / grid " + ($profileIds -join ','))
        $null = Add-Check $list 'no profiling identifier appears twice' `
            ((@($profileIds | Select-Object -Unique)).Count -eq $profileIds.Count)

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $afterAdd = @()
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $afterAdd += $row[0] }
        }
        $null = Add-Check $list 'adding a driver creates exactly one matching profiling row' `
            ($afterAdd.Count -eq ($profileIds.Count + 1)) ("grid rows " + $afterAdd.Count)

        $victim = $afterAdd[$afterAdd.Count - 1]
        $excel.Run('PCCM_DeleteCostLineById', $victim) | Out-Null
        $afterDelete = @()
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $afterDelete += $row[0] }
        }
        $null = Add-Check $list 'deleting a driver removes its profiling row' ($afterDelete -notcontains $victim)

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports no fault' ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'K' 'Profiling synchronisation by permanent ID' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'K' 'Profiling synchronisation' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # L. Runtime failure containment
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $costBefore = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $inflBefore = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
        $appliedBefore = @(
            (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
        )

        # Arm a failure AFTER mutation has begun: the applied triple has already
        # been written and the profiling columns already reshaped.
        $excel.Run('PCCM_AutomationBegin', $true, 'apply.after_profiling_columns') | Out-Null
        Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered' -Value ([double]$appliedBefore[2] + 2)
        $excel.Run('PCCM_ApplyTimeline') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected mid-operation failure was reported, not swallowed' ($outcome -like 'FAIL|*') ("outcome '$outcome'")

        $appliedAfter = @(
            (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
        )
        $null = Add-Check $list 'the applied triple was logically restored' `
            (($appliedBefore -join '/') -eq ($appliedAfter -join '/')) `
            ("before " + ($appliedBefore -join '/') + ", after " + ($appliedAfter -join '/'))

        $costAfter = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $inflAfter = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
        $null = Add-Check $list 'the cost profiling grid was logically restored' `
            ((($costBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($costAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the inflation grid was logically restored' `
            ((($inflBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($inflAfter | ForEach-Object { $_ -join '|' }) -join ';'))

        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'the restored workbook still passes structural revalidation' ([string]::IsNullOrWhiteSpace($report)) $report

        $null = Add-Check $list 'application state was restored (ScreenUpdating is on)' ([bool]$excel.ScreenUpdating)
        $null = Add-Check $list 'application state was restored (EnableEvents is on)' ([bool]$excel.EnableEvents)

        Add-Result 'L' 'Runtime failure containment and logical restore' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'L' 'Runtime failure containment' 'FAIL' (Format-Err $_)
    }

    $excel.Run('PCCM_AutomationEnd') | Out-Null
  } catch {
    Add-Result 'XX' 'Driving the Stage-B workbook' 'FAIL' (Format-Err $_)
  }

  # --- shutdown, leaf before parent ---------------------------------------
  $rel = New-ReleaseLedger 'functional test instance'
  try {
      Invoke-NamedRelease $rel $worksheets 'Worksheets'; $worksheets = $null
      if ($null -ne $wb) {
          try { $wb.Close($false); $rel.WorkbookClosed = $true } catch { $null = $rel.Failed.Add('Workbook.Close') }
      }
      Invoke-NamedRelease $rel $wb        'Workbook';  $wb        = $null
      Invoke-NamedRelease $rel $workbooks 'Workbooks'; $workbooks = $null
      if ($null -ne $excel) {
          try { $excel.Quit(); $rel.QuitCalled = $true } catch { $null = $rel.Failed.Add('Application.Quit') }
      }
      Invoke-NamedRelease $rel $excel 'Application'; $excel = $null

      [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
      [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

      $rel.NaturalExit = Wait-ExcelExit -Identity $id
      if (-not $rel.NaturalExit) {
          $rel.EmergencyRequired = $true
          Add-Note (Invoke-EmergencyExcelCleanup -Identity $id -Label 'functional test instance')
      }
  } catch {
      Add-Note ('Shutdown raised: ' + (Format-Err $_))
  }

  if ($rel.NaturalExit -and $rel.Failed.Count -eq 0) {
      Add-Result 'Z' 'Excel closed naturally after the functional run' 'PASS' ("pid {0} exited without a forced stop" -f $id.ProcessId)
  } else {
      Add-Result 'Z' 'Excel closed naturally after the functional run' 'FAIL' `
          ("natural exit={0}; failed releases={1}" -f $rel.NaturalExit, (($rel.Failed | Select-Object -Unique) -join ', '))
  }
} else {
    Add-Result 'B-Z' 'Structural runtime scenarios' 'SKIP' 'the Stage-B build did not complete'
}

# ===========================================================================
# Report
# ===========================================================================
$transient = Get-TransientFailures
if ($transient.Count -gt 0) {
    Add-Result 'Y' 'Transient COM releases' 'FAIL' ($transient -join '; ')
} else {
    Add-Result 'Y' 'Transient COM releases' 'PASS' 'every transient object released cleanly'
}

Write-Host ''
Write-Host 'Shutdown ledger' -ForegroundColor Cyan
Write-Host '---------------'
Write-Host (Format-ReleaseLedger $rel)

if ($notes.Count -gt 0) {
    Write-Host ''
    Write-Host 'Notes' -ForegroundColor Yellow
    Write-Host '-----'
    foreach ($n in $notes) { Write-Host ("  " + $n) }
}

if (-not $KeepArtifacts) {
    try { Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue } catch { }
} else {
    Add-Note ("Artifacts kept at " + $tempRoot)
}

$failed = @($results | Where-Object { $_.Status -eq 'FAIL' })
Write-Host ''
Write-Host ("  {0} passed, {1} failed, {2} skipped" -f `
    (@($results | Where-Object { $_.Status -eq 'PASS' })).Count, `
    $failed.Count, `
    (@($results | Where-Object { $_.Status -eq 'SKIP' })).Count)
Write-Host ''
if ($failed.Count -eq 0) {
    Write-Host 'PHASE-4 FUNCTIONAL TEST: ALL CHECKS PASSED' -ForegroundColor Green
    Write-Host ''
    exit 0
}
Write-Host ("PHASE-4 FUNCTIONAL TEST FAILED: " + (($failed | ForEach-Object { $_.Id }) -join ', ')) -ForegroundColor Red
Write-Host ''
exit 1
