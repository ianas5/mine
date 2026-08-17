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
      A   Stage-B build: .xlsm, FileFormat 52, 14 CodeNames, modules, five buttons,
          natural COM shutdown
      B   Permanent Cost Line IDs: sequence and non-reuse after deletion
      B2  A REAL ListObject reorder: identity travels with its own row data
      C   Permanent Risk IDs, independently sequenced
      D   First timeline application
      E   Duration increase
      F   Start-year shift with unchanged duration, and BLANK preservation
      G   Duration decrease: destructive confirmation, cancelled then accepted
      H   Base-Year movement, earlier and later
      I   Combined three-way change
      J   Degenerate inflation span
      K   Profiling synchronisation by permanent ID
      L   Apply failure after mutation: logical restore
      M   Growth beyond the 25 reserved rows, with presentation and Data Validation
          proved on the runtime-created row
      N   An Inflation Profile removed from Config: destructive, cancelled then
          accepted, with the timeline unchanged throughout
      O   Non-numeric content in a removed profiling cell counts as data loss
      P   Oversized pasted timeline values rejected without a VBA overflow
      Q   Add failure after row mutation: rows, values and counter restored
      R   Delete failure after row mutation: rows, values and counter restored

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
# Every helper below acquires each COM object into its OWN named variable and
# releases it in the narrowest scope. There are no chained member expressions such
# as $Workbook.Names.Item(...) or $body.Rows.Count: each of those creates an
# intermediate RCW that nothing owns and nothing releases, which is precisely the
# pattern the Phase-1.6 readiness gate ruled out.
function Get-NamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $v = $rng.Value2
        if ($null -eq $v) { return '' }
        return [string]$v
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Set-NamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Get-TableColumnNames {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $cols = $null
    $out = @()
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $colCount = [int]$cols.Count
        for ($i = 1; $i -le $colCount; $i++) {
            $c = $null
            try { $c = $cols.Item($i); $out += [string]$c.Name }
            finally { if ($null -ne $c) { Release-Transient $c 'ListColumn'; $c = $null } }
        }
    } finally {
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
    return $out
}

# Reads the whole data body of a table as a jagged array of strings, so a later
# comparison is a plain value comparison and never a live COM read.
function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null
    $rowsObj = $null; $colsObj = $null
    $rows = @()
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        if ($null -eq $body) { return $rows }

        # Row and column counts are read through named, released objects rather than
        # through $body.Rows.Count, which would mint an unowned Range on every
        # iteration of the loop.
        $rowsObj = $body.Rows
        $colsObj = $body.Columns
        $rowCount = [int]$rowsObj.Count
        $colCount = [int]$colsObj.Count
        Release-Transient $rowsObj 'Range(rows)'; $rowsObj = $null
        Release-Transient $colsObj 'Range(columns)'; $colsObj = $null

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = @()
            for ($c = 1; $c -le $colCount; $c++) {
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
        if ($null -ne $rowsObj)         { Release-Transient $rowsObj         'Range(rows)';    $rowsObj         = $null }
        if ($null -ne $colsObj)         { Release-Transient $colsObj         'Range(columns)'; $colsObj         = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';    $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';     $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';    $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';      $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';     $localWorksheets = $null }
    }
    return $rows
}

function Set-TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        if ($null -eq $Value) {
            # A genuine blank, not zero. The two are different assumptions and the
            # harness has to be able to create each of them deliberately.
            $null = $cell.ClearContents()
        } elseif ($Value -is [string]) {
            $cell.Value2 = [string]$Value
        } else {
            $cell.Value2 = [double]$Value
        }
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# A REAL table reorder. Sorting the ListObject moves whole rows, which is what the
# permanent-ID invariant is actually about: an identifier must travel with its own
# row data. Writing a marker into row 1 and reading row 1 back proves nothing.
function Invoke-TableSort {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$KeyColumnIndex, [int]$Order)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $sortObj = $null; $sortFields = $null; $body = $null; $keyRange = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $keyRange = $body.Columns($KeyColumnIndex)
        $sortObj = $lo.Sort
        $sortFields = $sortObj.SortFields
        $sortFields.Clear()
        $null = $sortFields.Add($keyRange, 0, $Order)
        $sortObj.Apply()
        $sortFields.Clear()
    } finally {
        if ($null -ne $sortFields)      { Release-Transient $sortFields      'SortFields';  $sortFields      = $null }
        if ($null -ne $sortObj)         { Release-Transient $sortObj         'Sort';        $sortObj         = $null }
        if ($null -ne $keyRange)        { Release-Transient $keyRange        'Range(key)';  $keyRange        = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# Fill colour of one table cell, as a six-digit RGB hex string, so the harness can
# assert that a row created at RUNTIME still carries the Phase-2 input language.
function Get-TableCellFill {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $body = $null; $cell = $null; $interior = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        $interior = $cell.Interior
        $bgr = [int]$interior.Color
        $r = $bgr -band 0xFF
        $g = ($bgr -shr 8) -band 0xFF
        $b = ($bgr -shr 16) -band 0xFF
        return ('{0:X2}{1:X2}{2:X2}' -f $r, $g, $b)
    } finally {
        if ($null -ne $interior)        { Release-Transient $interior        'Interior';    $interior        = $null }
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# $true when the cell carries Data Validation. Excel raises when a cell has none, so
# the absence is detected by the raise rather than by a sentinel value.
function Test-TableCellValidation {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $body = $null; $cell = $null; $validation = $null
    $present = $false
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        try {
            $validation = $cell.Validation
            $null = $validation.Type
            $present = $true
        } catch {
            $present = $false
        }
    } finally {
        if ($null -ne $validation)      { Release-Transient $validation      'Validation';  $validation      = $null }
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
    return $present
}

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return (Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
}

function Add-ConfigProfile {
    param($Workbook, [string]$ProfileName, [int]$RowIndex)
    Set-TableCell -Workbook $Workbook -SheetName 'Config' -TableName 'tblInflationProfiles' `
        -RowIndex $RowIndex -ColumnIndex 1 -Value $ProfileName
}

function Clear-ConfigProfile {
    param($Workbook, [int]$RowIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item('Config')
        $los = $ws.ListObjects
        $lo = $los.Item('tblInflationProfiles')
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, 1)
        $null = $cell.ClearContents()
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
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

        Add-Result 'B' 'Permanent Cost Line IDs: sequence and non-reuse' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'B' 'Permanent Cost Line IDs' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # B2. A REAL reorder of the Cost Lines table
    # -------------------------------------------------------------------
    # The invariant is that a permanent identifier travels with its OWN row data
    # when whole rows move. Writing a marker into row 1 and reading row 1 back does
    # not move anything and therefore tests nothing.
    try {
        $list = New-Checklist

        # Give each identified row a distinct, sortable marker in the Description
        # column, and a distinct profiling percentage in project year 1 if a grid
        # exists yet.
        $before = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $markers = @{}
        $rowIndex = 0
        foreach ($row in $before) {
            $rowIndex++
            if ($row[0] -ne '') {
                $marker = 'MARKER-' + $row[0]
                Set-TableCell -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
                    -RowIndex $rowIndex -ColumnIndex 3 -Value $marker
                $markers[$row[0]] = $marker
            }
        }
        $countersBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'

        # Sort descending by Description so the physical row order genuinely changes.
        Invoke-TableSort -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
            -KeyColumnIndex 3 -Order 2

        $after = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $orderBefore = @(); foreach ($row in $before) { if ($row[0] -ne '') { $orderBefore += $row[0] } }
        $orderAfter  = @(); foreach ($row in $after)  { if ($row[0] -ne '') { $orderAfter  += $row[0] } }

        $null = Add-Check $list 'the physical row order actually changed' `
            (($orderBefore -join ',') -ne ($orderAfter -join ',')) `
            ("before " + ($orderBefore -join ',') + " / after " + ($orderAfter -join ','))
        $null = Add-Check $list 'the same set of identifiers is still present' `
            ((($orderBefore | Sort-Object) -join ',') -eq (($orderAfter | Sort-Object) -join ','))

        $travelled = $true
        foreach ($row in $after) {
            if ($row[0] -ne '') {
                if ($markers[$row[0]] -ne $row[2]) { $travelled = $false }
            }
        }
        $null = Add-Check $list 'each permanent ID still sits on its own row data' $travelled

        $null = Add-Check $list 'no identifier was regenerated by the reorder' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $countersBefore)

        # Profiling ownership must follow the ID, not the row position.
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $excel.Run('PCCM_DeleteCostLineById', (Get-IdColumnValues -Workbook $wb -Info $costReg)[-1]) | Out-Null
        $profileIds = @()
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'the profiling grid follows the reordered register by ID' `
            (($profileIds -join ',') -eq ($orderAfter -join ',')) `
            ("register " + ($orderAfter -join ',') + " / grid " + ($profileIds -join ','))

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports no fault after the reorder' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'B2' 'Real table reorder: identity travels with its row' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'B2' 'Real table reorder' 'FAIL' (Format-Err $_)
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
                # Deliberately leave one identified row's first project year BLANK.
                # It is invalid data, and it must SURVIVE every structural operation
                # so Model Check can report it later. Silently filling it with 0%
                # would repair user data inside a structural synchronisation.
                if ($costBefore.Count -ge 2) {
                    if ($costBefore[1][0] -ne '') {
                        Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                            -RowIndex 2 -ColumnIndex ($fixedCost + 1) -Value $null
                    }
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

                # A blank is a value too. Synchronisation must NOT quietly fill it
                # with 0%: an invalid blank has to survive for Model Check to report.
                $blankSurvived = $true
                for ($r = 0; $r -lt $costBefore.Count; $r++) {
                    for ($c = $fixedCost; $c -lt $costBefore[$r].Count; $c++) {
                        if ($costBefore[$r][$c] -eq '' -and $costAfter[$r][$c] -ne '') { $blankSurvived = $false }
                    }
                }
                $null = Add-Check $list 'an existing BLANK profiling cell is still blank after synchronisation' `
                    $blankSurvived 'structural synchronisation must not repair invalid user data'
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

    # -------------------------------------------------------------------
    # M. Real table growth beyond the reserved capacity
    # -------------------------------------------------------------------
    # Stage A materialises 25 reserved rows. Until a 26th identified driver exists,
    # the ListRows.Add path in modDrivers has never executed, and neither has
    # whatever Excel does to a row it creates inside a ListObject.
    try {
        $list = New-Checklist
        $reserved = [int]$costReg.reserved_rows
        $existing = @(Get-IdColumnValues -Workbook $wb -Info $costReg).Count

        for ($i = $existing; $i -le $reserved; $i++) {
            $excel.Run('PCCM_AddCostLine') | Out-Null
        }
        $ids = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $rowCount = Get-TableRowCount -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name

        $null = Add-Check $list 'the register now holds more identified rows than its reserved capacity' `
            ($ids.Count -gt $reserved) ("identified " + $ids.Count + " / reserved " + $reserved)
        $null = Add-Check $list 'the ListObject itself grew' ($rowCount -gt $reserved) ("rows " + $rowCount)

        $newId = $ids[$ids.Count - 1]
        $null = Add-Check $list 'the grown row carries a valid next permanent ID' `
            ($newId -match '^CL-\d{3,}$') $newId
        $null = Add-Check $list 'no identifier is duplicated after growth' `
            ((@($ids | Select-Object -Unique)).Count -eq $ids.Count)

        $profileIds = @()
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'the grown driver has exactly one profiling row' `
            (@($profileIds | Where-Object { $_ -eq $newId }).Count -eq 1)

        # Presentation and validation on the RUNTIME-created row. If Excel Table
        # propagation is what supplies these, this is where that reliance is proved
        # rather than assumed.
        $grownRow = 0
        $rowIdx = 0
        foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
            $rowIdx++
            if ($row[0] -eq $newId) { $grownRow = $rowIdx }
        }
        $idFill = Get-TableCellFill -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name -RowIndex $grownRow -ColumnIndex 1
        $userFill = Get-TableCellFill -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name -RowIndex $grownRow -ColumnIndex 2
        $null = Add-Check $list 'the grown row ID cell keeps the model-controlled treatment' `
            ($idFill -eq $manifest.presentation.locked_fill) ("fill " + $idFill)
        $null = Add-Check $list 'the grown row user cells keep the editable input treatment' `
            ($userFill -eq $manifest.presentation.input_fill) ("fill " + $userFill)
        $null = Add-Check $list 'the grown row ID cell carries NO user Data Validation' `
            (-not (Test-TableCellValidation -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name -RowIndex $grownRow -ColumnIndex 1))

        $validationOk = $true
        foreach ($key in $manifest.driver_validation_columns.cost_lines) {
            $columnIndex = [array]::IndexOf($costReg.columns, $key) + 1
            if (-not (Test-TableCellValidation -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name -RowIndex $grownRow -ColumnIndex $columnIndex)) {
                $validationOk = $false
            }
        }
        $null = Add-Check $list 'every validated user column keeps its Data Validation on the grown row' $validationOk

        # Dynamic year cells must also read as editable input, not as locked cells.
        if ((Get-TableColumnNames -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name).Count -gt $fixedCost) {
            $gridRow = 0; $rowIdx = 0
            foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                $rowIdx++
                if ($row[0] -eq $newId) { $gridRow = $rowIdx }
            }
            if ($gridRow -gt 0) {
                $yearFill = Get-TableCellFill -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name -RowIndex $gridRow -ColumnIndex ($fixedCost + 1)
                $null = Add-Check $list 'a generated profiling year cell is visually an editable input' `
                    ($yearFill -ne $manifest.presentation.locked_fill) ("fill " + $yearFill)
            }
        }

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports no fault after growth' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'M' 'Table growth beyond reserved capacity, with presentation and validation intact' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'M' 'Table growth beyond reserved capacity' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # N. An Inflation Profile removed from Config is destructive
    # -------------------------------------------------------------------
    # A second, INDEPENDENT loss mechanism: deleting a profile from the Config
    # master destroys that row's annual rates on the next synchronisation even when
    # Base Year, Start Year and Duration are completely unchanged.
    try {
        $list = New-Checklist
        $profileName = 'HARNESS TEMP PROFILE'
        $profileRow = [int]$inflGrid.reserved_rows

        Add-ConfigProfile -Workbook $wb -ProfileName $profileName -RowIndex $profileRow
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $excel.Run('PCCM_ApplyTimeline') | Out-Null

        $inflRows = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
        $targetRow = 0; $rowIdx = 0
        foreach ($row in $inflRows) { $rowIdx++; if ($row[0] -eq $profileName) { $targetRow = $rowIdx } }
        $null = Add-Check $list 'the new Config profile gained an inflation row' ($targetRow -gt 0)

        if ($targetRow -gt 0 -and $inflRows[0].Count -gt $fixedInfl) {
            Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                -RowIndex $targetRow -ColumnIndex ($fixedInfl + 1) -Value 0.042
            $inflBefore = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name

            # Remove the profile from Config. The timeline is NOT touched.
            Clear-ConfigProfile -Workbook $wb -RowIndex $profileRow

            # --- cancelled -------------------------------------------------
            $excel.Run('PCCM_AutomationBegin', $false, '') | Out-Null
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $prompt = [string]$excel.Run('PCCM_AutomationPrompt')
            $null = Add-Check $list 'the confirmation names the removed profile' `
                ($prompt -match [regex]::Escape($profileName)) 'the prompt must identify the profile leaving Config'
            $null = Add-Check $list 'the confirmation reports permanent deletion' `
                ($prompt -match 'PERMANENTLY DELETED')
            $null = Add-Check $list 'the confirmation counts the rates that would be lost' `
                ($prompt -match 'inflation rates lost')

            $inflAfterCancel = Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name
            $null = Add-Check $list 'cancelling leaves the inflation row and all its rates unchanged' `
                ((($inflBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($inflAfterCancel | ForEach-Object { $_ -join '|' }) -join ';'))

            # --- accepted --------------------------------------------------
            $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $remaining = @()
            foreach ($row in (Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)) {
                if ($row[0] -ne '') { $remaining += $row[0] }
            }
            $null = Add-Check $list 'accepting removes the obsolete inflation row' ($remaining -notcontains $profileName)
        }

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports no fault' ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'N' 'Removed Config inflation profile is assessed as destructive' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'N' 'Removed Config inflation profile' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # O. Non-numeric content in a removed profiling cell is a data loss
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $appliedDuration = [double](Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
        if ($appliedDuration -ge 2) {
            $ids = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
            $gridRow = 0; $rowIdx = 0
            foreach ($row in (Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                $rowIdx++; if ($row[0] -eq $ids[0]) { $gridRow = $rowIdx }
            }
            # Zero the tail first, so ONLY the pasted text can make this destructive.
            for ($r = 1; $r -le (Get-TableRowCount -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name); $r++) {
                Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                    -RowIndex $r -ColumnIndex ($fixedCost + [int]$appliedDuration) -Value 0
            }
            for ($r = 1; $r -le (Get-TableRowCount -Workbook $wb -SheetName $riskGrid.sheet -TableName $riskGrid.table_name); $r++) {
                Set-TableCell -Workbook $wb -SheetName $riskGrid.sheet -TableName $riskGrid.table_name `
                    -RowIndex $r -ColumnIndex ($fixedCost + [int]$appliedDuration) -Value 0
            }
            Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                -RowIndex $gridRow -ColumnIndex ($fixedCost + [int]$appliedDuration) -Value 'PASTED TEXT'

            Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered' -Value ($appliedDuration - 1)
            $excel.Run('PCCM_AutomationBegin', $false, '') | Out-Null
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $prompt = [string]$excel.Run('PCCM_AutomationPrompt')

            $null = Add-Check $list 'text in a removed profiling cell triggers a destructive warning' `
                ($prompt -match 'PERMANENTLY DELETED') `
                'blank and numeric zero are not data; anything else is'
            $null = Add-Check $list 'the affected permanent ID is named' ($prompt -match [regex]::Escape($ids[0]))

            $stillThere = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
            $null = Add-Check $list 'cancelling leaves the pasted value in place' `
                ($stillThere[$gridRow - 1][$fixedCost + [int]$appliedDuration - 1] -eq 'PASTED TEXT')

            # Clean up so later sections start from a numeric grid.
            Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                -RowIndex $gridRow -ColumnIndex ($fixedCost + [int]$appliedDuration) -Value 0
            Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered' -Value $appliedDuration
        } else {
            $null = Add-Check $list 'the applied duration is long enough to shrink' $false "duration $appliedDuration"
        }

        Add-Result 'O' 'Non-numeric profiling content counts as destructive data loss' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'O' 'Non-numeric profiling content' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # P. An oversized pasted timeline value is rejected cleanly
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $appliedBefore = @(
            (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
        )
        $enteredBefore = @(
            (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Entered'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Entered'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered')
        )

        # A whole number far beyond a VBA Long. Data Validation does not stop a paste.
        Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered' -Value 10000000000
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $excel.Run('PCCM_ApplyTimeline') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')

        $null = Add-Check $list 'an oversized Duration is rejected by prevalidation, not by an overflow' `
            ($outcome -like 'FAIL|*') ("outcome '$outcome'")
        $null = Add-Check $list 'the rejection names the structural protection limit' `
            ($outcome -match 'structural protection limit')

        $appliedAfter = @(
            (Get-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Applied'),
            (Get-NamedValue -Workbook $wb -DefinedName 'nmDuration_Applied')
        )
        $null = Add-Check $list 'the applied state is untouched' `
            (($appliedBefore -join '/') -eq ($appliedAfter -join '/'))

        # An oversized YEAR must be rejected before any arithmetic uses it.
        Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered' -Value $enteredBefore[2]
        Set-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Entered' -Value 99999999
        $excel.Run('PCCM_ApplyTimeline') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'an oversized Start Year is rejected cleanly' ($outcome -like 'FAIL|*') ("outcome '$outcome'")
        $null = Add-Check $list 'the rejection names the supported calendar-year range' `
            ($outcome -match 'outside the supported range')

        Set-NamedValue -Workbook $wb -DefinedName 'nmBaseYear_Entered'  -Value $enteredBefore[0]
        Set-NamedValue -Workbook $wb -DefinedName 'nmStartYear_Entered' -Value $enteredBefore[1]
        Set-NamedValue -Workbook $wb -DefinedName 'nmDuration_Entered'  -Value $enteredBefore[2]

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports no fault' ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'P' 'Oversized pasted timeline values are rejected without overflow' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'P' 'Oversized pasted timeline values' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # Q. Add failure after row mutation has begun
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $registerBefore = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $gridBefore = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
        $idsBefore = @(Get-IdColumnValues -Workbook $wb -Info $costReg)

        # Armed AFTER the identifier has been allocated and written, so the failure
        # lands with the register already mutated.
        $excel.Run('PCCM_AutomationBegin', $true, 'add.after_write_id') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected Add failure was reported' ($outcome -like 'FAIL|*') ("outcome '$outcome'")

        $registerAfter = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $gridAfter = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $null = Add-Check $list 'the driver table row count was restored' ($registerAfter.Count -eq $registerBefore.Count) `
            ("before " + $registerBefore.Count + " / after " + $registerAfter.Count)
        $null = Add-Check $list 'the profiling row count was restored' ($gridAfter.Count -eq $gridBefore.Count)
        $null = Add-Check $list 'the register values were restored' `
            ((($registerBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($registerAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the profiling values were restored' `
            ((($gridBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($gridAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the ID counter was restored' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $counterBefore)
        $idsAfter = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'no identifier issued by the failed Add survives' `
            (($idsAfter -join ',') -eq ($idsBefore -join ','))

        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'the restored workbook still passes structural revalidation' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'Q' 'Add failure after row mutation: full logical restore' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'Q' 'Add failure after row mutation' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # R. Delete failure after row mutation has begun
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $registerBefore = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $gridBefore = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
        $victim = @(Get-IdColumnValues -Workbook $wb -Info $costReg)[0]

        $excel.Run('PCCM_AutomationBegin', $true, 'delete.after_remove') | Out-Null
        $excel.Run('PCCM_DeleteCostLineById', $victim) | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected Delete failure was reported' ($outcome -like 'FAIL|*') ("outcome '$outcome'")

        $registerAfter = Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name
        $gridAfter = Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name
        $null = Add-Check $list 'the driver table row count was restored' ($registerAfter.Count -eq $registerBefore.Count) `
            ("before " + $registerBefore.Count + " / after " + $registerAfter.Count)
        $null = Add-Check $list 'the profiling row count was restored' ($gridAfter.Count -eq $gridBefore.Count)
        $null = Add-Check $list 'the deleted identifier is back' `
            ((Get-IdColumnValues -Workbook $wb -Info $costReg) -contains $victim)
        $null = Add-Check $list 'the register values were restored' `
            ((($registerBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($registerAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the ID counter was restored' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $counterBefore)

        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'the restored workbook still passes structural revalidation' `
            ([string]::IsNullOrWhiteSpace($report)) $report
        $null = Add-Check $list 'application state was restored (ScreenUpdating is on)' ([bool]$excel.ScreenUpdating)

        Add-Result 'R' 'Delete failure after row mutation: full logical restore' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'R' 'Delete failure after row mutation' 'FAIL' (Format-Err $_)
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
