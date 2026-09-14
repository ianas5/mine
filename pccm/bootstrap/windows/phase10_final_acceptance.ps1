<#
.SYNOPSIS
    PCCM Phase 10 - the FINAL WINDOWS ACCEPTANCE runner for the release workbook.
.DESCRIPTION
    THE CONTRACTED HARDENING RUNNER, IN THE ACCEPTED P9-1 SHAPE. One owned Excel
    instance, one disposable Stage-B build under %TEMP%, one workbook, one clean
    shutdown. It asks the release workbook the questions no accepted runner has
    asked: does the built artifact carry its release identity and the Source
    Revision the builder stamped; do Reset Results and Repair Profiling behave
    through their REAL entry points; is protection intact after every success,
    refusal and injected failure; and does every accepted state still present.

    REUSED, NOT REINVENTED. build_stage_b.ps1 builds and reopen-verifies the
    workbook (FileFormat, every sheet and CodeName, every module, every button and
    its OnAction). com_lifecycle.ps1 owns COM. phase5_gate_b_scenarios.ps1 owns
    the accepted W4-sized fixture and the production-operation invoker.
    phase6_gate_b_scenarios.ps1 owns the simulation invoker and the persisted
    simulation-state reader. phase10_fixture_window.bas is the accepted setup-only
    protection window, imported into the DISPOSABLE copy exactly as the benchmark
    imports it, and used here only to create Repair preconditions and to let the
    accepted fixture write where it always wrote.

    EVERY EXPECTATION COMES FROM A PROJECTION. Sheets, CodeNames and modules from
    stage_b_manifest.json; the metadata rows from
    phase10_methodology_inspection.json; the protection sheet set from
    phase10_protection_inspection.json; what Reset clears and preserves from
    phase10_reset_inspection.json; the state vocabulary from
    phase7_acceptance_inspection.json; the Model Check surface from
    phase9_model_check_inspection.json. No inventory is a literal in this file.
    A count that is reported was derived first.

    FAIL FAST. A failed acceptance check throws; the throw reaches the one
    shutdown path; the verdict is FAIL; the exit code is 1. Nothing after a
    failure is reported as evidence.

    ITERATIONS. The business minimum for Simulation, Sensitivity and Annual, and
    the minimum plus one exactly once, to produce STALE. No larger request exists
    in this file.

    NO UI AUTOMATION. Confirmation and refusal are driven through
    PCCM_AutomationBegin, the accepted automation seam, which answers the
    confirmation deterministically and injects the contracted failpoint.
.NOTES
    THIS FILE HAS NOT RUN. There is no Windows and no Excel where it was
    written, and nothing in it claims otherwise.
#>
[CmdletBinding()]
param(
    [string]$BuildDir
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# DEFINITION-ONLY AT TOP LEVEL. Dot-sourcing defines functions and script
# variables and runs no scenario. Nothing here calls Invoke-Phase5GateBScenarios
# or Invoke-Phase6GateBScenarios.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

$script:FaLines = New-Object System.Collections.ArrayList
$script:FaPath = ''
$script:FaChecks = New-Object System.Collections.ArrayList
$script:FaResidual = New-Object System.Collections.ArrayList
$script:FaErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}
# The accepted setup-only protection window shim, exactly as the benchmark
# imports it. Never declared in the manifest; refused if it ever is.
$script:FixtureWindowModule = 'modPhase10FixtureWindow'
$script:FixtureWindowSource = 'phase10_fixture_window.bas'
# The contracted Reset failpoint this runner injects: after the simulation clear,
# so both the calculation and the simulation owners have something to put back.
$script:ResetFailpoint = 'Phase10ResetSimulation'
# The two contracted Repair failpoints: after the cost grid and after the risk grid.
$script:RepairFailpointCost = 'Phase10RepairCost'
$script:RepairFailpointRisk = 'Phase10RepairRisk'
# The grid digests the Repair contract scenarios compare against; set by Save-FaGridsBefore.
$script:FaCostBefore = ''
$script:FaRiskBefore = ''
# The contracted Workbook_Open failpoint (matrix row O): the handler's own constant.
$script:OpenFailpoint = 'Phase10WorkbookOpen'

# ===========================================================================
# THE HELPERS THE DOT-SOURCED FILES CALL, AND THIS RUNNER DEFINES
# ===========================================================================
# phase5_gate_b_scenarios.ps1 and phase6_gate_b_scenarios.ps1 call these by name
# and define none of them: their definitions live in the Phase-4 driver this
# runner deliberately does not dot-source. They are the accepted P9-1 copies,
# verbatim, so tests/powershell_command_resolution_audit.ps1 resolves every
# reachable command exactly once.
function Write-RowObject {
    param([object[]]$Row)
    Write-Output -NoEnumerate $Row
}

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

function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null
    $rowsObj = $null; $colsObj = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        # An empty body is a valid outcome: emit NOTHING. The caller's @(...) turns
        # zero pipeline objects into an empty collection, which is exactly right.
        if ($null -eq $body) { return }

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
            Write-RowObject $line
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
    # No trailing return: every row has already been emitted, one object each.
}

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
}

function Add-BlankTableRow {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $added = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $added = $rows.Add()
        return [int]$added.Index
    } finally {
        if ($null -ne $added)           { Release-Transient $added           'ListRow';     $added           = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Remove-TableRow {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $victim = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $victim = $rows.Item($RowIndex)
        $victim.Delete()
    } finally {
        if ($null -ne $victim)          { Release-Transient $victim          'ListRow';     $victim          = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# A PROJECT-YEAR COLUMN ADDED TO OR REMOVED FROM A GRID - used ONLY inside the
# accepted setup window to create the width preconditions Repair Profiling is
# asked to repair or refuse. Structural, so production's own window guards it.
function Add-FaTableColumn {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $cols = $null; $added = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $added = $cols.Add()
        return [int]$added.Index
    } finally {
        if ($null -ne $added)           { Release-Transient $added           'ListColumn';  $added           = $null }
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Remove-FaLastTableColumn {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $cols = $null; $victim = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $victim = $cols.Item([int]$cols.Count)
        $victim.Delete()
    } finally {
        if ($null -ne $victim)          { Release-Transient $victim          'ListColumn';  $victim          = $null }
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# One typed weight, or a blank, written to a grid cell by (row, project year).
function Set-FaWeight {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$FixedColumns, [int]$RowIndex, [int]$Year, $Weight)
    if ($null -eq $Weight) {
        Set-TableCell -Workbook $Workbook -SheetName $SheetName -TableName $TableName -RowIndex $RowIndex -ColumnIndex ($FixedColumns + $Year) -Value $null
    } else {
        Set-TableCell -Workbook $Workbook -SheetName $SheetName -TableName $TableName -RowIndex $RowIndex -ColumnIndex ($FixedColumns + $Year) -Value ([double]$Weight)
    }
}

function Get-IdColumnValues {
    param($Workbook, $Info)
    $out = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}

# ===========================================================================
# REPORTING - ONE LINE PER CHECK, FAIL FAST
# ===========================================================================
function Write-FaLine {
    param([string]$Text = '')
    $null = $script:FaLines.Add($Text)
    Write-Host $Text
    if (-not [string]::IsNullOrWhiteSpace($script:FaPath)) {
        try {
            Set-Content -LiteralPath $script:FaPath `
                -Value ($script:FaLines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# EVERY CHECK IS A LINE `PASS|scenario|detail` OR `FAIL|scenario|detail`. A
# failed acceptance check THROWS, so nothing after a failure is reported as
# evidence; -Continue is for the shutdown checks, which run after the session
# and must all be reported together.
function Add-FaCheck {
    param([string]$Scenario, [bool]$Ok, [string]$Detail = '', [switch]$Continue)
    $null = $script:FaChecks.Add([pscustomobject]@{
        Scenario = $Scenario; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    Write-FaLine ($verdict + '|' + $Scenario + '|' + $Detail)
    if ((-not $Ok) -and (-not $Continue)) {
        throw ('ACCEPTANCE FAILED at ' + $Scenario + ': ' + $Detail)
    }
    return $Ok
}

function Invoke-FaRelease {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:FaResidual.Add(
                ([string]$rec.Label + ' left ' + [string]$rec.Count + ' outstanding'))
        }
    } elseif ($rec.Status -eq 'FAIL') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $null = $Ledger.Failed.Add($rec.Label)
        $null = $Ledger.Lines.Add(("      {0,-24} | FAIL    | {1}" -f $rec.Label, $rec.Error))
    } else {
        $null = $Ledger.Lines.Add(("      {0,-24} | SKIPPED | {1}" -f $rec.Label, $rec.Error))
    }
}

# ===========================================================================
# THE SOURCE REVISION THE BUILDER STAMPS, DERIVED THE WAY THE BUILDER DERIVES IT
# ===========================================================================
# builder/pccm_builder/workbook_builder.py stamps `<short HEAD> (clean|dirty)`,
# clean meaning `git status --porcelain --untracked-files=no` is empty for the
# whole tree. This runner refuses a dirty tree before Excel starts: the workbook
# it would accept could not be attributed to one commit.
function Get-FaSourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse --short HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; the release workbook cannot be attributed to a source revision')
    }
    $dirty = @()
    $lines = @()
    try { $lines = @(& git -C $RepoRoot status --porcelain --untracked-files=no 2>$null) }
    catch { $lines = @() }
    foreach ($line in $lines) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $dirty += [string]$line }
    }
    return [pscustomobject]@{ Head = $head; Dirty = $dirty; Expected = ($head + ' (clean)') }
}

# ===========================================================================
# CELL READS - ONE COM CALL PER RECTANGLE
# ===========================================================================
function Get-FaBlock {
    param($Workbook, [string]$SheetName, [string]$Address)
    $worksheets = $null; $ws = $null; $rng = $null
    try {
        $worksheets = $Workbook.Worksheets
        $ws = $worksheets.Item($SheetName)
        $rng = $ws.Range($Address)
        return [pscustomobject]@{ Sheet = $SheetName; Address = $Address; Rect = $rng.Value2 }
    } finally {
        if ($null -ne $rng)        { Release-Transient $rng        'Range(block)'; $rng        = $null }
        if ($null -ne $ws)         { Release-Transient $ws         'Worksheet';    $ws         = $null }
        if ($null -ne $worksheets) { Release-Transient $worksheets 'Worksheets';   $worksheets = $null }
    }
}

function Get-FaBlockCell {
    param($Block, [int]$Row, [int]$Column)
    $where = 'block ' + [string]$Block.Sheet + '!' + [string]$Block.Address +
             ' at (' + [string]$Row + ',' + [string]$Column + ')'
    $values = $Block.Rect
    if ($values -is [array]) {
        if ($values.Rank -ne 2) {
            throw ($where + ': the block arrived with rank ' + [string]$values.Rank +
                   ' instead of 2, so the rectangle has been flattened')
        }
        $cell = $values.GetValue($Row, $Column)
    } else {
        if (($Row -ne 1) -or ($Column -ne 1)) {
            throw ($where + ': the block is a single cell and only (1,1) exists in it')
        }
        $cell = $values
    }
    if ($cell -is [array]) {
        throw ($where + ': the read produced ' + [string]@($cell).Count +
               ' values where exactly one was expected')
    }
    return $cell
}

function Format-FaCell {
    param($Value)
    if ($null -eq $Value) { return '<blank>' }
    if (($Value -is [int]) -and $script:FaErrorCodes.ContainsKey([int]$Value)) {
        return [string]$script:FaErrorCodes[[int]$Value]
    }
    if ($Value -is [string]) { return $Value }
    return [string]$Value
}

# A rectangle rendered as one string, so two reads of the same rectangle
# compare with -ceq and a difference is one line.
function Get-FaBlockDigest {
    param($Workbook, [string]$SheetName, [string]$Address)
    $block = Get-FaBlock -Workbook $Workbook -SheetName $SheetName -Address $Address
    $values = $block.Rect
    $parts = @()
    if ($values -is [array]) {
        $rows = $values.GetLength(0); $cols = $values.GetLength(1)
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $parts += (Format-FaCell ($values.GetValue($r, $c)))
            }
        }
    } else {
        $parts += (Format-FaCell $values)
    }
    return ($SheetName + '!' + $Address + '=' + ($parts -join [char]31))
}

function Test-FaBlockBlank {
    param($Workbook, [string]$SheetName, [string]$Address)
    $block = Get-FaBlock -Workbook $Workbook -SheetName $SheetName -Address $Address
    $values = $block.Rect
    if ($values -is [array]) {
        $rows = $values.GetLength(0); $cols = $values.GetLength(1)
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $v = $values.GetValue($r, $c)
                if (($null -ne $v) -and ([string]$v -ne '')) { return $false }
            }
        }
        return $true
    }
    return (($null -eq $values) -or ([string]$values -eq ''))
}

# The cells the reset projection lists one by one, digested one by one: the
# projection owns the addresses, this runner owns nothing about them.
function Get-FaCellListDigest {
    param($Workbook, [string]$SheetName, [string[]]$Cells)
    $parts = @()
    foreach ($address in $Cells) {
        $block = Get-FaBlock -Workbook $Workbook -SheetName $SheetName -Address ([string]$address)
        $parts += ([string]$address + '=' + (Format-FaCell $block.Rect))
    }
    return ($SheetName + ':' + ($parts -join [char]31))
}

function Get-FaTableDigest {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $rows = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName)) {
        $rows += (@($row) -join [char]31)
    }
    return ($TableName + '=' + ($rows -join [char]30))
}

# THE DIAGNOSIS OF A GRID THAT DID NOT COME BACK TO ITS BASELINE. Pure data in,
# one string out: no workbook, no COM, no write. Final acceptance run 11 at
# 34fcb69 reached 103 checks and failed exactly one, repair.grids-restored,
# whose detail said only that the digests differed. The digest is the body's
# Value2 content, so the mismatch is content - a value, a row or a column - and
# this says WHICH, and at WHICH of the two stages the restoration passes
# through: the grid as it stood after the runner's own restoration writes and
# BEFORE the final PCCM_ApplyTimeline resync, and the grid AFTER it. The pass
# predicate of the check is untouched; this only explains a failure.
function ConvertTo-FaBodyKey {
    param([object[]]$Body)
    $rows = @()
    foreach ($row in @($Body)) { $rows += (@($row) -join [char]31) }
    return ($rows -join [char]30)
}

function Get-FaBodyCell {
    param([object[]]$Body, [int]$Row, [int]$Column)
    if (($Row -lt 1) -or ($Row -gt @($Body).Count)) { return '' }
    $line = @(@($Body)[$Row - 1])
    if (($Column -lt 1) -or ($Column -gt $line.Count)) { return '' }
    return [string]$line[$Column - 1]
}

function Get-FaBodyWidth {
    param([object[]]$Body)
    $width = 0
    foreach ($row in @($Body)) { if (@($row).Count -gt $width) { $width = @($row).Count } }
    return $width
}

function Format-FaGridDifferences {
    param([string]$Label, [object[]]$Baseline, [object[]]$BeforeResync, [object[]]$AfterResync,
          [string[]]$BaselineColumns, [string[]]$BeforeColumns, [string[]]$AfterColumns, [int]$Limit = 5)
    $baseKey = ConvertTo-FaBodyKey -Body $Baseline
    $preKey = ConvertTo-FaBodyKey -Body $BeforeResync
    $postKey = ConvertTo-FaBodyKey -Body $AfterResync
    $preDiffers = ($preKey -cne $baseKey)
    $postDiffers = ($postKey -cne $baseKey)
    $resyncChanged = ($postKey -cne $preKey)
    if ((-not $preDiffers) -and (-not $postDiffers)) { return '' }
    $stage = ''
    if ($preDiffers -and (-not $resyncChanged)) {
        $stage = 'CASE 1: the mismatch already existed before the final PCCM_ApplyTimeline, which changed nothing - the restoration choreography did not return the grid to its baseline'
    } elseif ((-not $preDiffers) -and $postDiffers) {
        $stage = 'CASE 2: the final PCCM_ApplyTimeline introduced the mismatch on a grid that was baseline-identical before it'
    } elseif ($preDiffers -and $postDiffers) {
        $stage = 'MIXED: the grid was not at its baseline before the final PCCM_ApplyTimeline, and ApplyTimeline changed it further'
    } else {
        $stage = 'NOTE: the grid was not at its baseline before the final PCCM_ApplyTimeline, and ApplyTimeline returned it to the baseline'
    }
    $rowMax = [math]::Max([math]::Max(@($Baseline).Count, @($BeforeResync).Count), @($AfterResync).Count)
    $colMax = [math]::Max([math]::Max((Get-FaBodyWidth -Body $Baseline), (Get-FaBodyWidth -Body $BeforeResync)), (Get-FaBodyWidth -Body $AfterResync))
    $differences = @()
    $extra = 0
    for ($r = 1; $r -le $rowMax; $r++) {
        for ($c = 1; $c -le $colMax; $c++) {
            $b = Get-FaBodyCell -Body $Baseline -Row $r -Column $c
            $q = Get-FaBodyCell -Body $BeforeResync -Row $r -Column $c
            $a = Get-FaBodyCell -Body $AfterResync -Row $r -Column $c
            if (($b -ceq $a) -and ($b -ceq $q)) { continue }
            if ($differences.Count -ge $Limit) { $extra = $extra + 1; continue }
            $key = Get-FaBodyCell -Body $Baseline -Row $r -Column 1
            if ($key -eq '') { $key = Get-FaBodyCell -Body $BeforeResync -Row $r -Column 1 }
            if ($key -eq '') { $key = Get-FaBodyCell -Body $AfterResync -Row $r -Column 1 }
            if ($key -eq '') { $key = '<no key>' }
            $header = '<no header>'
            foreach ($names in @($BaselineColumns, $BeforeColumns, $AfterColumns)) {
                if (($header -eq '<no header>') -and ($c -le @($names).Count)) { $header = [string]@($names)[$c - 1] }
            }
            $differences += ('row ' + [string]$r + ' (' + $key + ') column ' + [string]$c + ' [' + $header + ']: baseline <' + $b +
                             '>, pre-resync <' + $q + '>, post-resync <' + $a + '>')
        }
    }
    $text = ($Label + ': ' + $stage + '; rows baseline/pre-resync/post-resync=' + [string]@($Baseline).Count + '/' +
             [string]@($BeforeResync).Count + '/' + [string]@($AfterResync).Count + '; columns baseline/pre-resync/post-resync=' +
             [string]@($BaselineColumns).Count + '/' + [string]@($BeforeColumns).Count + '/' + [string]@($AfterColumns).Count +
             '; ' + ($differences -join '; '))
    if ($extra -gt 0) { $text = $text + '; ' + [string]$extra + ' more difference(s)' }
    return $text
}

# THE CAPACITY PLAN FOR A GRID THE RUNNER'S OWN FIXTURE SHORTENED. Pure: two
# counts in, a plan out. Final acceptance run 12 at 01b9a1a diagnosed
# repair.grids-restored as CASE 1 with rows 23/22/22 and no value or column
# difference: the runner had deleted a physical profiling ListRow to make the
# rollback precondition, production recreated the driver in an existing blank
# row - SyncRows promises one row per identified id in register order and a
# cleared tail, never a prior physical count - and the runner then demanded
# byte-identical equality without undoing its own deletion. So the runner
# appends exactly the blank rows it owes, and never deletes: an excess is not
# the runner's to explain away.
function Get-FaCapacityPlan {
    param([int]$Current, [int]$Baseline)
    if ($Current -gt $Baseline) {
        return [pscustomobject]@{ Append = 0; Excess = ($Current - $Baseline) }
    }
    return [pscustomobject]@{ Append = ($Baseline - $Current); Excess = 0 }
}

function Get-FaRegister {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if ([string]$register.key -ceq $Key) { return $register }
    }
    throw ('the manifest declares no register ' + $Key)
}

function Get-FaGrid {
    param($Manifest, [string]$Key)
    foreach ($grid in @($Manifest.grids)) {
        if ([string]$grid.key -ceq $Key) { return $grid }
    }
    throw ('the manifest declares no grid ' + $Key)
}

function Get-FaColumnOrdinal {
    param($Register, [string]$ColumnKey)
    $columns = @($Register.columns)
    for ($index = 0; $index -lt $columns.Count; $index++) {
        if ([string]$columns[$index] -ceq $ColumnKey) { return ($index + 1) }
    }
    throw ('the register ' + [string]$Register.key + ' declares no column ' + $ColumnKey)
}

function Find-FaTableRow {
    param([object[]]$Body, [string]$Id)
    for ($index = 0; $index -lt $Body.Count; $index++) {
        if ([string]$Body[$index][0] -ceq $Id) { return ($index + 1) }
    }
    return 0
}

# THE PRECONDITION THE PRODUCTION SEMANTIC GATE ASSESSES, read from a grid body:
# every row with an id is either blank across the first $YearCount project years or
# totals 100% there, within the tolerance the Repair suite uses. A row that is
# populated and does not is named, so a fixture can be proved before Repair runs.
function Get-FaProfileProblems {
    param([object[]]$Body, [int]$FixedColumns, [int]$YearCount, [string]$Label)
    $problems = @()
    foreach ($entry in $Body) {
        $row = @($entry)
        if ([string]$row[0] -eq '') { continue }
        $total = 0.0
        $populated = $false
        for ($y = 1; $y -le $YearCount; $y++) {
            $cell = [string]$row[$FixedColumns + $y - 1]
            if ($cell -eq '') { continue }
            $populated = $true
            $total = $total + [double]$cell
        }
        if ($populated -and ([math]::Abs($total - 1.0) -gt 0.000000001)) {
            $problems += ($Label + ' ' + [string]$row[0] + ' is populated and totals ' + [string]$total +
                          ' over project years 1-' + [string]$YearCount + ', not 100%')
        }
    }
    return $problems
}

# THE LOCK STATE OF ONE TABLE CELL, READ, NEVER WRITTEN: its worksheet address
# and its Locked property, so a runtime-generated cell can be compared with the
# protection projection's declared unlocked set and with the property itself
# rather than inferred from whether a write happened to succeed.
function Get-FaCellLockState {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        return [pscustomobject]@{ Address = [string]$cell.Address($false, $false); Locked = [bool]$cell.Locked }
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

# THE OPEN/READY BOUNDARY. Final acceptance runs 15 and 16 at b5c3f9a both died
# between Workbooks.Open and the first workbook read - once with
# RPC_E_CALL_REJECTED, once with a null-valued expression - with the Stage-B
# bootstrap already complete and the shutdown clean: the classic post-open race
# the Stage-B bootstrap already guards with Wait-StageBWorkbookReady. This is the
# same policy, through the same shared authority: every read goes through
# Invoke-ComRetryRead (com_lifecycle.ps1), which retries ONLY the recognised
# transient rejections (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER) within
# its accepted bounds and rethrows everything else untouched; a rejection the
# helper could not resolve within its own bounds is 'not ready yet' for one more
# bounded outer attempt. Two harmless reads decide readiness - the workbook's
# FullName (identity: the wrong path is an error, never a delay) and its
# Worksheets.Count (the first structural fact every later read needs). No
# production endpoint, no acceptance predicate and no write sit inside the loop.
function Wait-FaWorkbookReady {
    param($Workbook, [string]$ExpectedPath,
          [int]$MaxAttempts   = 12,
          [int]$FirstDelayMs  = 250,
          [int]$MaxDelayMs    = 2000,
          [int]$TotalBudgetMs = 15000)
    if ($null -eq $Workbook) { throw 'Wait-FaWorkbookReady: no workbook.' }
    if ([string]::IsNullOrWhiteSpace($ExpectedPath)) { throw 'Wait-FaWorkbookReady: no expected path to recognise the workbook by.' }
    if ($MaxAttempts -lt 1) { throw 'Wait-FaWorkbookReady: MaxAttempts must be at least 1.' }
    if ($TotalBudgetMs -lt 0) { throw 'Wait-FaWorkbookReady: TotalBudgetMs may not be negative.' }
    $attempt = 0
    $waitedMs = 0
    $delay = $FirstDelayMs
    $nameState = 'unresolved'
    $sheetState = 'unresolved'
    $fullName = ''
    $sheetCount = 0
    $rejections = @()
    $ready = $false
    while ($attempt -lt $MaxAttempts) {
        $attempt = $attempt + 1
        $nameState = 'unresolved'
        $sheetState = 'unresolved'
        $nameValue = $null
        try {
            $nameRead = Invoke-ComRetryRead -Target $Workbook -Member 'FullName' -Description 'the opened workbook FullName'
            if ($nameRead.Attempts -gt 1) { $rejections += ('FullName:' + $nameRead.Rejections) }
            $nameValue = $nameRead.Value
        } catch {
            if ([string]::IsNullOrWhiteSpace((Get-ComRejectionName $_))) { throw }
            $rejections += ('FullName:' + (Get-ComRejectionName $_))
            $nameValue = $null
        }
        if (($null -eq $nameValue) -or [string]::IsNullOrWhiteSpace([string]$nameValue)) {
            $nameState = 'no-answer'
        } elseif (([string]$nameValue).Trim().ToLowerInvariant() -ne $ExpectedPath.Trim().ToLowerInvariant()) {
            throw ('READY: the opened workbook is bound to ' + [string]$nameValue + ' and not to ' + $ExpectedPath + '. Waiting cannot change which workbook this is.')
        } else {
            $fullName = [string]$nameValue
            $nameState = 'ok'
        }
        $sheetsObject = $null
        $countValue = $null
        try {
            $sheetsRead = Invoke-ComRetryRead -Target $Workbook -Member 'Worksheets' -Description 'the opened workbook Worksheets'
            if ($sheetsRead.Attempts -gt 1) { $rejections += ('Worksheets:' + $sheetsRead.Rejections) }
            $sheetsObject = $sheetsRead.Value
            if ($null -ne $sheetsObject) {
                $countRead = Invoke-ComRetryRead -Target $sheetsObject -Member 'Count' -Description 'the opened workbook Worksheets.Count'
                if ($countRead.Attempts -gt 1) { $rejections += ('Worksheets.Count:' + $countRead.Rejections) }
                $countValue = $countRead.Value
            }
        } catch {
            if ([string]::IsNullOrWhiteSpace((Get-ComRejectionName $_))) { throw }
            $rejections += ('Worksheets:' + (Get-ComRejectionName $_))
            $countValue = $null
        } finally {
            if ($null -ne $sheetsObject) { Release-Transient $sheetsObject 'Worksheets'; $sheetsObject = $null }
        }
        if ($null -eq $countValue) {
            $sheetState = 'no-answer'
        } elseif ((([string]$countValue).Trim()) -notmatch '^[0-9]+$') {
            throw ('READY: the opened workbook Worksheets.Count answered ' + ([string]$countValue).Trim() + ', which is not a count.')
        } elseif ([int](([string]$countValue).Trim()) -lt 1) {
            $sheetState = 'no-answer'
        } else {
            $sheetCount = [int](([string]$countValue).Trim())
            $sheetState = 'ok'
        }
        if (($nameState -eq 'ok') -and ($sheetState -eq 'ok')) { $ready = $true; break }
        if ($attempt -ge $MaxAttempts) { break }
        if (($waitedMs + $delay) -gt $TotalBudgetMs) { break }
        Start-Sleep -Milliseconds $delay
        $waitedMs = $waitedMs + $delay
        $delay = [Math]::Min(($delay + $FirstDelayMs), $MaxDelayMs)
    }
    return [pscustomobject]@{
        Ready      = $ready
        Attempts   = $attempt
        WaitedMs   = $waitedMs
        FullName   = $fullName
        Sheets     = $sheetCount
        NameState  = $nameState
        SheetState = $sheetState
        Rejections = (@($rejections | Select-Object -Unique) -join ', ')
    }
}

# THE OPEN/READY LINE AND CHECK. One line whatever happened - attempts, the
# milliseconds waited and every rejection name - and a fail-fast check that
# names the session, so a workbook that never becomes ready is an OPEN/READY
# harness failure before any acceptance scenario is claimed.
function Assert-FaWorkbookReady {
    param($Workbook, [string]$ExpectedPath, [string]$Session)
    $state = Wait-FaWorkbookReady -Workbook $Workbook -ExpectedPath $ExpectedPath
    Write-FaLine ('READY|' + $Session + '|attempts=' + [string]$state.Attempts + '|waited=' + [string]$state.WaitedMs + 'ms' +
                  '|fullname=' + $state.NameState + '|worksheets=' + $state.SheetState +
                  '|rejections=' + $(if ($state.Rejections -eq '') { 'none' } else { $state.Rejections }))
    $null = Add-FaCheck ('session.open-ready.' + $Session) $state.Ready `
        $(if ($state.Ready) { ($state.FullName + ' answered ' + [string]$state.Sheets + ' worksheets after ' + [string]$state.Attempts + ' attempt(s) and ' + [string]$state.WaitedMs + ' ms') }
          else { ('OPEN/READY harness failure: the workbook did not answer after ' + [string]$state.Attempts + ' attempt(s) and ' + [string]$state.WaitedMs + ' ms (fullname=' + $state.NameState + ', worksheets=' + $state.SheetState + '; rejections: ' + $state.Rejections + ')') })
    return $state
}

# THE FATAL DIAGNOSTICS. Plain data from the error record: the exception type,
# the message, the script line, the offending source line and a bounded script
# stack, so a null-valued expression is located in one run. The verdict is not
# touched: a fatal is still a fatal.
function Format-FaFatalLines {
    param($ErrorRecord, [int]$MaxStackLines = 8)
    $lines = @()
    $type = ''
    $message = ''
    try { $type = [string]$ErrorRecord.Exception.GetType().FullName } catch { $type = '' }
    try { $message = [string]$ErrorRecord.Exception.Message } catch { $message = '' }
    if ($message -eq '') { try { $message = [string]$ErrorRecord } catch { $message = 'unknown error' } }
    $lines += ('FATAL.TYPE|' + $(if ($type -eq '') { 'unknown' } else { $type }))
    $lines += ('FATAL.MESSAGE|' + ($message -replace '[\r\n]+', ' '))
    $inner = ''
    try { if ($null -ne $ErrorRecord.Exception.InnerException) { $inner = [string]$ErrorRecord.Exception.InnerException.Message } } catch { $inner = '' }
    if ($inner -ne '') { $lines += ('FATAL.INNER|' + ($inner -replace '[\r\n]+', ' ')) }
    $lineNumber = ''
    $sourceLine = ''
    $position = ''
    try {
        if ($null -ne $ErrorRecord.InvocationInfo) {
            $lineNumber = [string]$ErrorRecord.InvocationInfo.ScriptLineNumber
            $sourceLine = ([string]$ErrorRecord.InvocationInfo.Line).Trim()
            $position = ([string]$ErrorRecord.InvocationInfo.PositionMessage -replace '[\r\n]+', ' ').Trim()
        }
    } catch { }
    $lines += ('FATAL.LINE|' + $(if ($lineNumber -eq '') { 'unknown' } else { $lineNumber }))
    if ($sourceLine -ne '') { $lines += ('FATAL.SOURCE|' + $sourceLine) }
    if ($position -ne '') { $lines += ('FATAL.POSITION|' + $position) }
    $stack = @()
    try { $stack = @(([string]$ErrorRecord.ScriptStackTrace) -split '\r?\n' | Where-Object { $_.Trim() -ne '' }) } catch { $stack = @() }
    $shown = 0
    foreach ($frame in $stack) {
        if ($shown -ge $MaxStackLines) { $lines += ('FATAL.STACK|... ' + [string]($stack.Count - $MaxStackLines) + ' more frame(s)'); break }
        $lines += ('FATAL.STACK|' + $frame.Trim())
        $shown = $shown + 1
    }
    return $lines
}

# ===========================================================================
# THE PRODUCTION ENTRY POINTS, THROUGH THE ACCEPTED AUTOMATION SEAM
# ===========================================================================
# Invoke-Phase5ProductionOperation (accepted) THROWS on a FAIL announcement. The
# scenarios that EXPECT a refusal need the announcement itself, so this returns
# it, exactly as Invoke-Phase6Simulation does, and clears the failpoint after.
# THE ENDPOINT, OBSERVED. Final acceptance run 13 at 1cbcf5e failed
# reset.declined although production had asked the destructive confirmation:
# the prompt was read AFTER this helper's finally had re-begun the seam, and
# PCCM_AutomationBegin clears the recorded prompt and result. So the evidence is
# read inside the try - the result, then the prompt - and returned as plain
# data; only then does the finally reset the seam, so no failpoint and no reply
# ever outlive the call. Invoke-FaEndpoint is the same call for the callers
# that need the result alone.
function Invoke-FaObservedEndpoint {
    param($Excel, [string]$Operation, [bool]$ConfirmReply = $true,
          [string]$FailAfterStage = '')
    $Excel.Run('PCCM_AutomationBegin', $ConfirmReply, $FailAfterStage) | Out-Null
    try {
        $Excel.Run($Operation) | Out-Null
        $result = [string]$Excel.Run('PCCM_AutomationResult')
        $prompt = [string]$Excel.Run('PCCM_AutomationPrompt')
        return [pscustomobject]@{ Result = $result; Prompt = $prompt }
    } finally {
        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    }
}

function Invoke-FaEndpoint {
    param($Excel, [string]$Operation, [bool]$ConfirmReply = $true,
          [string]$FailAfterStage = '')
    $observed = Invoke-FaObservedEndpoint -Excel $Excel -Operation $Operation -ConfirmReply $ConfirmReply -FailAfterStage $FailAfterStage
    return [string]$observed.Result
}

function Get-FaRunText {
    param($Excel, [string]$Procedure)
    $value = $Excel.Run($Procedure)
    if ($null -eq $value) { return '' }
    return [string]$value
}

# ===========================================================================
# THE PROTECTION WINDOW - THE BENCHMARK'S ACCEPTED SHIM, COPIED
# ===========================================================================
function Import-FaFixtureWindow {
    param($Excel, $Workbook, $Manifest, [string]$ScriptDir)
    $source = Join-Path $ScriptDir $script:FixtureWindowSource
    if (-not (Test-Path -LiteralPath $source)) {
        throw ('the fixture-window shim ' + $source + ' is missing, so the runner ' +
               'cannot open the accepted protection window')
    }
    $declared = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
    if ($declared -contains $script:FixtureWindowModule) {
        throw ('the manifest declares ' + $script:FixtureWindowModule + ' as a production ' +
               'module; the runner will not import a test module over production')
    }
    $project = $null; $components = $null; $imported = $null
    try {
        $project = $Workbook.VBProject
        $components = $project.VBComponents
        $imported = $components.Import($source)
        $name = [string]$imported.Name
        if ($name -ne $script:FixtureWindowModule) {
            throw ('the fixture-window shim imported as ' + $name + ', not ' +
                   $script:FixtureWindowModule)
        }
    } finally {
        if ($null -ne $imported)   { Release-Transient $imported   'VBComponent(shim)'; $imported   = $null }
        if ($null -ne $components) { Release-Transient $components 'VBComponents';      $components = $null }
        if ($null -ne $project)    { Release-Transient $project    'VBProject';         $project    = $null }
    }
    $ping = [string]$Excel.Run('P10FW_Ping')
    if ($ping -ne ('OK|' + $script:FixtureWindowModule)) {
        throw ('the fixture-window shim imported but does not answer: ' + $ping)
    }
}

function Get-FaProtectionState {
    param($Excel)
    $raw = [string]$Excel.Run('P10FW_State')
    if ($raw -notlike 'OK|*') {
        throw ('the workbook could not report its protection state: ' + $raw)
    }
    $state = @{}
    foreach ($field in $raw.Substring(3).Split([char]124)) {
        $pair = $field.Split([char]61)
        if ($pair.Count -eq 2) { $state[$pair[0]] = $pair[1] }
    }
    foreach ($required in @('applied', 'depth', 'structure', 'sheets', 'protected')) {
        if (-not $state.ContainsKey($required)) {
            throw ('the protection state is missing ' + $required + ': ' + $raw)
        }
    }
    return [pscustomobject]@{
        Applied   = ([string]$state['applied'] -eq 'True')
        Depth     = [int]$state['depth']
        Structure = ([string]$state['structure'] -eq 'True')
        Sheets    = [int]$state['sheets']
        Protected = [int]$state['protected']
        Raw       = $raw
    }
}

# THE ONE PROTECTION ASSERTION, used after open, around the window, and after
# every command, refusal and injected failure. The expected sheet count is the
# protection projection's declared set, never a literal.
function Assert-FaProtectionApplied {
    param($Excel, $Protection, [string]$Scenario)
    $expected = @($Protection.sheets).Count
    $state = Get-FaProtectionState -Excel $Excel
    $problems = @()
    if (-not $state.Applied)            { $problems += 'the workbook does not report itself protected' }
    if (-not $state.Structure)          { $problems += 'workbook structure protection is not applied' }
    if ($state.Sheets -ne $expected)    { $problems += ('the workbook holds ' + [string]$state.Sheets +
                                                        ' worksheets where the projection declares ' +
                                                        [string]$expected) }
    if ($state.Protected -ne $expected) { $problems += ([string]$state.Protected + ' of ' +
                                                        [string]$expected + ' worksheets are protected') }
    if ($state.Depth -ne 0)             { $problems += ('the structural window is still open at depth ' +
                                                        [string]$state.Depth) }
    $null = Add-FaCheck $Scenario ($problems.Count -eq 0) `
        ($(if ($problems.Count -eq 0) {
            ('structure=' + [string]$state.Structure + ' protected=' + [string]$state.Protected +
             '/' + [string]$state.Sheets + ' depth=' + [string]$state.Depth)
          } else { ($problems -join '; ') + '. Reported: ' + $state.Raw }))
    return $state
}

function Open-FaFixtureWindow {
    param($Excel, $Protection, [string]$Scenario)
    $null = Assert-FaProtectionApplied -Excel $Excel -Protection $Protection `
        -Scenario ($Scenario + '.window-before')
    $reply = [string]$Excel.Run('P10FW_Begin')
    if ($reply -notlike 'OK|*') {
        throw ('the setup window could not be opened: ' + $reply)
    }
    try {
        $state = Get-FaProtectionState -Excel $Excel
        if ($state.Depth -ne 1) {
            throw ('the setup window opened to depth ' + [string]$state.Depth +
                   ', not 1. The runner opens exactly one outer window and closes it.')
        }
        if (-not $state.Structure) {
            throw ('opening the setup window released WORKBOOK STRUCTURE protection, ' +
                   'which it must never do. Reported: ' + $state.Raw)
        }
        return $state
    } catch {
        $original = [string]$_.Exception.Message
        $recovery = ''
        try {
            $rollback = [string]$Excel.Run('P10FW_End')
            $recovery = 'a compensating close reported ' + $rollback
        } catch {
            $recovery = 'a compensating close RAISED: ' + [string]$_.Exception.Message
        }
        throw ($original + ' The window was open when this failed; ' + $recovery)
    }
}

function Close-FaFixtureWindow {
    param($Excel, $Protection, [string]$Scenario)
    $reply = [string]$Excel.Run('P10FW_End')
    if ($reply -notlike 'OK|*') {
        throw ('the setup window could not be closed and protection was not restored: ' +
               $reply + '. No acceptance may be read from this workbook.')
    }
    return (Assert-FaProtectionApplied -Excel $Excel -Protection $Protection `
        -Scenario ($Scenario + '.window-after'))
}

# ===========================================================================
# THE RESET PROJECTION, READ AS DIGESTS
# ===========================================================================
# Everything Reset PRESERVES, digested: every declared editable input cell, the
# applied timeline cells, the permanent-id counters, the identity cells the
# simulation owner keeps, and the build metadata block.
function Get-FaPreservedDigest {
    param($Workbook, $Reset, $Methodology, [string]$MetadataRange)
    $parts = @()
    foreach ($entry in @($Reset.preserved.editable_inputs)) {
        $parts += (Get-FaCellListDigest -Workbook $Workbook -SheetName ([string]$entry.sheet) `
            -Cells @($entry.cells | ForEach-Object { [string]$_ }))
    }
    $timeline = $Reset.preserved.applied_timeline
    $parts += (Get-FaCellListDigest -Workbook $Workbook -SheetName ([string]$timeline.sheet) `
        -Cells @($timeline.cells | ForEach-Object { [string]$_ }))
    $counters = $Reset.preserved.permanent_id_counters
    $parts += (Get-FaCellListDigest -Workbook $Workbook -SheetName ([string]$counters.sheet) `
        -Cells @($counters.cells | ForEach-Object { [string]$_ }))
    $identity = $Reset.publications.simulation.preserved
    $parts += (Get-FaCellListDigest -Workbook $Workbook `
        -SheetName ([string]$Reset.publications.simulation.sheet) `
        -Cells @([string]$identity.next_auto_nonce, [string]$identity.last_run_id,
                 [string]$identity.pending_auto_nonce))
    $parts += (Get-FaBlockDigest -Workbook $Workbook -SheetName ([string]$Methodology.sheet) `
        -Address $MetadataRange)
    return ($parts -join [char]29)
}

function Get-FaColumnRange {
    param($Columns, [int]$FirstRow, [int]$Rows)
    $first = [string]$Columns[0]
    $last = [string]$Columns[$Columns.Count - 1]
    return ($first + [string]$FirstRow + ':' + $last + [string]($FirstRow + $Rows - 1))
}

function Get-FaBlockRange {
    param($Block, [int]$Rows)
    return ([string]$Block.first_column + [string]$Block.first_row + ':' +
            [string]$Block.last_column + [string]([int]$Block.first_row + $Rows - 1))
}

# Every rectangle the reset projection says a reset CLEARS, as (sheet, address)
# pairs. The iteration banks and the annual and sensitivity blocks are open-ended
# downwards; the runner bounds them at the request it made, which is where a
# publication of that request can have written.
function Get-FaClearedRectangles {
    param($Reset, [int]$Iterations)
    $out = @()
    $calc = $Reset.publications.calculation
    $out += [pscustomobject]@{ Sheet = [string]$calc.sheet; Address = [string]$calc.cleared.state }
    $out += [pscustomobject]@{ Sheet = [string]$calc.sheet; Address = [string]$calc.cleared.totals }
    $sim = $Reset.publications.simulation
    $sheet = [string]$sim.sheet
    $cleared = $sim.cleared
    foreach ($bank in @('A', 'B')) {
        $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.bank_snapshots.$bank }
        $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.summary.$bank }
        $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.contingency.$bank }
        $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.annual_stamps.$bank }
        $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.sensitivity_stamps.$bank }
        $iterationBank = $cleared.iteration_banks.$bank
        $out += [pscustomobject]@{ Sheet = $sheet
                                   Address = (Get-FaColumnRange -Columns @($iterationBank.columns) `
                                                  -FirstRow ([int]$iterationBank.first_row) -Rows $Iterations) }
        $out += [pscustomobject]@{ Sheet = $sheet
                                   Address = (Get-FaBlockRange -Block $cleared.annual_blocks.$bank -Rows $Iterations) }
        $out += [pscustomobject]@{ Sheet = $sheet
                                   Address = (Get-FaBlockRange -Block $cleared.sensitivity_records.$bank -Rows $Iterations) }
    }
    $out += [pscustomobject]@{ Sheet = $sheet; Address = [string]$cleared.attempt_and_selector }
    return $out
}

function Get-FaPublicationDigest {
    param($Workbook, $Reset, [int]$Iterations)
    $parts = @()
    foreach ($rect in @(Get-FaClearedRectangles -Reset $Reset -Iterations $Iterations)) {
        $parts += (Get-FaBlockDigest -Workbook $Workbook -SheetName $rect.Sheet -Address $rect.Address)
    }
    $calc = $Reset.publications.calculation
    foreach ($table in @($calc.cleared.tables)) {
        $parts += (Get-FaTableDigest -Workbook $Workbook -SheetName ([string]$calc.sheet) `
            -TableName ([string]$table))
    }
    return ($parts -join [char]29)
}

# THE SEMANTIC POST-RESET STATE, AS PRODUCTION DEFINES IT. Final acceptance run
# 14 at 1e62de2 failed reset.confirmed although production had reset exactly as
# its own source says it does: Reset clears CONTENTS and preserves table shape
# (CalcReportClearPublication clears each _Calc table body by address and never
# resizes it), clears the calculation state block and then writes the accepted
# NONE attempt value back into its last-attempt field, and clears the simulation
# publication record and then writes NONE into ITS last-attempt field. The old
# verifier demanded zero ListRows and blank sentinel cells - conditions no
# correct reset can meet. These helpers are pure - plain data in, problem
# strings out - and the COM reader beneath them assembles the data. Every
# ordinary publication rectangle is still required strictly blank.
function ConvertTo-FaRectCells {
    param($Rect)
    $cells = @()
    if ($Rect -is [array]) {
        $rows = $Rect.GetLength(0); $cols = $Rect.GetLength(1)
        for ($r = 1; $r -le $rows; $r++) {
            for ($c = 1; $c -le $cols; $c++) {
                $v = $Rect.GetValue($r, $c)
                if ($null -eq $v) { $cells += '' } else { $cells += [string]$v }
            }
        }
    } elseif ($null -eq $Rect) { $cells += '' } else { $cells += [string]$Rect }
    return , $cells
}

function Get-FaSemanticBlockProblems {
    param([string]$Label, [string[]]$Cells, [int]$SentinelIndex, [string]$Sentinel)
    $problems = @()
    if (($SentinelIndex -lt 1) -or ($SentinelIndex -gt @($Cells).Count)) {
        $problems += ($Label + ' sentinel index ' + [string]$SentinelIndex + ' is outside its ' + [string]@($Cells).Count + ' cell(s)')
        return , $problems
    }
    for ($i = 1; $i -le @($Cells).Count; $i++) {
        $value = [string]@($Cells)[$i - 1]
        if ($i -eq $SentinelIndex) {
            if ($value -cne $Sentinel) { $problems += ($Label + ' field ' + [string]$i + ' holds <' + $value + '>, expected the ' + $Sentinel + ' sentinel') }
        } elseif ($value -ne '') {
            $problems += ($Label + ' field ' + [string]$i + ' holds <' + $value + '>, expected blank')
        }
    }
    return , $problems
}

function Get-FaTableBodyProblems {
    param([string]$Label, [object[]]$Body, [int]$ExpectedRows, [int]$ExpectedColumns)
    $problems = @()
    if (@($Body).Count -ne $ExpectedRows) { $problems += ($Label + ' holds ' + [string]@($Body).Count + ' row(s) where the precondition held ' + [string]$ExpectedRows + '; Reset owns contents, not geometry') }
    $found = 0
    for ($r = 1; $r -le @($Body).Count; $r++) {
        $line = @(@($Body)[$r - 1])
        if (($r -eq 1) -and ($line.Count -ne $ExpectedColumns)) { $problems += ($Label + ' holds ' + [string]$line.Count + ' column(s) where the precondition held ' + [string]$ExpectedColumns) }
        for ($c = 1; $c -le $line.Count; $c++) {
            if ([string]$line[$c - 1] -eq '') { continue }
            $found = $found + 1
            if ($found -le 5) { $problems += ($Label + ' row ' + [string]$r + ' column ' + [string]$c + ' holds <' + [string]$line[$c - 1] + '>, expected blank') }
        }
    }
    if ($found -gt 5) { $problems += ($Label + ': ' + [string]($found - 5) + ' more populated cell(s)') }
    return , $problems
}

function Get-FaCalcTableShapes {
    param($Workbook, $Reset)
    $calc = $Reset.publications.calculation
    $shapes = @{}
    foreach ($table in @($calc.cleared.tables)) {
        $shapes[[string]$table] = [pscustomobject]@{
            Rows    = (Get-TableRowCount -Workbook $Workbook -SheetName ([string]$calc.sheet) -TableName ([string]$table))
            Columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName ([string]$calc.sheet) -TableName ([string]$table)).Count
        }
    }
    return $shapes
}

function Get-FaResetProblems {
    param($Workbook, $Reset, [int]$Iterations, [string]$CalcAttemptCell, [string]$SimAttemptCell, [string]$Sentinel, $TableShapes)
    $problems = @()
    $calc = $Reset.publications.calculation
    $sim = $Reset.publications.simulation
    $stateAddress = [string]$calc.cleared.state
    $recordAddress = [string]$sim.cleared.attempt_and_selector
    foreach ($rect in @(Get-FaClearedRectangles -Reset $Reset -Iterations $Iterations)) {
        if (($rect.Address -eq $stateAddress) -or ($rect.Address -eq $recordAddress)) { continue }
        if (-not (Test-FaBlockBlank -Workbook $Workbook -SheetName $rect.Sheet -Address $rect.Address)) {
            $problems += ('publication rectangle ' + $rect.Sheet + '!' + $rect.Address + ' is not blank')
        }
    }
    $stateCells = ConvertTo-FaRectCells -Rect ((Get-FaBlock -Workbook $Workbook -SheetName ([string]$calc.sheet) -Address $stateAddress).Rect)
    $problems += @(Get-FaSemanticBlockProblems -Label ('calculation state ' + [string]$calc.sheet + '!' + $stateAddress) -Cells $stateCells `
        -SentinelIndex (Get-FaCellOffset -Address $stateAddress -Cell $CalcAttemptCell) -Sentinel $Sentinel)
    $recordCells = ConvertTo-FaRectCells -Rect ((Get-FaBlock -Workbook $Workbook -SheetName ([string]$sim.sheet) -Address $recordAddress).Rect)
    $problems += @(Get-FaSemanticBlockProblems -Label ('simulation record ' + [string]$sim.sheet + '!' + $recordAddress) -Cells $recordCells `
        -SentinelIndex (Get-FaCellOffset -Address $recordAddress -Cell $SimAttemptCell) -Sentinel $Sentinel)
    foreach ($table in @($calc.cleared.tables)) {
        $shape = $TableShapes[[string]$table]
        $body = @(Get-TableBody -Workbook $Workbook -SheetName ([string]$calc.sheet) -TableName ([string]$table))
        $problems += @(Get-FaTableBodyProblems -Label ([string]$table) -Body $body -ExpectedRows ([int]$shape.Rows) -ExpectedColumns ([int]$shape.Columns))
    }
    return , $problems
}

# A cell's 1-based position inside a single-column block address: row minus the block's first row plus one.
function Get-FaCellOffset {
    param([string]$Address, [string]$Cell)
    $firstRow = [int]($Address.Split(':')[0] -replace '[A-Za-z]', '')
    $column = ($Address.Split(':')[0] -replace '[0-9]', '')
    $cellRow = [int]($Cell -replace '[A-Za-z]', '')
    $cellColumn = ($Cell -replace '[0-9]', '')
    if ($cellColumn -ne $column) { return 0 }
    return ($cellRow - $firstRow + 1)
}

# The live, DERIVED states, read through accessors that write nothing: the
# Model Check adapter for the calculation, the read-only simulation derivation,
# and the annual owner's own state.
function Get-FaStates {
    param($Excel)
    return [pscustomobject]@{
        Calculation = (Get-FaRunText -Excel $Excel -Procedure 'PCCM_ModelCheckCalculationState')
        Simulation  = (Get-FaRunText -Excel $Excel -Procedure 'SimReportDerivedStatus')
        Annual      = (Get-FaRunText -Excel $Excel -Procedure 'PCCM_AnnualDistributionState')
        Profile     = (Get-FaRunText -Excel $Excel -Procedure 'PCCM_AnnualProfileState')
    }
}

function Format-FaStates {
    param($States)
    return ('calc=' + $States.Calculation + ' sim=' + $(if ($States.Simulation -eq '') { '<blank>' } else { $States.Simulation }) +
            ' annual=' + $States.Annual + ' profile=' + $States.Profile)
}

# Read the Model Check surface through the Phase-9 projection: the summary
# rows and every shown register row (check id, group, severity, subject,
# message), two rectangles. Unused register slots hold #N/A and are dropped.
function Get-FaModelCheckSurface {
    param($Workbook, $Projection)
    $sheet = [string]$Projection.sheet
    $valueColumn = [string]$Projection.summary.value_column
    $rows = @()
    foreach ($property in $Projection.summary.rows.PSObject.Properties) {
        $rows += [pscustomobject]@{ Key = [string]$property.Name; Row = [int]$property.Value.row }
    }
    $first = ($rows | Measure-Object -Property Row -Minimum).Minimum
    $last = ($rows | Measure-Object -Property Row -Maximum).Maximum
    $block = Get-FaBlock -Workbook $Workbook -SheetName $sheet `
        -Address ($valueColumn + [string]$first + ':' + $valueColumn + [string]$last)
    $summary = @{}
    foreach ($entry in $rows) {
        $summary[$entry.Key] = Get-FaBlockCell -Block $block -Row ($entry.Row - $first + 1) -Column 1
    }
    $columns = @($Projection.register.columns)
    $firstColumn = [string]$columns[0].column
    $lastColumn = [string]$columns[$columns.Count - 1].column
    $firstRow = [int]$Projection.register.first_row
    $lastRow = [int]$Projection.register.last_row
    $register = Get-FaBlock -Workbook $Workbook -SheetName $sheet `
        -Address ($firstColumn + [string]$firstRow + ':' + $lastColumn + [string]$lastRow)
    $shown = @()
    for ($offset = 0; $offset -lt ($lastRow - $firstRow + 1); $offset++) {
        $record = @{}
        for ($index = 0; $index -lt $columns.Count; $index++) {
            $record[[string]$columns[$index].key] = Get-FaBlockCell -Block $register -Row ($offset + 1) -Column ($index + 1)
        }
        $id = $record['check_id']
        if (($null -eq $id) -or (($id -is [int]) -and $script:FaErrorCodes.ContainsKey([int]$id))) { continue }
        $shown += [pscustomobject]$record
    }
    return [pscustomobject]@{ Summary = $summary; Shown = $shown }
}

function Format-FaActionable {
    param($Rows)
    $parts = @()
    foreach ($row in @($Rows)) {
        $parts += ([string]$row.check_id + '(' + [string]$row.severity + ')[' + (Format-FaCell $row.subject) + ']')
    }
    if ($parts.Count -eq 0) { return '<none>' }
    return ($parts -join ', ')
}

# AN EXPECTED ENTRY, VALIDATED BEFORE IT IS READ. The expected entries are
# hashtables with OPTIONAL fields, and under Set-StrictMode reading a key that
# is absent is itself an error - run 4 died on `$entry.Subject` for the advisory,
# which names no Subject. So every field is tested for PRESENCE with ContainsKey
# before it is read, exactly once, here; the matcher below reads only this
# descriptor, whose properties always exist. A malformed definition - both Id
# and AnyOf, neither, no Severity, an empty AnyOf or a blank id - is a RUNNER
# DEFINITION ERROR and throws before any row is compared; it is never tolerated.
function Test-FaExpectedEntry {
    param($Entry, [string]$Scenario)
    $where = 'RUNNER DEFINITION ERROR at ' + $Scenario + ': an expected Model Check entry '
    if ($Entry -isnot [hashtable]) { throw ($where + 'is not a hashtable') }
    $hasId = $Entry.ContainsKey('Id')
    $hasAnyOf = $Entry.ContainsKey('AnyOf')
    if ($hasId -and $hasAnyOf) { throw ($where + 'names both Id and AnyOf; exactly one selector is allowed') }
    if (-not ($hasId -or $hasAnyOf)) { throw ($where + 'names neither Id nor AnyOf; exactly one selector is required') }
    if (-not $Entry.ContainsKey('Severity')) { throw ($where + 'has no Severity') }
    $severity = [string]$Entry['Severity']
    if ($severity -eq '') { throw ($where + 'has a blank Severity') }
    $ids = @()
    if ($hasId) { $ids = @([string]$Entry['Id']) }
    else { $ids = @(@($Entry['AnyOf']) | ForEach-Object { [string]$_ }) }
    if ($ids.Count -lt 1) { throw ($where + 'has an empty AnyOf') }
    foreach ($id in $ids) { if ($id -eq '') { throw ($where + 'names a blank check id') } }
    $hasSubject = $Entry.ContainsKey('Subject')
    $hasMessage = $Entry.ContainsKey('Message')
    $subject = ''
    if ($hasSubject) { $subject = [string]$Entry['Subject'] }
    $message = ''
    if ($hasMessage) { $message = [string]$Entry['Message'] }
    $wanted = $ids[0]
    if ($hasAnyOf) { $wanted = 'one of ' + ($ids -join '/') }
    return [pscustomobject]@{
        Ids = $ids; Wanted = $wanted; Severity = $severity
        HasSubject = $hasSubject; Subject = $subject
        HasMessage = $hasMessage; Message = $message
    }
}

# THE ONE MODEL CHECK ASSERTION. It takes the EXACT expected actionable set and
# derives everything else from it: the overall status from the vocabulary, the
# error and warning counts from the set, and the requirement that no actionable
# row is shown beyond the set - so an unrelated WARNING cannot satisfy a
# checkpoint, and a missing advisory cannot hide behind a matching overall word.
# An expected entry names an Id or an AnyOf list, a Severity, and optionally a
# Subject and a Message that must match the row exactly; each is validated by
# Test-FaExpectedEntry before anything is read from it.
function Assert-FaModelCheck {
    param($Workbook, $Projection, [string]$Scenario, $Expected)
    $states = @($Projection.vocabulary.overall_states | ForEach-Object { [string]$_ })
    $actionable = @($Projection.vocabulary.actionable_severities | ForEach-Object { [string]$_ })
    $entries = @()
    foreach ($raw in @($Expected)) { $entries += (Test-FaExpectedEntry -Entry $raw -Scenario $Scenario) }
    $expectedErrors = 0; $expectedWarnings = 0
    foreach ($entry in $entries) {
        if ($entry.Severity -ceq $actionable[0]) { $expectedErrors = $expectedErrors + 1 }
        if ($entry.Severity -ceq $actionable[1]) { $expectedWarnings = $expectedWarnings + 1 }
    }
    $expectedOverall = $states[0]
    if ($expectedWarnings -gt 0) { $expectedOverall = $states[1] }
    if ($expectedErrors -gt 0) { $expectedOverall = $states[2] }
    $surface = Get-FaModelCheckSurface -Workbook $Workbook -Projection $Projection
    $overall = Format-FaCell $surface.Summary['overall_status']
    $errors = [int]$surface.Summary['error_count']
    $warnings = [int]$surface.Summary['warning_count']
    $rows = @($surface.Shown | Where-Object { $actionable -ccontains [string]$_.severity })
    $problems = @()
    if ($overall -cne $expectedOverall) { $problems += ('overall ' + $overall + ', expected ' + $expectedOverall) }
    if ($errors -ne $expectedErrors) { $problems += ('errors ' + [string]$errors + ', expected ' + [string]$expectedErrors) }
    if ($warnings -ne $expectedWarnings) { $problems += ('warnings ' + [string]$warnings + ', expected ' + [string]$expectedWarnings) }
    $unmatched = @($rows)
    foreach ($entry in $entries) {
        $found = $null
        foreach ($row in $unmatched) {
            if (-not ($entry.Ids -ccontains [string]$row.check_id)) { continue }
            if ([string]$row.severity -cne $entry.Severity) { continue }
            if ($entry.HasSubject -and ((Format-FaCell $row.subject) -cne $entry.Subject)) { continue }
            if ($entry.HasMessage -and ((Format-FaCell $row.message) -cne $entry.Message)) { continue }
            $found = $row; break
        }
        if ($null -eq $found) {
            $problems += ('expected ' + $entry.Wanted + '(' + $entry.Severity + ') is not shown as expected')
        } else {
            $unmatched = @($unmatched | Where-Object { -not [object]::ReferenceEquals($_, $found) })
        }
    }
    if ($unmatched.Count -gt 0) { $problems += ('unexpected actionable row(s): ' + (Format-FaActionable $unmatched)) }
    $null = Add-FaCheck $Scenario ($problems.Count -eq 0) `
        $(if ($problems.Count -eq 0) {
            ('overall=' + $overall + ' errors=' + [string]$errors + ' warnings=' + [string]$warnings +
             ' actionable=' + (Format-FaActionable $rows))
          } else { ($problems -join '; ') + '; shown actionable=' + (Format-FaActionable $rows) })
    return $surface
}

# ===========================================================================
# PREFLIGHT - EVERYTHING THAT CAN BE REFUSED BEFORE EXCEL STARTS
# ===========================================================================
$manifestPath    = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath     = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath  = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasesPath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$casesPath       = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$p7InspectPath   = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$modelCheckPath  = Join-Path $BuildDir 'phase9_model_check_inspection.json'
$methodologyPath = Join-Path $BuildDir 'phase10_methodology_inspection.json'
$protectionPath  = Join-Path $BuildDir 'phase10_protection_inspection.json'
$resetPath       = Join-Path $BuildDir 'phase10_reset_inspection.json'
$artefacts = @($manifestPath, $inspectPath, $simInspectPath, $gateBCasesPath, $casesPath,
               $p7InspectPath, $modelCheckPath, $methodologyPath, $protectionPath, $resetPath)
foreach ($required in $artefacts) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ('REFUSED, BEFORE EXCEL WAS STARTED: ' + $required +
                    ' does not exist. Build Stage A first.') -ForegroundColor Red
        exit 1
    }
}

$manifest    = Get-Content -LiteralPath $manifestPath    -Raw | ConvertFrom-Json
$inspection  = Get-Content -LiteralPath $inspectPath     -Raw | ConvertFrom-Json
$simInspect  = Get-Content -LiteralPath $simInspectPath  -Raw | ConvertFrom-Json
$gateBCases  = Get-Content -LiteralPath $gateBCasesPath  -Raw | ConvertFrom-Json
$cases       = Get-Content -LiteralPath $casesPath       -Raw | ConvertFrom-Json
$p7          = Get-Content -LiteralPath $p7InspectPath   -Raw | ConvertFrom-Json
$projection  = Get-Content -LiteralPath $modelCheckPath  -Raw | ConvertFrom-Json
$methodology = Get-Content -LiteralPath $methodologyPath -Raw | ConvertFrom-Json
$protection  = Get-Content -LiteralPath $protectionPath  -Raw | ConvertFrom-Json
$reset       = Get-Content -LiteralPath $resetPath       -Raw | ConvertFrom-Json

# THE W4 MODEL: the accepted small fixture every Phase-7 to Phase-9 runner built.
$case = $null
foreach ($scenario in @($cases.scenarios)) { if ([string]$scenario.id -ceq 'W4') { $case = $scenario } }
if ($null -eq $case) { throw 'the acceptance corpus carries no W4 scenario' }
$model = $case.model
$suppliedSeed = [double]$case.supplied_seed
$durationYears = [int]$model.timeline.duration

# THE REQUEST SIZES, DERIVED: the business minimum, and the minimum plus one for
# the STALE drift. Nothing larger exists in this runner.
$acceptanceIterations = [int]$gateBCases.bounds.business_minimum_iterations
$staleIterations = $acceptanceIterations + 1
$iterationsName = [string]$simInspect.controls.monte_carlo_iterations.defined_name
$seedName = [string]$simInspect.controls.random_seed.defined_name

# THE STATE VOCABULARY, PROJECTED: derived statuses and attempt results.
$statusNotCalculated = [string]$p7.model_states.derived_status[0]
$statusCurrent       = [string]$p7.model_states.derived_status[1]
$statusStale         = [string]$p7.model_states.derived_status[2]
$statusInvalid       = [string]$p7.model_states.derived_status[3]
$attemptNone         = [string]$p7.model_states.attempt_result[0]
$attemptSuccess      = [string]$p7.model_states.attempt_result[1]
$attemptRefused      = [string]$p7.model_states.attempt_result[2]
$derivedVocabulary   = @($p7.model_states.derived_status | ForEach-Object { [string]$_ })
$severityError       = [string]$projection.vocabulary.actionable_severities[0]
$severityWarning     = [string]$projection.vocabulary.actionable_severities[1]
$groupCalculation    = [string]$projection.vocabulary.group_order[2]

# THE MODEL CHECK EXPECTATIONS, FROM THE PHASE-9 PROJECTION. The runner's request
# is the business minimum, which the Phase-9 contract places STRICTLY BELOW the
# recommendation, so the low-iteration advisory is shown as an actionable WARNING
# at every checkpoint; it refuses nothing and is never an error. The INVALID
# checkpoint shows the one declared Calculation ERROR with the refused driver as
# its subject and the simulation's invalidity only as context; the reset
# checkpoint shows the one Calculation WARNING that a NOT CALCULATED live state
# raises, with no subject. Nothing here is a literal check id.
if (-not ($acceptanceIterations -lt [int]$projection.advisory.threshold)) {
    Write-Host ('REFUSED, BEFORE EXCEL WAS STARTED: the business minimum ' + [string]$acceptanceIterations +
                ' is not below the advisory threshold ' + [string]$projection.advisory.threshold +
                ', so the Model Check expectations in this runner would not hold.') -ForegroundColor Red
    exit 1
}
$advisoryExpected = @{ Id = [string]$projection.advisory.check_id; Severity = [string]$projection.advisory.severity
                       Message = [string]$projection.advisory.message }
$calcErrorChecks = @($projection.evaluation.declared_checks | Where-Object {
    ([string]$_.group -ceq $groupCalculation) -and ([string]$_.severity -ceq $severityError) })
if ($calcErrorChecks.Count -ne 1) { throw ('the projection declares ' + [string]$calcErrorChecks.Count + ' Calculation ERROR checks; exactly one is expected') }
$calcWarningIds = @($projection.evaluation.declared_checks | Where-Object {
    ([string]$_.group -ceq $groupCalculation) -and ([string]$_.severity -ceq $severityWarning) } |
    ForEach-Object { [string]$_.check_id })
if ($calcWarningIds.Count -lt 1) { throw 'the projection declares no Calculation WARNING check' }
# A MODEL-WIDE BLANK SUBJECT IS THE EMPTY STRING, NOT A BLANK CELL. The Phase-9
# builder writes a check with no subject as `=""` - an empty STRING, which INDEX
# returns as one - precisely so that Excel does not read an empty cell back as a
# hard 0. The register therefore answers '' for these rows, and run 6 showed it:
# both historical Annual rows carried an empty subject. The expectation is the worksheet's
# representation, never the '<blank>' token Format-FaCell reserves for a $null
# cell read. The Subject constraint stays: it is what excludes the Annual
# WARNING whose subject is a profile confidence level.
$notCalculatedExpected = @{ AnyOf = $calcWarningIds; Severity = $severityWarning; Subject = '' }
# THE ANNUAL OUTPUTS BECOME HISTORICAL when the model they were produced for is
# invalidated after they were published - which is this runner's sequence - and
# the Phase-9 contract raises one Annual WARNING for the historical profile and
# one for the historical distributions, both without a subject. The Annual
# WARNING population comes from the projection; the two shown at the INVALID
# checkpoint are the two with no subject, and the third declared Annual WARNING
# names a profile confidence level as its subject and cannot fire beside a
# HISTORICAL profile.
$groupAnnual = [string]$projection.vocabulary.group_order[4]
$annualWarningIds = @($projection.evaluation.declared_checks | Where-Object {
    ([string]$_.group -ceq $groupAnnual) -and ([string]$_.severity -ceq $severityWarning) } |
    ForEach-Object { [string]$_.check_id })
if ($annualWarningIds.Count -lt 2) { throw ('the projection declares ' + [string]$annualWarningIds.Count + ' Annual WARNING checks; at least two are expected') }
$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning; Subject = '' }
$annualHistorical  = [string]$p7.handoff.distribution_states[2]
$profileHistorical = [string]$p7.handoff.profile_states[3]

# THE CALCULATION STATE BLOCK, from the Phase-5 inspection.
$calcSheet = [string]$inspection.calc.sheet
$calcStateBlock = $inspection.calc.scalar_blocks.calc_state
$calcValueColumn = [string]$calcStateBlock.value_column
$calcStatusRange = ($calcValueColumn + [string]$calcStateBlock.rows.calculation_status + ':' +
                    $calcValueColumn + [string]$calcStateBlock.rows.status_evaluated_at)
$calcAttemptCell = ($calcValueColumn + [string]$calcStateBlock.rows.last_attempt_result)
# THE SIMULATION RECORD'S LAST-ATTEMPT FIELD, from the Phase-6 run-identity
# projection; the reset projection's attempt_and_selector span begins on it, and
# the reset projection's calculation attempt_result_initial must name the same
# calc cell and the same NONE the Phase-7 vocabulary names. Two contracts, one answer.
$simRunIdentity = $simInspect.sim_data.run_identity
$simAttemptCell = ([string]$simRunIdentity.value_column + [string]$simRunIdentity.rows.last_attempt_result)
$resetAttemptInitial = $reset.publications.calculation.attempt_result_initial
if ([string]$resetAttemptInitial.cell -cne $calcAttemptCell) { throw ('the reset projection seeds ' + [string]$resetAttemptInitial.cell + ' after a reset; the calculation state block names ' + $calcAttemptCell) }
if ([string]$resetAttemptInitial.value -cne $attemptNone) { throw ('the reset projection seeds ' + [string]$resetAttemptInitial.value + ' after a reset; the accepted vocabulary names ' + $attemptNone) }
if ([string]$reset.publications.simulation.cleared.attempt_and_selector -notlike ($simAttemptCell + ':*')) { throw ('the reset projection clears ' + [string]$reset.publications.simulation.cleared.attempt_and_selector + ', which does not begin on the simulation last-attempt field ' + $simAttemptCell) }
$calcPersistedStatusCell = ($calcValueColumn + [string]$calcStateBlock.rows.calculation_status)
$calcFingerprintCell = ($calcValueColumn + [string]$calcStateBlock.rows.last_successful_fingerprint)

# THE METADATA BLOCK, from the methodology projection: label column, text column,
# and the rows the emitter wrote. The Source Revision row is required to exist.
$metadataRows = @($methodology.metadata)
if ($metadataRows.Count -lt 1) { throw 'the methodology projection carries no metadata rows' }
$metadataFirst = ($metadataRows | Measure-Object -Property row -Minimum).Minimum
$metadataLast = ($metadataRows | Measure-Object -Property row -Maximum).Maximum
$metadataRange = ([string]$methodology.text_column + [string]$metadataFirst + ':' +
                  [string]$methodology.text_column + [string]$metadataLast)
$metadataLabelRange = ([string]$methodology.label_column + [string]$metadataFirst + ':' +
                       [string]$methodology.label_column + [string]$metadataLast)
$sourceRevisionRow = $null
foreach ($row in $metadataRows) { if ([string]$row.label -ceq 'Source Revision') { $sourceRevisionRow = $row } }
if ($null -eq $sourceRevisionRow) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED: the methodology projection has no Source Revision row.' -ForegroundColor Red
    exit 1
}

# THE SOURCE REVISION, derived before anything else runs.
$revision = $null
try { $revision = Get-FaSourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ('The working tree is modified, so the builder would have stamped (dirty) and ' +
                'a final acceptance could not be attributed to one commit:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase10-final-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in $artefacts) { Copy-Item -LiteralPath $artefact -Destination $tempRoot }
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse
$script:FaPath = Join-Path $tempRoot 'phase10_final_acceptance.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

# ===========================================================================
# 1. STAGE-B BOOTSTRAP AND REOPEN VERIFICATION - THE ACCEPTED BUILD, UNCHANGED
# ===========================================================================
$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

Write-FaLine 'PCCM - PHASE 10 FINAL WINDOWS ACCEPTANCE'
Write-FaLine '========================================'
Write-FaLine ('expected source revision : ' + $revision.Expected)
Write-FaLine ('fixture                  : ' + [string]$case.id + ' model, ' +
              [string]@($model.cost_lines).Count + ' cost lines, ' +
              [string]@($model.risks).Count + ' risks, ' + [string]$durationYears + ' years, seed ' +
              [string]$suppliedSeed)
Write-FaLine ('iterations               : ' + [string]$acceptanceIterations + ' (business minimum); ' +
              [string]$staleIterations + ' once, for STALE')
Write-FaLine ''
$null = Add-FaCheck 'bootstrap' $bootstrapOk ('build_stage_b.ps1 exit ' + [string]$bootstrapExit +
    '; the bootstrap reopen-verified FileFormat, sheets, CodeNames, modules and buttons')

# ===========================================================================
# THE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$rel = New-ReleaseLedger
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0
$observedRevision = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1
    # THE OPEN/READY BOUNDARY: bounded, transient-only, before the first workbook read.
    $null = Assert-FaWorkbookReady -Workbook $wb -ExpectedPath $stageBPath -Session 'main'

    # 6a. THE PERSISTED CALCULATION HISTORY, read from the model's own persisted
    # cells BEFORE any accessor evaluates: no calculation has ever been
    # committed, so the last-evaluated status is NOT CALCULATED, the attempt
    # result is NONE and the last successful fingerprint is blank. Phase 9
    # settled the distinction: the LIVE state of the untouched workbook is
    # INVALID (6b below); the persisted history is not, and the two are asserted
    # separately. The compile check that follows evaluates and persists the live
    # status, which is why this read comes first.
    $persistedStatus = Format-FaCell ((Get-FaBlock -Workbook $wb -SheetName $calcSheet -Address $calcPersistedStatusCell).Rect)
    $persistedAttempt = Format-FaCell ((Get-FaBlock -Workbook $wb -SheetName $calcSheet -Address $calcAttemptCell).Rect)
    $persistedFingerprint = Format-FaCell ((Get-FaBlock -Workbook $wb -SheetName $calcSheet -Address $calcFingerprintCell).Rect)
    $null = Add-FaCheck 'state.initial.persisted' `
        (($persistedStatus -ceq $statusNotCalculated) -and ($persistedAttempt -ceq $attemptNone) -and
         ($persistedFingerprint -ceq '<blank>')) `
        ('persisted status ' + $persistedStatus + '; last attempt ' + $persistedAttempt +
         '; last successful fingerprint ' + $persistedFingerprint)

    # 3. THE ACCEPTED COMPILE CHECK: the first Application.Run compiles the project.
    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-FaCheck 'compile' ([string]::IsNullOrWhiteSpace($compileFailure)) `
        $(if ([string]::IsNullOrWhiteSpace($compileFailure)) { 'the VBAProject compiles in real Excel' } else { $compileFailure })

    # 2b. SHEETS AND CODENAMES, SET-BASED AGAINST THE MANIFEST.
    $expectedSheets = @{}
    foreach ($sheet in @($manifest.sheets)) { $expectedSheets[[string]$sheet.name] = [string]$sheet.codename }
    $observedSheets = @{}
    $sheets = $null
    try {
        $sheets = $wb.Worksheets
        $count = [int]$sheets.Count
        for ($index = 1; $index -le $count; $index++) {
            $ws = $null
            try {
                $ws = $sheets.Item($index)
                $observedSheets[[string]$ws.Name] = [string]$ws.CodeName
            } finally {
                if ($null -ne $ws) { Release-Transient $ws 'Worksheet'; $ws = $null }
            }
        }
    } finally {
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets'; $sheets = $null }
    }
    $sheetProblems = @()
    foreach ($name in $expectedSheets.Keys) {
        if (-not $observedSheets.ContainsKey($name)) { $sheetProblems += ('missing ' + $name); continue }
        if ($observedSheets[$name] -cne $expectedSheets[$name]) {
            $sheetProblems += ($name + ' CodeName ' + $observedSheets[$name] + ', expected ' + $expectedSheets[$name])
        }
    }
    foreach ($name in $observedSheets.Keys) {
        if (-not $expectedSheets.ContainsKey($name)) { $sheetProblems += ('undeclared ' + $name) }
    }
    $null = Add-FaCheck 'sheets' ($sheetProblems.Count -eq 0) `
        $(if ($sheetProblems.Count -eq 0) { ([string]$observedSheets.Count + ' worksheets, every name and CodeName as the manifest declares') }
          else { $sheetProblems -join '; ' })

    # 2c. THE VBA MODULE SET == THE MANIFEST-DECLARED SET, read BEFORE the shim is
    # imported. The declared set is the manifest's modules, its document module
    # and one document module per sheet CodeName; nothing is a literal.
    $expectedModules = @{}
    foreach ($module in @($manifest.vba.modules)) { $expectedModules[[string]$module.name] = $true }
    $expectedModules[[string]$manifest.vba.document_module.component] = $true
    foreach ($sheet in @($manifest.sheets)) { $expectedModules[[string]$sheet.codename] = $true }
    $observedModules = @{}
    $vbproj = $null; $comps = $null
    try {
        $vbproj = $wb.VBProject
        $comps = $vbproj.VBComponents
        $total = [int]$comps.Count
        for ($index = 1; $index -le $total; $index++) {
            $comp = $null
            try { $comp = $comps.Item($index); $observedModules[[string]$comp.Name] = $true }
            finally { if ($null -ne $comp) { Release-Transient $comp ('VBComponent[' + [string]$index + ']'); $comp = $null } }
        }
    } finally {
        if ($null -ne $comps)  { Release-Transient $comps  'VBComponents'; $comps  = $null }
        if ($null -ne $vbproj) { Release-Transient $vbproj 'VBProject';    $vbproj = $null }
    }
    $moduleProblems = @()
    foreach ($name in $expectedModules.Keys) { if (-not $observedModules.ContainsKey($name)) { $moduleProblems += ('missing ' + $name) } }
    foreach ($name in $observedModules.Keys) { if (-not $expectedModules.ContainsKey($name)) { $moduleProblems += ('undeclared ' + $name) } }
    $null = Add-FaCheck 'modules' ($moduleProblems.Count -eq 0) `
        $(if ($moduleProblems.Count -eq 0) { ([string]$observedModules.Count + ' components observed; the set equals the manifest-declared set') }
          else { $moduleProblems -join '; ' })

    # 4. RELEASE METADATA FROM THE ACTUAL WORKBOOK. Every projected row's label
    # and value, and the Source Revision compared to the checkout.
    $labelBlock = Get-FaBlock -Workbook $wb -SheetName ([string]$methodology.sheet) -Address $metadataLabelRange
    $valueBlock = Get-FaBlock -Workbook $wb -SheetName ([string]$methodology.sheet) -Address $metadataRange
    $shown = @{}
    $metadataProblems = @()
    foreach ($row in $metadataRows) {
        $offset = [int]$row.row - $metadataFirst + 1
        $label = Format-FaCell (Get-FaBlockCell -Block $labelBlock -Row $offset -Column 1)
        $value = Format-FaCell (Get-FaBlockCell -Block $valueBlock -Row $offset -Column 1)
        $shown[[string]$row.label] = $value
        if ($label -cne [string]$row.label) { $metadataProblems += ('row ' + [string]$row.row + ' is labelled ' + $label + ', expected ' + [string]$row.label) }
        if ($value -cne [string]$row.value) { $metadataProblems += ([string]$row.label + ' reads ' + $value + ', the projection says ' + [string]$row.value) }
    }
    $null = Add-FaCheck 'metadata.rows' ($metadataProblems.Count -eq 0) `
        $(if ($metadataProblems.Count -eq 0) { ([string]$metadataRows.Count + ' metadata rows, each label and value as projected') }
          else { $metadataProblems -join '; ' })
    $null = Add-FaCheck 'metadata.model-version' ($shown['PCCM Model Version'] -ceq '1.0.0') ('PCCM Model Version = ' + $shown['PCCM Model Version'])
    $null = Add-FaCheck 'metadata.builder-version' ($shown['Builder Version'] -ceq '1.0.0') ('Builder Version = ' + $shown['Builder Version'])
    $null = Add-FaCheck 'metadata.build-phase' ($shown['Build Phase'] -ceq 'Release 1.0 - Production') ('Build Phase = ' + $shown['Build Phase'])
    $observedRevision = $shown['Source Revision']
    $null = Add-FaCheck 'metadata.source-revision' ($observedRevision -ceq $revision.Expected) `
        ('expected ' + $revision.Expected + '; observed ' + $observedRevision)

    # 5. INITIAL PROTECTION, through the accepted shim.
    Import-FaFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.initial'

    # 6b. THE LIVE, DERIVED STATE of the untouched workbook, as Phase 9 accepted
    # it: required inputs are unresolved, so the calculation is INVALID, the
    # simulation is INVALID as its consequence, and nothing annual was produced.
    $states0 = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'state.initial.live' `
        (($states0.Calculation -ceq $statusInvalid) -and ($states0.Simulation -ceq $statusInvalid) -and
         ($states0.Annual -like 'NOT PRODUCED*') -and ($states0.Profile -like 'NOT PRODUCED*')) `
        (Format-FaStates $states0)

    # 7. THE ACCEPTED W4 FIXTURE, inside the setup window, exactly as the benchmark
    # builds it: the fixture writes ClearContents on locked tables and needs it.
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'fixture'
    $applied = ''
    try {
        $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
            -Inspection $inspection -Model $model)
    } finally {
        $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'fixture'
    }
    $null = Add-FaCheck 'fixture' ($applied -like 'OK|*') $applied
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$acceptanceIterations)
    Set-NamedValue -Workbook $wb -DefinedName $seedName -Value $suppliedSeed

    # 8. THE REAL STRUCTURAL USER WORKFLOW, with no window: production opens its own.
    $costRegister = Get-FaRegister -Manifest $manifest -Key 'cost_lines'
    $riskRegister = Get-FaRegister -Manifest $manifest -Key 'risk_register'
    foreach ($pair in @(
        @{ Register = $costRegister; Add = 'PCCM_AddCostLine'; Delete = 'PCCM_DeleteCostLineById'; Label = 'cost-line' },
        @{ Register = $riskRegister; Add = 'PCCM_AddRisk';     Delete = 'PCCM_DeleteRiskById';     Label = 'risk' })) {
        $before = @(Get-IdColumnValues -Workbook $wb -Info $pair.Register)
        $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation ([string]$pair.Add) -Stage ('structural ' + [string]$pair.Label)
        $after = @(Get-IdColumnValues -Workbook $wb -Info $pair.Register)
        $added = @($after | Where-Object { $before -notcontains $_ })
        if ($added.Count -ne 1) {
            throw ([string]$pair.Add + ' reported success but the register gained ' + [string]$added.Count + ' identifier(s): ' + ($added -join ', '))
        }
        $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation ([string]$pair.Delete) `
            -Argument ([string]$added[0]) -WithArgument -Stage ('structural ' + [string]$pair.Label)
        $restored = @(Get-IdColumnValues -Workbook $wb -Info $pair.Register)
        $null = Add-FaCheck ('structural.add-delete.' + [string]$pair.Label) `
            (($restored -join ',') -ceq ($before -join ',')) `
            ([string]$pair.Add + ' issued ' + [string]$added[0] + '; ' + [string]$pair.Delete + ' removed it; ' +
             [string]$restored.Count + ' identifier(s) as before')
    }
    $timelineResult = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_ApplyTimeline' -Stage 'structural timeline')
    $structuralReport = [string](Assert-Phase5StructurallyCoherent -Excel $excel -Stage 'after the structural workflow')
    $null = Add-FaCheck 'structural.apply-timeline' ($timelineResult -like 'OK|*') ($timelineResult + '; structural report empty')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-structural'

    # 9. CALCULATE -> CURRENT.
    $calcResult = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'calculate')
    $fingerprint = Get-FaRunText -Excel $excel -Procedure 'PCCM_CalculationFingerprint'
    $statesCalc = Get-FaStates -Excel $excel
    $attemptCalc = Get-FaRunText -Excel $excel -Procedure 'PCCM_CalculationAttemptResult'
    $null = Add-FaCheck 'calculate.current' `
        (($calcResult -like 'OK|*') -and ($statesCalc.Calculation -ceq $statusCurrent) -and
         ($attemptCalc -ceq $attemptSuccess) -and (-not [string]::IsNullOrWhiteSpace($fingerprint))) `
        ($calcResult + '; ' + (Format-FaStates $statesCalc) + '; attempt=' + $attemptCalc + '; fingerprint=' + $fingerprint)
    $excel.Calculate()
    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.calculated' `
        -Expected @($advisoryExpected)

    # WORKSHEET SAFETY, as P9-1 proved it: a recalculation rewrites no persisted status cell.
    $statusBefore = Get-FaBlockDigest -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    $excel.Calculate(); $excel.Calculate()
    $statusAfter = Get-FaBlockDigest -Workbook $wb -SheetName $calcSheet -Address $calcStatusRange
    $null = Add-FaCheck 'modelcheck.worksheet-safety' ($statusBefore -ceq $statusAfter) 'two recalculations rewrote no persisted status cell'

    # 9. SIMULATION, SENSITIVITY, ANNUAL at the business minimum.
    $simResult = [string](Invoke-Phase6Simulation -Excel $excel)
    $statesSim = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'simulation.current' (($simResult -like 'OK|*') -and ($statesSim.Simulation -ceq $statusCurrent)) `
        ([string]$acceptanceIterations + ' iterations; ' + $simResult + '; ' + (Format-FaStates $statesSim))
    $sensResult = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RunSensitivity'
    $null = Add-FaCheck 'sensitivity' ($sensResult -like 'OK|*') $sensResult
    $annualResult = Invoke-FaEndpoint -Excel $excel -Operation ([string]$p7.command_surface.annual_endpoint)
    $statesAnnual = Get-FaStates -Excel $excel
    $yearCount = Get-FaRunText -Excel $excel -Procedure 'PCCM_AnnualYearCount'
    $null = Add-FaCheck 'annual' `
        (($annualResult -like 'OK|*') -and ($statesAnnual.Annual -ceq $statusCurrent) -and
         ($statesAnnual.Profile -ceq $statusCurrent) -and ([int]$yearCount -eq $durationYears)) `
        ($annualResult + '; ' + (Format-FaStates $statesAnnual) + '; years=' + $yearCount)
    $excel.Calculate()
    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.simulated' `
        -Expected @($advisoryExpected)
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-commands'

    # 11. STALE: the request drifts by one iteration; the state is derived, not written.
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$staleIterations)
    $statesStale = Get-FaStates -Excel $excel
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$acceptanceIterations)
    $statesBack = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'state.stale' (($statesStale.Simulation -ceq $statusStale) -and ($statesBack.Simulation -ceq $statusCurrent)) `
        ('at ' + [string]$staleIterations + ': ' + (Format-FaStates $statesStale) + '; restored: ' + (Format-FaStates $statesBack))

    # 11/12. INVALID, and REFUSED as the attempt outcome of the refused Calculate.
    $victimRow = 1
    $maxOrdinal = Get-FaColumnOrdinal -Register $costRegister -ColumnKey 'unit_cost_max'
    $minOrdinal = Get-FaColumnOrdinal -Register $costRegister -ColumnKey 'unit_cost_min'
    $costBody = @(Get-TableBody -Workbook $wb -SheetName ([string]$costRegister.sheet) -TableName ([string]$costRegister.table_name))
    $victimId = [string]$costBody[$victimRow - 1][0]
    $originalMax = [double]$costBody[$victimRow - 1][$maxOrdinal - 1]
    $invalidMax = [double]$costBody[$victimRow - 1][$minOrdinal - 1] - 1.0
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) -TableName ([string]$costRegister.table_name) `
        -RowIndex $victimRow -ColumnIndex $maxOrdinal -Value $invalidMax
    $refusedCalc = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_Calculate'
    $statesInvalid = Get-FaStates -Excel $excel
    $attemptInvalid = Get-FaRunText -Excel $excel -Procedure 'PCCM_CalculationAttemptResult'
    $null = Add-FaCheck 'state.invalid' (($refusedCalc -like 'FAIL|*') -and ($statesInvalid.Calculation -ceq $statusInvalid)) `
        ($victimId + ' maximum below minimum; ' + $refusedCalc + '; ' + (Format-FaStates $statesInvalid))
    $null = Add-FaCheck 'refused.outcome.calculate' `
        (($attemptInvalid -ceq $attemptRefused) -and ($derivedVocabulary -cnotcontains $attemptRefused)) `
        ('attempt result ' + $attemptInvalid + '; derived status ' + $statesInvalid.Calculation + '; REFUSED is not a derived status')
    $excel.Calculate()
    $refusalSubject = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckRefusalSubject'
    $null = Add-FaCheck 'modelcheck.invalid.subject' ($refusalSubject -ceq $victimId) `
        ('the refusal subject is ' + $refusalSubject + '; the invalidated driver is ' + $victimId)
    # THE ANNUAL OUTPUTS PUBLISHED BEFORE THE INVALIDATION ARE NOW HISTORICAL.
    $null = Add-FaCheck 'state.invalid.annual-historical' `
        (($statesInvalid.Annual -ceq $annualHistorical) -and ($statesInvalid.Profile -ceq $profileHistorical)) `
        ('annual=' + $statesInvalid.Annual + ' profile=' + $statesInvalid.Profile)
    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.invalid' `
        -Expected @(@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId },
                    $advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) -TableName ([string]$costRegister.table_name) `
        -RowIndex $victimRow -ColumnIndex $maxOrdinal -Value $originalMax
    $recalcResult = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'calculate after restore')
    $fingerprintBack = Get-FaRunText -Excel $excel -Procedure 'PCCM_CalculationFingerprint'
    $statesRestored = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'state.current-restored' `
        (($recalcResult -like 'OK|*') -and ($statesRestored.Calculation -ceq $statusCurrent) -and ($fingerprintBack -ceq $fingerprint)) `
        ((Format-FaStates $statesRestored) + '; fingerprint identical to the first CURRENT')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-refusal'

    # 14. REPAIR PROFILING, through the real entry point. The window creates the
    # preconditions only; production repairs, or refuses, under its own protection.
    $grid = Get-FaGrid -Manifest $manifest -Key 'cost_profiling'
    $gridSheet = [string]$grid.sheet; $gridTable = [string]$grid.table_name
    $fixedColumns = @($grid.fixed_columns).Count
    $inputFingerprint0 = Get-FaRunText -Excel $excel -Procedure 'PCCM_CurrentInputFingerprint'
    $gridBody0 = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $gridDigest0 = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable

    # 14a. NO-OP.
    $noop = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $null = Add-FaCheck 'repair.noop' `
        (($noop -like 'OK|*') -and ($noop -like '*Nothing was changed*') -and
         ((Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable) -ceq $gridDigest0) -and
         ((Get-FaRunText -Excel $excel -Procedure 'PCCM_CurrentInputFingerprint') -ceq $inputFingerprint0)) `
        ($noop + '; grid and input fingerprint unchanged')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-noop'

    # 14b. MISSING ROW: the second driver's profiling row is deleted inside the
    # window; Repair must recreate it id-only with blank weights and touch no other row.
    $registerIds = @(Get-IdColumnValues -Workbook $wb -Info $costRegister)
    $missingId = [string]$registerIds[1]
    $missingRowIndex = Find-FaTableRow -Body $gridBody0 -Id $missingId
    if ($missingRowIndex -lt 1) { throw ('the profiling grid carries no row for ' + $missingId) }
    $missingRowBefore = @($gridBody0[$missingRowIndex - 1])
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.missing-row'
    try { Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.missing-row' }
    $repairMissing = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $gridBody1 = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $repairProblems = @()
    if ($repairMissing -notlike 'OK|*') { $repairProblems += ('the repair did not succeed: ' + $repairMissing) }
    $orderedIds = @($gridBody1 | ForEach-Object { [string]$_[0] } | Where-Object { $_ -ne '' })
    if (($orderedIds -join ',') -cne ($registerIds -join ',')) {
        $repairProblems += ('grid ids ' + ($orderedIds -join ',') + ' are not the register order ' + ($registerIds -join ','))
    }
    $restoredIndex = Find-FaTableRow -Body $gridBody1 -Id $missingId
    if ($restoredIndex -lt 1) { $repairProblems += ($missingId + ' was not restored') }
    else {
        $restoredRow = @($gridBody1[$restoredIndex - 1])
        for ($c = $fixedColumns; $c -lt $restoredRow.Count; $c++) {
            if ([string]$restoredRow[$c] -ne '') { $repairProblems += ($missingId + ' column ' + [string]($c + 1) + ' is ' + [string]$restoredRow[$c] + ', not blank') }
        }
    }
    foreach ($row in $gridBody0) {
        $id = [string]$row[0]
        if ($id -ceq $missingId) { continue }
        $index1 = Find-FaTableRow -Body $gridBody1 -Id $id
        if ($index1 -lt 1) { $repairProblems += ($id + ' vanished'); continue }
        if ((@($gridBody1[$index1 - 1]) -join [char]31) -cne (@($row) -join [char]31)) { $repairProblems += ($id + ' changed') }
    }
    $null = Add-FaCheck 'repair.missing-row' ($repairProblems.Count -eq 0) `
        $(if ($repairProblems.Count -eq 0) { ($repairMissing + '; ' + $missingId + ' recreated id-only with blank weights; every other row byte-identical') }
          else { $repairProblems -join '; ' })
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-missing'
    # The user's weights for the recreated row are typed back into the UNLOCKED
    # weight cells, the way a user would, and the fixture state is proved recovered
    # by the calculation fingerprint further down.
    for ($c = $fixedColumns; $c -lt $missingRowBefore.Count; $c++) {
        $weight = [string]$missingRowBefore[$c]
        if ($weight -ne '') {
            Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable `
                -RowIndex $restoredIndex -ColumnIndex ($c + 1) -Value ([double]$weight)
        }
    }

    # 14c. ORDER: the first two rows are swapped inside the window; Repair must
    # restore register order with every weight travelling with its id, and the
    # input fingerprint must be what it was before the swap.
    $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_ApplyTimeline' -Stage 'resync after the missing-row repair'
    $inputFingerprint1 = Get-FaRunText -Excel $excel -Procedure 'PCCM_CurrentInputFingerprint'
    $gridBody2 = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $gridDigest2 = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $rowA = @($gridBody2[0]); $rowB = @($gridBody2[1])
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.order'
    try {
        for ($c = 0; $c -lt $rowA.Count; $c++) {
            $valueA = [string]$rowA[$c]; $valueB = [string]$rowB[$c]
            if ($c -lt $fixedColumns) {
                Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 1 -ColumnIndex ($c + 1) -Value $valueB
                Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex ($c + 1) -Value $valueA
            } else {
                if ($valueB -eq '') { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 1 -ColumnIndex ($c + 1) -Value $null }
                else { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 1 -ColumnIndex ($c + 1) -Value ([double]$valueB) }
                if ($valueA -eq '') { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex ($c + 1) -Value $null }
                else { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex ($c + 1) -Value ([double]$valueA) }
            }
        }
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.order' }
    $swappedDigest = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    if ($swappedDigest -ceq $gridDigest2) { throw 'the order precondition did not take: the grid is unchanged after the swap' }
    $repairOrder = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $orderDigest = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $inputFingerprint2 = Get-FaRunText -Excel $excel -Procedure 'PCCM_CurrentInputFingerprint'
    $null = Add-FaCheck 'repair.order' `
        (($repairOrder -like 'OK|*') -and ($orderDigest -ceq $gridDigest2) -and ($inputFingerprint2 -ceq $inputFingerprint1)) `
        ($repairOrder + '; grid byte-identical to before the swap; input fingerprint identical')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-order'

    # 14d. AMBIGUOUS DUPLICATE: the second row is given the first row's id and a
    # different weight inside the window; Repair must refuse, name the id, and
    # change nothing. The precondition is then undone through the window and a
    # no-op repair proves the grid is exactly what it was.
    $duplicateId = [string]$rowA[0]
    $secondId = [string]$rowB[0]
    $weightColumn = $fixedColumns + 1
    $secondWeight = [string]$rowB[$fixedColumns]
    $alteredWeight = $(if ($secondWeight -eq '') { 0.5 } else { [double]$secondWeight + 0.1 })
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.duplicate'
    try {
        Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex 1 -Value $duplicateId
        Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex $weightColumn -Value ([double]$alteredWeight)
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.duplicate' }
    $damagedDigest = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $repairDuplicate = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $afterRefusal = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $null = Add-FaCheck 'repair.duplicate-refused' `
        (($repairDuplicate -like 'FAIL|*') -and ($repairDuplicate -like ('*' + $duplicateId + '*')) -and
         ($repairDuplicate -like '*more than one row*') -and ($afterRefusal -ceq $damagedDigest)) `
        ($repairDuplicate + '; the grid is exactly as it was before the refusal')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-refusal'
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.duplicate-undo'
    try {
        Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex 1 -Value $secondId
        if ($secondWeight -eq '') { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex $weightColumn -Value $null }
        else { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 -ColumnIndex $weightColumn -Value ([double]$secondWeight) }
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.duplicate-undo' }
    $repairAfterUndo = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $undoneDigest = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $null = Add-FaCheck 'repair.restored' (($repairAfterUndo -like 'OK|*') -and ($undoneDigest -ceq $gridDigest2)) `
        ($repairAfterUndo + '; grid byte-identical to before the duplicate')
    # --- P10-R2 REPAIR CONTRACT SCENARIOS: begin ------------------------------
    # Added at the bounded correction round after the independent review, against
    # the settled contract (section 5 and matrix rows F, G, G2, M): width growth,
    # a blank-only shrink, a typed-zero shrink refused, a nonzero shrink refused,
    # a fixed-width populated non-100% profile refused, an all-blank profile
    # allowed, a populated zero-total signed profile refused, both Repair
    # failpoints rolled back, both grids compared around every refusal, and
    # protection after each. Every precondition is made inside the setup window
    # or through an unlocked weight cell; every repair or refusal is production's.
    $riskGrid = Get-FaGrid -Manifest $manifest -Key 'risk_profiling'
    $riskSheet = [string]$riskGrid.sheet; $riskTable = [string]$riskGrid.table_name
    $baselineCost = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $baselineRisk = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
    # THE SAME BASELINES AS PLAIN DATA, for the diagnosis of repair.grids-restored.
    $baselineCostBody = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $baselineRiskBody = @(Get-TableBody -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $baselineCostColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $baselineRiskColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $rowOneWeights = @($gridBody2[0])
    $rowOneYears = @()
    for ($c = $fixedColumns; $c -lt $rowOneWeights.Count; $c++) { $rowOneYears += [string]$rowOneWeights[$c] }
    $rowOneId = [string]$rowOneWeights[0]
    if ($rowOneYears.Count -ne $durationYears) { throw ('the cost profiling grid holds ' + [string]$rowOneYears.Count + ' project years, the applied timeline ' + [string]$durationYears) }

    function Restore-FaRowOne {
        for ($y = 1; $y -le $rowOneYears.Count; $y++) {
            $w = $rowOneYears[$y - 1]
            if ($w -eq '') { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year $y -Weight $null }
            else { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year $y -Weight ([double]$w) }
        }
    }
    function Assert-FaGridsUnchanged {
        param([string]$Scenario, [string]$Announcement)
        $costNow = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
        $riskNow = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
        $null = Add-FaCheck $Scenario (($Announcement -like 'FAIL|*') -and ($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)) `
            ($Announcement + '; both profiling grids byte-identical to before the refusal')
    }
    function Save-FaGridsBefore {
        $script:FaCostBefore = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
        $script:FaRiskBefore = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
    }

    # (a) WIDTH GROWTH. THE FIXTURE FIRST, FOR EVERY ROW. Losing the last project
    # year must leave EVERY retained populated row totalling 100%, because the
    # production semantic gate assesses every retained row of BOTH grids before
    # any width is repaired. Final acceptance run 8 proved it: only row one had
    # been shaped, CL-002 still totalled 0.75 without its last year, and Repair
    # refused - correctly. So every populated cost row becomes [0.5, 0.5, 0, ..., 0]:
    # 100% inside the years that will remain, a typed 0 in the year that will be
    # deleted. The risk grid is not written and is proved to satisfy the rule as it
    # stands. The precondition is asserted from the grids themselves before the
    # defect is made, and the original fixture is put back exactly afterwards.
    $riskFixedColumns = @($riskGrid.fixed_columns).Count
    $originalCostBody = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    function Restore-FaCostRows {
        param([object[]]$Original)
        $current = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
        foreach ($entry in $Original) {
            $row = @($entry)
            if ([string]$row[0] -eq '') { continue }
            $index = Find-FaTableRow -Body $current -Id ([string]$row[0])
            if ($index -lt 1) { throw ([string]$row[0] + ' has no profiling row to restore') }
            for ($y = 1; $y -le $durationYears; $y++) {
                $w = [string]$row[$fixedColumns + $y - 1]
                if ($w -eq '') { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex $index -Year $y -Weight $null }
                else { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex $index -Year $y -Weight ([double]$w) }
            }
        }
    }
    for ($r = 0; $r -lt $originalCostBody.Count; $r++) {
        if ([string]$originalCostBody[$r][0] -eq '') { continue }
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex ($r + 1) -Year 1 -Weight 0.5
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex ($r + 1) -Year 2 -Weight 0.5
        for ($y = 3; $y -le $durationYears; $y++) { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex ($r + 1) -Year $y -Weight 0.0 }
    }
    $growthBefore = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $riskBeforeGrowth = @(Get-TableBody -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $fixtureProblems = @()
    $fixtureProblems += @(Get-FaProfileProblems -Body $growthBefore -FixedColumns $fixedColumns -YearCount ($durationYears - 1) -Label 'Cost Profiling')
    foreach ($entry in $growthBefore) {
        $row = @($entry)
        if ([string]$row[0] -eq '') { continue }
        if ([string]$row[$fixedColumns + $durationYears - 1] -ne '0') { $fixtureProblems += ('Cost Profiling ' + [string]$row[0] + ' project year ' + [string]$durationYears + ' is ' + [string]$row[$fixedColumns + $durationYears - 1] + ', not a typed 0') }
    }
    $fixtureProblems += @(Get-FaProfileProblems -Body $riskBeforeGrowth -FixedColumns $riskFixedColumns -YearCount $durationYears -Label 'Risk Profiling')
    $null = Add-FaCheck 'repair.width-growth.fixture' ($fixtureProblems.Count -eq 0) `
        $(if ($fixtureProblems.Count -eq 0) { ('every populated cost row totals 100% over project years 1-' + [string]($durationYears - 1) + ' with a typed 0 in year ' + [string]$durationYears + '; every populated risk row totals 100%') } else { $fixtureProblems -join '; ' })
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.width-growth'
    try { Remove-FaLastTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.width-growth' }
    $grown = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $growthAfter = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $growthProblems = @()
    if ($grown -notlike 'OK|*') { $growthProblems += ('the growth repair did not succeed: ' + $grown) }
    if ($grown -notlike '*left blank*') { $growthProblems += 'the success wording does not say what was left blank' }
    for ($r = 0; $r -lt $growthBefore.Count; $r++) {
        $before = @($growthBefore[$r]); $after = @($growthAfter[$r])
        if ([string]$before[0] -eq '') { continue }
        if ($after.Count -ne $before.Count) { $growthProblems += ([string]$before[0] + ' has ' + [string]$after.Count + ' columns, expected ' + [string]$before.Count); continue }
        for ($c = 0; $c -lt $before.Count - 1; $c++) {
            if ([string]$after[$c] -cne [string]$before[$c]) { $growthProblems += ([string]$before[0] + ' column ' + [string]($c + 1) + ' changed from ' + [string]$before[$c] + ' to ' + [string]$after[$c]) }
        }
        if ([string]$after[$before.Count - 1] -ne '') { $growthProblems += ([string]$before[0] + ' regrown project year ' + [string]$durationYears + ' is ' + [string]$after[$before.Count - 1] + ', not blank') }
    }
    $null = Add-FaCheck 'repair.width-growth' ($growthProblems.Count -eq 0) `
        $(if ($growthProblems.Count -eq 0) { ($grown + '; the regrown project year is blank on every row and every existing cell is unchanged') } else { $growthProblems -join '; ' })
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-growth'
    # THE LOCK STATE OF THE REGROWN CELLS, INSPECTED WHILE PROTECTION IS ON. The
    # accepted rule is Stage A's: the year columns over every reserved body row
    # are unlocked, keyed or not, and the fixed columns stay locked - the
    # protection projection declares exactly that set. A regrown project-year
    # cell must therefore be declared unlocked AND read Locked=False, on a keyed
    # row and on an unkeyed reserved row alike, while the permanent-id cell is
    # neither. The Locked property is read; nothing is inferred from a write,
    # because a COM value write is permitted on a locked cell.
    $costProtection = $null
    foreach ($entry in @($protection.sheets)) { if ([string]$entry.sheet -ceq $gridSheet) { $costProtection = $entry } }
    if ($null -eq $costProtection) { throw ('the protection projection declares no ' + $gridSheet + ' sheet') }
    $declaredUnlocked = @(@($costProtection.unlocked) | ForEach-Object { [string]$_ })
    $unkeyedRowIndex = 0
    for ($r = 0; $r -lt $growthAfter.Count; $r++) { if ([string]$growthAfter[$r][0] -eq '') { $unkeyedRowIndex = $r + 1; break } }
    if ($unkeyedRowIndex -lt 1) { throw 'the cost profiling grid holds no unkeyed reserved row to inspect' }
    $lockProbes = @(
        [pscustomobject]@{ Label = 'regrown keyed weight';            Row = 1;                Column = ($fixedColumns + $durationYears); ExpectLocked = $false },
        [pscustomobject]@{ Label = 'existing keyed weight';           Row = 1;                Column = ($fixedColumns + 1);              ExpectLocked = $false },
        [pscustomobject]@{ Label = 'regrown unkeyed reserved weight'; Row = $unkeyedRowIndex; Column = ($fixedColumns + $durationYears); ExpectLocked = $false },
        [pscustomobject]@{ Label = 'permanent id';                    Row = 1;                Column = 1;                                ExpectLocked = $true })
    $lockProblems = @()
    $lockDetails = @()
    foreach ($probe in $lockProbes) {
        $state = Get-FaCellLockState -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $probe.Row -ColumnIndex $probe.Column
        $declared = ($declaredUnlocked -ccontains $state.Address)
        $lockDetails += ($probe.Label + ' ' + $state.Address + ' Locked=' + [string]$state.Locked + ' declared-unlocked=' + [string]$declared)
        if ($state.Locked -ne $probe.ExpectLocked) { $lockProblems += ($probe.Label + ' ' + $state.Address + ' reads Locked=' + [string]$state.Locked + ', expected ' + [string]$probe.ExpectLocked) }
        if ($declared -eq $probe.ExpectLocked) { $lockProblems += ($probe.Label + ' ' + $state.Address + ' is ' + $(if ($declared) { 'declared unlocked' } else { 'not declared unlocked' }) + ' in the protection projection, expected ' + $(if ($probe.ExpectLocked) { 'locked' } else { 'unlocked' })) }
    }
    $null = Add-FaCheck 'repair.width-growth.lock-state' ($lockProblems.Count -eq 0) `
        $(if ($lockProblems.Count -eq 0) { ($lockDetails -join '; ') } else { ($lockProblems -join '; ') + '. Read: ' + ($lockDetails -join '; ') })
    # THE ORIGINAL FIXTURE, PUT BACK EXACTLY: every weight of every cost row,
    # blanks as blanks and never as a zero, proved by both digests. The risk grid
    # was never written and must still equal its baseline. A restored blank is a
    # ClearContents, which an external COM client may only issue inside the window.
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.width-growth.restore'
    try { Restore-FaCostRows -Original $originalCostBody }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.width-growth.restore' }
    $costAfterGrowthRestore = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $riskAfterGrowthRestore = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
    $null = Add-FaCheck 'repair.width-growth.restored' (($costAfterGrowthRestore -ceq $baselineCost) -and ($riskAfterGrowthRestore -ceq $baselineRisk)) `
        'both profiling grids byte-identical to before the width-growth fixture'
    # THE SHRINK AND SEMANTIC SCENARIOS' OWN FIXTURE: row one [0.5, 0.5, 0, ..., 0],
    # 100% with typed zeros, so only the gate each scenario names can refuse.
    Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 1 -Weight 0.5
    Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 2 -Weight 0.5
    for ($y = 3; $y -le $durationYears; $y++) { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year $y -Weight 0.0 }

    # (b) BLANK-ONLY SHRINK: a fifth project-year column added inside the window,
    # left blank; Repair trims it and changes nothing else.
    Save-FaGridsBefore
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-blank'
    try { $null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-blank' }
    $shrunkBlank = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $costAfterShrink = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $null = Add-FaCheck 'repair.shrink-blank' (($shrunkBlank -like 'OK|*') -and ($costAfterShrink -ceq $script:FaCostBefore)) `
        ($shrunkBlank + '; the blank project year was trimmed and the grid is byte-identical to before it was added')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-shrink'

    # (c) TYPED-ZERO SHRINK REFUSED: the extra column holds 0 on row one; the
    # row still totals 100%, so only the trim gate can refuse - and must, naming
    # the id and the project year.
    Save-FaGridsBefore
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-zero'
    try {
        $null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year ($durationYears + 1) -Weight 0.0
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-zero' }
    $costWithZero = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $script:FaCostBefore = $costWithZero
    $zeroRefused = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    Assert-FaGridsUnchanged -Scenario 'repair.shrink-zero-refused' -Announcement $zeroRefused
    $null = Add-FaCheck 'repair.shrink-zero-refused.names' `
        (($zeroRefused -like ('*' + $rowOneId + '*')) -and ($zeroRefused -like ('*project year ' + [string]($durationYears + 1) + '*')) -and ($zeroRefused -like '*typed zero counts as populated*')) `
        ('the refusal names ' + $rowOneId + ' and project year ' + [string]($durationYears + 1))
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-shrink-zero'
    # the zero is cleared inside the window (the column's cells are not declared inputs) and the blank column trimmed
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-zero-clear'
    try { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year ($durationYears + 1) -Weight $null }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-zero-clear' }
    $trimmedAgain = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    if ($trimmedAgain -notlike 'OK|*') { throw ('the blank column could not be trimmed after the zero was cleared: ' + $trimmedAgain) }

    # (d) NONZERO SHRINK REFUSED: row one becomes [0.15, 0.25, 0.25, 0.25 | 0.1]
    # so it still totals 100% and only the trim gate refuses.
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-nonzero'
    try {
        $null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 1 -Weight 0.4
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year ($durationYears + 1) -Weight 0.1
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-nonzero' }
    Save-FaGridsBefore
    $nonzeroRefused = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    Assert-FaGridsUnchanged -Scenario 'repair.shrink-nonzero-refused' -Announcement $nonzeroRefused
    $null = Add-FaCheck 'repair.shrink-nonzero-refused.names' `
        (($nonzeroRefused -like ('*' + $rowOneId + '*')) -and ($nonzeroRefused -like ('*project year ' + [string]($durationYears + 1) + '*'))) `
        ('the refusal names ' + $rowOneId + ' and project year ' + [string]($durationYears + 1))
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-shrink-nonzero'
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-nonzero-clear'
    try {
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year ($durationYears + 1) -Weight $null
        Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 1 -Weight 0.5
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.shrink-nonzero-clear' }
    $trimmedAgain = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    if ($trimmedAgain -notlike 'OK|*') { throw ('the blank column could not be trimmed after the weight was cleared: ' + $trimmedAgain) }

    # (e) FIXED-WIDTH POPULATED NON-100%: no width drift at all; row one's first
    # weight becomes 0.55 through its unlocked cell; Repair must refuse by the
    # semantic gate, naming the id, and change nothing.
    Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 1 -Weight 0.55
    Save-FaGridsBefore
    $non1Refused = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    Assert-FaGridsUnchanged -Scenario 'repair.semantic-non1-refused' -Announcement $non1Refused
    $null = Add-FaCheck 'repair.semantic-non1-refused.names' `
        (($non1Refused -like ('*' + $rowOneId + '*')) -and ($non1Refused -like '*populated and total*') -and ($non1Refused -like '*not 100%*')) `
        ('the refusal names ' + $rowOneId + ' as populated and not 100%')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-semantic'

    # (f) ALL-BLANK PROFILE: row one cleared entirely; an unmade assumption is
    # allowed through, and with nothing structural to repair the command is a no-op.
    # THE CLEARS GO THROUGH THE ACCEPTED WINDOW. A genuine blank is ClearContents,
    # and Benchmark Run 4 proved that an external COM client's ClearContents is
    # refused on a protected sheet even on a declared-unlocked cell, while its
    # value writes are permitted: the capability split belongs to the client, not
    # to the cell. Final acceptance run 9 stopped on exactly that at the first of
    # these clears - project year 1, a Stage-A unlocked cell. Every other clear in
    # this runner already sits inside the window; this one now does too. Repair
    # itself runs after the window closes, as everywhere else.
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.blank-profile'
    try { for ($y = 1; $y -le $durationYears; $y++) { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year $y -Weight $null } }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.blank-profile' }
    Save-FaGridsBefore
    $blankAllowed = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $costAfterBlank = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $null = Add-FaCheck 'repair.blank-profile-allowed' (($blankAllowed -like 'OK|*') -and ($blankAllowed -like '*Nothing was changed*') -and ($costAfterBlank -ceq $script:FaCostBefore)) `
        ($blankAllowed + '; the all-blank row passed as an unmade assumption and nothing moved')

    # (g) POPULATED ZERO-TOTAL SIGNED PROFILE: row one becomes [1, -1, 0, 0]; it
    # totals zero and must be refused as populated, never read as blank.
    Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 1 -Weight 1.0
    Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year 2 -Weight -1.0
    for ($y = 3; $y -le $durationYears; $y++) { Set-FaWeight -Workbook $wb -SheetName $gridSheet -TableName $gridTable -FixedColumns $fixedColumns -RowIndex 1 -Year $y -Weight 0.0 }
    Save-FaGridsBefore
    $signedRefused = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    Assert-FaGridsUnchanged -Scenario 'repair.signed-zero-total-refused' -Announcement $signedRefused
    $null = Add-FaCheck 'repair.signed-zero-total-refused.names' `
        (($signedRefused -like ('*' + $rowOneId + '*')) -and ($signedRefused -like '*populated and total 0*') -and ($signedRefused -like '*not 100%*')) `
        ('the refusal names ' + $rowOneId + ' as populated, total 0, not 100% - not as blank')
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-signed'
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.signed-zero-total.restore'
    try { Restore-FaRowOne }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.signed-zero-total.restore' }

    # (h) BOTH FAILPOINTS ROLL BACK: with row two deleted (a repairable fault),
    # the contracted failpoint after the cost grid and the one after the risk grid
    # each turn a repair into a FAIL whose rollback leaves BOTH grids byte-identical
    # to before the call; then the real repair recreates the row blank.
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.rollback'
    try { Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex 2 }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.rollback' }
    foreach ($failpoint in @($script:RepairFailpointCost, $script:RepairFailpointRisk)) {
        Save-FaGridsBefore
        $rolledBack = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling' -FailAfterStage $failpoint
        $costNow = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
        $riskNow = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
        $null = Add-FaCheck ('repair.rollback.' + $failpoint) `
            (($rolledBack -like 'FAIL|*') -and ($rolledBack -like ('*' + $failpoint + '*')) -and ($rolledBack -like '*restored to the state they were in before*') -and
             ($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)) `
            ($rolledBack + '; both profiling grids byte-identical to before the call')
        $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario ('protection.after-repair-rollback.' + $failpoint)
    }
    $repairedAfterRollback = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'
    $bodyAfterRollback = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $rowTwoIndex = Find-FaTableRow -Body $bodyAfterRollback -Id $missingId
    $rowTwoBlank = $true
    if ($rowTwoIndex -lt 1) { $rowTwoBlank = $false }
    else { $rowTwo = @($bodyAfterRollback[$rowTwoIndex - 1]); for ($c = $fixedColumns; $c -lt $rowTwo.Count; $c++) { if ([string]$rowTwo[$c] -ne '') { $rowTwoBlank = $false } } }
    $null = Add-FaCheck 'repair.rollback.then-repaired' (($repairedAfterRollback -like 'OK|*') -and $rowTwoBlank) `
        ($repairedAfterRollback + '; ' + $missingId + ' recreated blank after the rolled-back attempts')
    for ($c = $fixedColumns; $c -lt $missingRowBefore.Count; $c++) {
        $weight = [string]$missingRowBefore[$c]
        if ($weight -ne '') { Set-TableCell -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $rowTwoIndex -ColumnIndex ($c + 1) -Value ([double]$weight) }
    }
    # FIXTURE CLEANUP: THE PHYSICAL ROWS THE RUNNER DELETED, PUT BACK. The
    # rollback precondition deleted a profiling ListRow; production recreated the
    # driver in an existing blank row, so each grid may now hold fewer physical
    # body rows than the repair-contract baseline. Inside the accepted window the
    # runner appends exactly the blank rows it owes - no id, no trace text, no
    # weight - and refuses an excess rather than deleting anything.
    $baselineCostRows = @($baselineCostBody).Count
    $baselineRiskRows = @($baselineRiskBody).Count
    $capacityProblems = @()
    $capacityDetails = @()
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.grids-restored.capacity'
    try {
        foreach ($capacity in @(
                [pscustomobject]@{ Label = ($gridSheet + '!' + $gridTable); Sheet = $gridSheet; Table = $gridTable; Baseline = $baselineCostRows },
                [pscustomobject]@{ Label = ($riskSheet + '!' + $riskTable); Sheet = $riskSheet; Table = $riskTable; Baseline = $baselineRiskRows })) {
            $currentRows = Get-TableRowCount -Workbook $wb -SheetName $capacity.Sheet -TableName $capacity.Table
            $plan = Get-FaCapacityPlan -Current $currentRows -Baseline $capacity.Baseline
            if ($plan.Excess -gt 0) {
                $capacityProblems += ($capacity.Label + ' holds ' + [string]$currentRows + ' body rows where the baseline held ' + [string]$capacity.Baseline + '; an excess is not the runner''s to delete')
                continue
            }
            for ($i = 0; $i -lt $plan.Append; $i++) { $null = Add-BlankTableRow -Workbook $wb -SheetName $capacity.Sheet -TableName $capacity.Table }
            $afterAppend = @(Get-TableBody -Workbook $wb -SheetName $capacity.Sheet -TableName $capacity.Table)
            if ($afterAppend.Count -ne $capacity.Baseline) { $capacityProblems += ($capacity.Label + ' holds ' + [string]$afterAppend.Count + ' body rows after appending ' + [string]$plan.Append + ', expected ' + [string]$capacity.Baseline) }
            for ($r = $currentRows + 1; $r -le $afterAppend.Count; $r++) {
                $appendedRow = @($afterAppend[$r - 1])
                for ($c = 0; $c -lt $appendedRow.Count; $c++) {
                    if ([string]$appendedRow[$c] -ne '') { $capacityProblems += ($capacity.Label + ' appended row ' + [string]$r + ' column ' + [string]($c + 1) + ' holds ' + [string]$appendedRow[$c] + ', not blank') }
                }
            }
            $capacityDetails += ($capacity.Label + ' ' + [string]$currentRows + '->' + [string]$afterAppend.Count + ' rows (baseline ' + [string]$capacity.Baseline + ', appended ' + [string]$plan.Append + ' blank)')
        }
    } finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'repair.grids-restored.capacity' }
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-repair-capacity'
    $null = Add-FaCheck 'repair.grids-restored.capacity' ($capacityProblems.Count -eq 0) `
        $(if ($capacityProblems.Count -eq 0) { ($capacityDetails -join '; ') } else { ($capacityProblems -join '; ') })
    # DIAGNOSTIC SNAPSHOT A: both grids as plain data after the runner's own
    # restoration writes and BEFORE the final resync. Reads only.
    $preResyncCostBody = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $preResyncRiskBody = @(Get-TableBody -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $preResyncCostColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $preResyncRiskColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_ApplyTimeline' -Stage 'resync after the repair contract scenarios'
    $costRestored = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable
    $riskRestored = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable
    # DIAGNOSTIC SNAPSHOT B: both grids AFTER the final resync, the same state
    # the two digests above were read from. Reads only.
    $postResyncCostBody = @(Get-TableBody -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $postResyncRiskBody = @(Get-TableBody -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $postResyncCostColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $gridSheet -TableName $gridTable)
    $postResyncRiskColumns = @(Get-TableColumnNames -Workbook $wb -SheetName $riskSheet -TableName $riskTable)
    $gridsDiagnosis = @()
    $costDiagnosis = Format-FaGridDifferences -Label ($gridSheet + '!' + $gridTable) -Baseline $baselineCostBody -BeforeResync $preResyncCostBody -AfterResync $postResyncCostBody `
        -BaselineColumns $baselineCostColumns -BeforeColumns $preResyncCostColumns -AfterColumns $postResyncCostColumns
    if ($costDiagnosis -ne '') { $gridsDiagnosis += $costDiagnosis }
    $riskDiagnosis = Format-FaGridDifferences -Label ($riskSheet + '!' + $riskTable) -Baseline $baselineRiskBody -BeforeResync $preResyncRiskBody -AfterResync $postResyncRiskBody `
        -BaselineColumns $baselineRiskColumns -BeforeColumns $preResyncRiskColumns -AfterColumns $postResyncRiskColumns
    if ($riskDiagnosis -ne '') { $gridsDiagnosis += $riskDiagnosis }
    $null = Add-FaCheck 'repair.grids-restored' (($costRestored -ceq $baselineCost) -and ($riskRestored -ceq $baselineRisk)) `
        $(if ($gridsDiagnosis.Count -eq 0) { 'both profiling grids byte-identical to before the repair contract scenarios' } else { $gridsDiagnosis -join ' ~~ ' })
    # --- P10-R2 REPAIR CONTRACT SCENARIOS: end --------------------------------
    $recalc2 = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'calculate after the repairs')
    $fingerprintAfterRepairs = Get-FaRunText -Excel $excel -Procedure 'PCCM_CalculationFingerprint'
    $null = Add-FaCheck 'repair.fingerprint' (($recalc2 -like 'OK|*') -and ($fingerprintAfterRepairs -ceq $fingerprint)) `
        ('the calculation fingerprint after every repair equals the first CURRENT fingerprint: ' + $fingerprintAfterRepairs)

    # 15/16. RESET RESULTS, through the real entry point. All four publications are
    # re-established first so every owner has something to clear.
    $simResult2 = [string](Invoke-Phase6Simulation -Excel $excel)
    $sensResult2 = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RunSensitivity'
    $annualResult2 = Invoke-FaEndpoint -Excel $excel -Operation ([string]$p7.command_surface.annual_endpoint)
    $statesFull = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'reset.precondition' `
        (($simResult2 -like 'OK|*') -and ($sensResult2 -like 'OK|*') -and ($annualResult2 -like 'OK|*') -and
         ($statesFull.Calculation -ceq $statusCurrent) -and ($statesFull.Simulation -ceq $statusCurrent) -and
         ($statesFull.Annual -ceq $statusCurrent)) `
        ('every publication re-established; ' + (Format-FaStates $statesFull))
    $preservedBefore = Get-FaPreservedDigest -Workbook $wb -Reset $reset -Methodology $methodology -MetadataRange $metadataRange
    $publicationBefore = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    # THE _Calc TABLE GEOMETRY AT THE PRECONDITION: Reset owns contents, not shape.
    $calcTableShapes = Get-FaCalcTableShapes -Workbook $wb -Reset $reset
    $simStateBefore = Format-Phase6State -State (Get-Phase6State -Workbook $wb -Inspection $simInspect) -Label 'before'

    # 15a. DECLINED, deterministically: the automation seam answers the destructive
    # confirmation with False. No dialog exists, nothing is clicked. The prompt is
    # the one production recorded while the seam was live, captured with the
    # result before the seam was reset (run 13 read it afterwards, and it was gone).
    $declinedObservation = Invoke-FaObservedEndpoint -Excel $excel -Operation 'PCCM_ResetResults' -ConfirmReply $false
    $declined = [string]$declinedObservation.Result
    $prompt = [string]$declinedObservation.Prompt
    $publicationDeclined = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    $statesDeclined = Get-FaStates -Excel $excel
    $promptExcerpt = ($prompt -replace '[\r\n]+', ' ')
    if ($promptExcerpt.Length -gt 160) { $promptExcerpt = $promptExcerpt.Substring(0, 160) + '...' }
    if ($promptExcerpt -eq '') { $promptExcerpt = '<blank>' }
    $null = Add-FaCheck 'reset.declined' `
        (($declined -like 'OK|*') -and ($prompt -like '*Reset Results clears every published result*') -and
         ($publicationDeclined -ceq $publicationBefore) -and ($statesDeclined.Simulation -ceq $statusCurrent)) `
        ('endpoint=' + $declined + '; prompt=' + $promptExcerpt + '; publication unchanged=' + [string]($publicationDeclined -ceq $publicationBefore) +
         '; simulation=' + $statesDeclined.Simulation + '; ' + (Format-FaStates $statesDeclined))

    # 15b. CONFIRMED. Verified against the SEMANTIC post-reset state production
    # defines: every ordinary publication rectangle blank; the calculation state
    # block blank except the projected last-attempt field, which holds NONE; the
    # simulation publication record blank except its last-attempt field, which
    # holds NONE; every _Calc table body blank with its geometry unchanged from
    # the precondition. Each component is reported on its own.
    $confirmed = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_ResetResults'
    $resetProblems = @(Get-FaResetProblems -Workbook $wb -Reset $reset -Iterations $acceptanceIterations `
        -CalcAttemptCell $calcAttemptCell -SimAttemptCell $simAttemptCell -Sentinel $attemptNone -TableShapes $calcTableShapes)
    $attemptCell = Format-FaCell ((Get-FaBlock -Workbook $wb -SheetName $calcSheet -Address $calcAttemptCell).Rect)
    $calcStateActual = (ConvertTo-FaRectCells -Rect ((Get-FaBlock -Workbook $wb -SheetName $calcSheet -Address ([string]$reset.publications.calculation.cleared.state)).Rect)) -join ','
    $simRecordActual = (ConvertTo-FaRectCells -Rect ((Get-FaBlock -Workbook $wb -SheetName ([string]$reset.publications.simulation.sheet) -Address ([string]$reset.publications.simulation.cleared.attempt_and_selector)).Rect)) -join ','
    $null = Add-FaCheck 'reset.confirmed' `
        (($confirmed -like 'OK|Results reset. Model inputs and identity counters were preserved.') -and
         ($resetProblems.Count -eq 0) -and ($attemptCell -ceq $attemptNone)) `
        ('endpoint=' + $confirmed + '; calc state [' + $calcStateActual + ']; sim record [' + $simRecordActual + ']; ' +
         $(if ($resetProblems.Count -eq 0) { 'every ordinary publication rectangle blank, both sentinels NONE, every _Calc table body blank with its shape kept' } else { ($resetProblems -join '; ') }))
    $preservedAfter = Get-FaPreservedDigest -Workbook $wb -Reset $reset -Methodology $methodology -MetadataRange $metadataRange
    $null = Add-FaCheck 'reset.preserved' ($preservedAfter -ceq $preservedBefore) `
        'every declared editable input, the applied timeline, both permanent-id counters, the next AUTO nonce, the run-id, the pending nonce and the metadata block compare exactly'
    $statesReset = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'reset.states' `
        (($statesReset.Calculation -ceq $statusNotCalculated) -and ($statesReset.Simulation -eq '') -and
         ($statesReset.Annual -like 'NOT PRODUCED*')) `
        (Format-FaStates $statesReset)
    $excel.Calculate()
    $modelCheckReset = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckCalculationState'
    $null = Add-FaCheck 'modelcheck.after-reset.adapter' ($modelCheckReset -ceq $statusNotCalculated) ('adapter=' + $modelCheckReset)
    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.after-reset' `
        -Expected @($advisoryExpected, $notCalculatedExpected)
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-reset'

    # 15c. IDEMPOTENT.
    $publicationReset = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    $again = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_ResetResults'
    $publicationAgain = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    $preservedAgain = Get-FaPreservedDigest -Workbook $wb -Reset $reset -Methodology $methodology -MetadataRange $metadataRange
    $null = Add-FaCheck 'reset.idempotent' `
        (($again -ceq $confirmed) -and ($publicationAgain -ceq $publicationReset) -and ($preservedAgain -ceq $preservedBefore)) `
        ('second reset: ' + $again + '; nothing moved')

    # 12. REFUSED AS AN ATTEMPT OUTCOME: the annual endpoint with nothing to draw
    # from is refused, and no derived state changes.
    $refusedAnnual = Invoke-FaEndpoint -Excel $excel -Operation ([string]$p7.command_surface.annual_endpoint)
    $statesAfterRefusal = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'refused.outcome.annual' `
        (($refusedAnnual -like 'FAIL|*') -and ($statesAfterRefusal.Calculation -ceq $statusNotCalculated) -and
         ($statesAfterRefusal.Simulation -eq '') -and ($statesAfterRefusal.Annual -like 'NOT PRODUCED*')) `
        ($refusedAnnual + '; ' + (Format-FaStates $statesAfterRefusal))
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-refused-annual'

    # 15d. INJECTED FAILURE AND ROLLBACK, through the contracted failpoint.
    $calcResult3 = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'calculate before the rollback scenario')
    $simResult3 = [string](Invoke-Phase6Simulation -Excel $excel)
    $sensResult3 = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RunSensitivity'
    $annualResult3 = Invoke-FaEndpoint -Excel $excel -Operation ([string]$p7.command_surface.annual_endpoint)
    if (($calcResult3 -notlike 'OK|*') -or ($simResult3 -notlike 'OK|*') -or ($sensResult3 -notlike 'OK|*') -or ($annualResult3 -notlike 'OK|*')) {
        throw ('the rollback scenario could not re-establish its publications: ' + $calcResult3 + ' / ' + $simResult3 + ' / ' + $sensResult3 + ' / ' + $annualResult3)
    }
    $publicationFull = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    $simStateFull = Format-Phase6State -State (Get-Phase6State -Workbook $wb -Inspection $simInspect) -Label 'full'
    $preservedFull = Get-FaPreservedDigest -Workbook $wb -Reset $reset -Methodology $methodology -MetadataRange $metadataRange
    $injected = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_ResetResults' -FailAfterStage $script:ResetFailpoint
    $publicationRolledBack = Get-FaPublicationDigest -Workbook $wb -Reset $reset -Iterations $acceptanceIterations
    $simStateRolledBack = Format-Phase6State -State (Get-Phase6State -Workbook $wb -Inspection $simInspect) -Label 'full'
    $preservedRolledBack = Get-FaPreservedDigest -Workbook $wb -Reset $reset -Methodology $methodology -MetadataRange $metadataRange
    $statesRolledBack = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'reset.rollback' `
        (($injected -like 'FAIL|*') -and ($injected -like ('*' + $script:ResetFailpoint + '*')) -and
         ($injected -like '*put back*') -and ($publicationRolledBack -ceq $publicationFull) -and
         ($simStateRolledBack -ceq $simStateFull) -and ($preservedRolledBack -ceq $preservedFull) -and
         ($statesRolledBack.Calculation -ceq $statusCurrent) -and ($statesRolledBack.Simulation -ceq $statusCurrent) -and
         ($statesRolledBack.Annual -ceq $statusCurrent)) `
        ($injected + '; every publication rectangle, the persisted simulation state and every preserved cell compare exactly; ' + (Format-FaStates $statesRolledBack))
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-rollback'

    # 18/19. FINAL PROTECTION AND ITS BEHAVIOUR. One declared unlocked input cell
    # accepts its own value back. One locked reference cell is proved LOCKED on a
    # PROTECTED sheet - the user-edit protection - and then proved WRITABLE by
    # code, because the accepted design protects with UserInterfaceOnly:=True so
    # that production can write while a user cannot. Corrected at the bounded
    # correction round: the earlier runner expected the code write to be refused,
    # which is not what the accepted contract says, and would have read an
    # unrelated COM exception as a protection refusal. The target sheet and cell
    # are established explicitly first, and any exception on the way is a FAIL of
    # the check, never evidence of protection.
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.final'
    $setupProtection = $null
    foreach ($entry in @($protection.sheets)) { if ([string]$entry.sheet -ceq 'Setup') { $setupProtection = $entry } }
    if ($null -eq $setupProtection) { throw 'the protection projection declares no Setup sheet' }
    $unlockedAddress = ''
    $unlockedValue = $null
    foreach ($candidate in @($setupProtection.unlocked)) {
        $probe = (Get-FaBlock -Workbook $wb -SheetName 'Setup' -Address ([string]$candidate)).Rect
        if (($null -ne $probe) -and ($probe -isnot [array])) { $unlockedAddress = [string]$candidate; $unlockedValue = $probe; break }
    }
    if ($unlockedAddress -eq '') { throw 'no declared unlocked Setup input holds a value to write back' }
    $unlockedFailure = ''
    $worksheets = $null; $setupWs = $null; $inputCell = $null
    try {
        $worksheets = $wb.Worksheets
        $setupWs = $worksheets.Item('Setup')
        $inputCell = $setupWs.Range($unlockedAddress)
        if ($unlockedValue -is [string]) { $inputCell.Value2 = [string]$unlockedValue }
        else { $inputCell.Value2 = [double]$unlockedValue }
    } catch { $unlockedFailure = (Format-Err $_) }
    finally {
        if ($null -ne $inputCell)  { Release-Transient $inputCell  'Range(input)'; $inputCell  = $null }
        if ($null -ne $setupWs)    { Release-Transient $setupWs    'Worksheet';    $setupWs    = $null }
        if ($null -ne $worksheets) { Release-Transient $worksheets 'Worksheets';   $worksheets = $null }
    }
    $null = Add-FaCheck 'protection.unlocked-writable' ($unlockedFailure -eq '') `
        $(if ($unlockedFailure -eq '') { ('Setup!' + $unlockedAddress + ' accepted its own value under protection') } else { $unlockedFailure })
    # THE LOCKED CELL: the Source Revision label on the Methodology sheet, which
    # the protection projection declares protected with nothing unlocked.
    $methodologyProtection = $null
    foreach ($entry in @($protection.sheets)) { if ([string]$entry.sheet -ceq [string]$methodology.sheet) { $methodologyProtection = $entry } }
    if ($null -eq $methodologyProtection) { throw ('the protection projection declares no ' + [string]$methodology.sheet + ' sheet') }
    if (-not [bool]$methodologyProtection.protect) { throw ([string]$methodology.sheet + ' is not declared protected') }
    if (@($methodologyProtection.unlocked).Count -ne 0) { throw ([string]$methodology.sheet + ' declares unlocked cells; the locked-cell check needs a sheet with none') }
    $lockedAddress = [string]$methodology.label_column + [string]$sourceRevisionRow.row
    $lockedIsLocked = $false; $sheetIsProtected = $false
    $lockedReadFailure = ''
    $codeWriteFailure = ''
    $valueAfterWrite = ''
    $worksheets = $null; $methodWs = $null; $lockedCell = $null
    try {
        $worksheets = $wb.Worksheets
        $methodWs = $worksheets.Item([string]$methodology.sheet)
        $sheetIsProtected = [bool]$methodWs.ProtectContents
        $lockedCell = $methodWs.Range($lockedAddress)
        $lockedIsLocked = [bool]$lockedCell.Locked
    } catch { $lockedReadFailure = (Format-Err $_) }
    finally {
        if ($null -ne $lockedCell) { Release-Transient $lockedCell 'Range(locked)'; $lockedCell = $null }
        if ($null -ne $methodWs)   { Release-Transient $methodWs   'Worksheet';     $methodWs   = $null }
        if ($null -ne $worksheets) { Release-Transient $worksheets 'Worksheets';    $worksheets = $null }
    }
    $null = Add-FaCheck 'protection.locked-cell.user-protected' (($lockedReadFailure -eq '') -and $lockedIsLocked -and $sheetIsProtected) `
        $(if ($lockedReadFailure -eq '') { ([string]$methodology.sheet + '!' + $lockedAddress + ' Locked=' + [string]$lockedIsLocked + ' on a sheet with ProtectContents=' + [string]$sheetIsProtected) } else { $lockedReadFailure })
    $worksheets = $null; $methodWs = $null; $lockedCell = $null
    try {
        $worksheets = $wb.Worksheets
        $methodWs = $worksheets.Item([string]$methodology.sheet)
        $lockedCell = $methodWs.Range($lockedAddress)
        $lockedCell.Value2 = [string]$sourceRevisionRow.label
        $valueAfterWrite = [string]$lockedCell.Value2
    } catch { $codeWriteFailure = (Format-Err $_) }
    finally {
        if ($null -ne $lockedCell) { Release-Transient $lockedCell 'Range(locked)'; $lockedCell = $null }
        if ($null -ne $methodWs)   { Release-Transient $methodWs   'Worksheet';     $methodWs   = $null }
        if ($null -ne $worksheets) { Release-Transient $worksheets 'Worksheets';    $worksheets = $null }
    }
    $null = Add-FaCheck 'protection.locked-cell.code-write-permitted' (($codeWriteFailure -eq '') -and ($valueAfterWrite -ceq [string]$sourceRevisionRow.label)) `
        $(if ($codeWriteFailure -eq '') { ('code wrote the locked cell its own value under the accepted protection; user-edit protection and code-write capability are different facts') } else { $codeWriteFailure })
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'protection.after-locked-cell'

    $excel.Run('PCCM_AutomationEnd') | Out-Null
    # --- P10-R2 DISTRIBUTION COPY: begin ----------------------------------------
    # Contract matrix row J: a distribution copy runs from a different path and
    # name. The main workbook is closed unsaved; the BUILT file is copied to a
    # different folder under a different name, opened in the same owned Excel
    # instance - so Workbook_Open re-applies protection in the copy - and the
    # copy is proved to compile, carry the same Source Revision, report itself
    # protected, take the accepted fixture and run Calculate and a 1,000-iteration
    # Simulation to CURRENT. Closed unsaved like the original.
    try { $wb.Close($false); $rel.WorkbookClosed = $true }
    catch { $null = $rel.Failed.Add('Workbook.Close(original)') }
    Invoke-FaRelease -Ledger $rel -Obj $wb -Label 'Workbook(original)'
    $wb = $null
    $copyDir = Join-Path $tempRoot 'distribution copy'
    $null = New-Item -ItemType Directory -Path $copyDir -Force
    $copyPath = Join-Path $copyDir 'PCCM distribution copy.xlsm'
    Copy-Item -LiteralPath $stageBPath -Destination $copyPath -Force
    $wb = $workbooks.Open($copyPath)
    $comAcquired = $comAcquired + 1
    $null = Assert-FaWorkbookReady -Workbook $wb -ExpectedPath $copyPath -Session 'copy'
    $copyCompile = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $copyCompile = (Format-Err $_) }
    $null = Add-FaCheck 'copy.compile' ($copyCompile -eq '') $(if ($copyCompile -eq '') { ('opened from ' + $copyPath + ' and the VBAProject compiles') } else { $copyCompile })
    $copyRevision = Format-FaCell ((Get-FaBlock -Workbook $wb -SheetName ([string]$methodology.sheet) -Address ([string]$methodology.text_column + [string]$sourceRevisionRow.row)).Rect)
    $null = Add-FaCheck 'copy.source-revision' ($copyRevision -ceq $revision.Expected) ('expected ' + $revision.Expected + '; observed ' + $copyRevision)
    Import-FaFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'copy.protection'
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection
    $null = Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'copy.fixture'
    $copyApplied = ''
    try { $copyApplied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest -Inspection $inspection -Model $model) }
    finally { $null = Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'copy.fixture' }
    $null = Add-FaCheck 'copy.fixture' ($copyApplied -like 'OK|*') $copyApplied
    Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$acceptanceIterations)
    Set-NamedValue -Workbook $wb -DefinedName $seedName -Value $suppliedSeed
    $copyCalc = [string](Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' -Stage 'calculate in the distribution copy')
    $copySim = [string](Invoke-Phase6Simulation -Excel $excel)
    $copyStates = Get-FaStates -Excel $excel
    $null = Add-FaCheck 'copy.run' (($copyCalc -like 'OK|*') -and ($copySim -like 'OK|*') -and ($copyStates.Calculation -ceq $statusCurrent) -and ($copyStates.Simulation -ceq $statusCurrent)) `
        ($copyCalc + '; ' + $copySim + '; ' + (Format-FaStates $copyStates))
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'copy.protection-after-run'
    $excel.Run('PCCM_AutomationEnd') | Out-Null
    # --- P10-R2 DISTRIBUTION COPY: end ------------------------------------------
    # --- P10-R3 WORKBOOK_OPEN FAILURE: begin ------------------------------------
    # Contract matrix row O: an injected Workbook_Open failure leaves the
    # application state clean and the workbook usable. The handler runs before
    # any post-open automation can be armed, so the pre-open trigger is Excel's
    # own event switch: the copy is closed unsaved, a third disposable copy is
    # opened with EnableEvents off so the handler does NOT run at open (proved:
    # the file carries no protection and nothing is recorded), the accepted seam
    # is begun in the fresh project with the handler's own failpoint, the switch
    # is put back, and then the REAL handler - the production procedure in the
    # workbook's own document module, not a copy - is run through Application.Run.
    # Its failpoint fires after a successful apply, so the failure path is
    # exercised from a fully protected workbook. An exception from the run is a
    # FAIL of the check, never the failure path. The application state is read
    # before and compared after; nothing is written to restore it.
    try { $wb.Close($false); $rel.WorkbookClosed = $true }
    catch { $null = $rel.Failed.Add('Workbook.Close(copy)') }
    Invoke-FaRelease -Ledger $rel -Obj $wb -Label 'Workbook(copy)'
    $wb = $null
    $openDir = Join-Path $tempRoot 'open failure'
    $null = New-Item -ItemType Directory -Path $openDir -Force
    $openPath = Join-Path $openDir 'PCCM_open_failure_copy.xlsm'
    Copy-Item -LiteralPath $stageBPath -Destination $openPath -Force
    if (-not [bool]$excel.EnableEvents) { throw 'application events are already off before the open-failure session' }
    $suppressed = $null
    $resultBeforeHandler = '<not read>'
    $excel.EnableEvents = $false
    try {
        $wb = $workbooks.Open($openPath)
        $comAcquired = $comAcquired + 1
        $null = Assert-FaWorkbookReady -Workbook $wb -ExpectedPath $openPath -Session 'open-failure'
        Import-FaFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir
        $suppressed = Get-FaProtectionState -Excel $excel
        $resultBeforeHandler = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'
        $excel.Run('PCCM_AutomationBegin', $true, $script:OpenFailpoint) | Out-Null
    } finally { $excel.EnableEvents = $true }
    $null = Add-FaCheck 'rowo.open-suppressed' `
        (($null -ne $suppressed) -and (-not $suppressed.Applied) -and ($suppressed.Protected -eq 0) -and (-not $suppressed.Structure) -and ($suppressed.Depth -eq 0) -and ($resultBeforeHandler -eq '')) `
        ('opened with events off: ' + $(if ($null -eq $suppressed) { '<no state>' } else { $suppressed.Raw }) + '; recorded=' + $resultBeforeHandler)
    $screenBefore = [bool]$excel.ScreenUpdating
    $calcBefore = [int]$excel.Calculation
    $eventsBefore = [bool]$excel.EnableEvents
    $handlerFailure = ''
    try { $excel.Run("'" + [string]$wb.Name + "'!ThisWorkbook.Workbook_Open") | Out-Null }
    catch { $handlerFailure = (Format-Err $_) }
    $recorded = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'
    $prompted = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationPrompt'
    $expectedRecord = "Workbook_Open: Injected structural failure after stage '" + $script:OpenFailpoint + "'."
    $null = Add-FaCheck 'rowo.handler-entered' (($handlerFailure -eq '') -and ($recorded -clike 'Workbook_Open: *')) `
        $(if ($handlerFailure -eq '') { ('the real handler ran and recorded: ' + $recorded) } else { $handlerFailure })
    $null = Add-FaCheck 'rowo.failure-path' ($recorded -ceq $expectedRecord) `
        ('expected exactly ' + $expectedRecord + '; recorded ' + $recorded)
    $null = Add-FaCheck 'rowo.disclosed-once' (($recorded -ceq $expectedRecord) -and ($prompted -eq '') -and $eventsBefore) `
        'one record through the accepted mechanism, no prompt, no dialog: the unattended runner is still running'
    $null = Add-FaCheck 'rowo.screen-updating-restored' ([bool]$excel.ScreenUpdating -eq $screenBefore) `
        ('ScreenUpdating before ' + [string]$screenBefore + ', after ' + [string]$excel.ScreenUpdating)
    $null = Add-FaCheck 'rowo.calculation-unchanged' ([int]$excel.Calculation -eq $calcBefore) `
        ('Calculation before ' + [string]$calcBefore + ', after ' + [string][int]$excel.Calculation)
    $null = Add-FaCheck 'rowo.events-unchanged' ([bool]$excel.EnableEvents -eq $eventsBefore) `
        ('EnableEvents before ' + [string]$eventsBefore + ', after ' + [string]$excel.EnableEvents)
    $released = Get-FaProtectionState -Excel $excel
    $null = Add-FaCheck 'rowo.released-not-half-protected' `
        ((-not $released.Applied) -and ($released.Protected -eq 0) -and (-not $released.Structure) -and ($released.Depth -eq 0)) `
        ('after the failed open: ' + $released.Raw)
    $usableFailure = ''
    $valueAfterOpenFailure = ''
    $worksheets = $null; $methodWs = $null; $usableCell = $null
    try {
        $worksheets = $wb.Worksheets
        $methodWs = $worksheets.Item([string]$methodology.sheet)
        $usableCell = $methodWs.Range($lockedAddress)
        $usableCell.Value2 = [string]$sourceRevisionRow.label
        $valueAfterOpenFailure = [string]$usableCell.Value2
    } catch { $usableFailure = (Format-Err $_) }
    finally {
        if ($null -ne $usableCell)  { Release-Transient $usableCell  'Range(usable)'; $usableCell  = $null }
        if ($null -ne $methodWs)    { Release-Transient $methodWs    'Worksheet';     $methodWs    = $null }
        if ($null -ne $worksheets)  { Release-Transient $worksheets  'Worksheets';    $worksheets  = $null }
    }
    $null = Add-FaCheck 'rowo.usable' (($usableFailure -eq '') -and ($valueAfterOpenFailure -ceq [string]$sourceRevisionRow.label)) `
        $(if ($usableFailure -eq '') { ([string]$methodology.sheet + '!' + $lockedAddress + ' accepted its own value with no protection in force') } else { $usableFailure })
    # THE CONTRACT'S LAST SENTENCE: protection is attempted again the next time
    # the handler runs. The seam stays begun - a real failure would record, not
    # deadlock - with no failpoint, and the same handler applies protection.
    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $reopenFailure = ''
    try { $excel.Run("'" + [string]$wb.Name + "'!ThisWorkbook.Workbook_Open") | Out-Null }
    catch { $reopenFailure = (Format-Err $_) }
    $recordedAfterReopen = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'
    $null = Add-FaCheck 'rowo.reopen-silent' (($reopenFailure -eq '') -and ($recordedAfterReopen -eq '')) `
        $(if ($reopenFailure -eq '') { ('the handler ran again and recorded nothing; ScreenUpdating ' + [string]$excel.ScreenUpdating) } else { $reopenFailure })
    $null = Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'rowo.reopen-applied'
    $excel.Run('PCCM_AutomationEnd') | Out-Null
    # --- P10-R3 WORKBOOK_OPEN FAILURE: end --------------------------------------
} catch {
    $fatal = (Format-Err $_)
    Write-FaLine ''
    Write-FaLine ('FATAL|' + $fatal)
    foreach ($fatalLine in @(Format-FaFatalLines -ErrorRecord $_)) { Write-FaLine ([string]$fatalLine) }
} finally {
    # --- shutdown, the accepted path, leaf before parent --------------------
    if ($null -ne $wb) {
        # NEVER SAVED. The disposable copy is discarded.
        try { $wb.Close($false); $rel.WorkbookClosed = $true }
        catch { $null = $rel.Failed.Add('Workbook.Close') }
        Invoke-FaRelease -Ledger $rel -Obj $wb -Label 'Workbook'
        $wb = $null
    }
    if ($null -ne $workbooks) {
        Invoke-FaRelease -Ledger $rel -Obj $workbooks -Label 'Workbooks'
        $workbooks = $null
    }
    if ($null -ne $excel) {
        try { $excel.Quit(); $rel.QuitCalled = $true }
        catch { $null = $rel.Failed.Add('Application.Quit') }
        Invoke-FaRelease -Ledger $rel -Obj $excel -Label 'Application'
        $excel = $null
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        $naturalExit = (Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90)
    }
    if (($null -ne $excelIdentity) -and (-not $naturalExit)) {
        $emergencyRequired = $true
        $null = Invoke-EmergencyExcelCleanup -Identity $excelIdentity
    }
    $rel.NaturalExit = $naturalExit
    $rel.EmergencyRequired = $emergencyRequired
    Write-FaLine ''
    foreach ($line in @(Format-ReleaseLedger -Ledger $rel)) { Write-FaLine ([string]$line) }
    foreach ($transient in @(Get-TransientFailures)) {
        Write-FaLine ('      transient release FAILED: ' + [string]$transient)
    }
}

# 20. CLEAN SHUTDOWN - reported together, after the session.
$null = Add-FaCheck 'shutdown.workbook-close' $rel.WorkbookClosed 'Workbook.Close' -Continue
$null = Add-FaCheck 'shutdown.application-quit' $rel.QuitCalled 'Application.Quit' -Continue
$null = Add-FaCheck 'shutdown.natural-exit' $rel.NaturalExit 'the owned Excel process exited naturally' -Continue
$null = Add-FaCheck 'shutdown.no-emergency' (-not $rel.EmergencyRequired) 'no emergency cleanup was required' -Continue
$null = Add-FaCheck 'shutdown.com-released' `
    (([int]$rel.Attempted -eq $comAcquired) -and ($rel.Failed.Count -eq 0) -and (@($script:FaResidual).Count -eq 0) -and
     (@(Get-TransientFailures).Count -eq 0)) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released, ' + [string]$rel.Failed.Count +
     ' failed, ' + [string]@($script:FaResidual).Count + ' residual, ' + [string]@(Get-TransientFailures).Count +
     ' transient failure(s)') -Continue

$failed = @($script:FaChecks | Where-Object { -not $_.Ok })
$ok = (($failed.Count -eq 0) -and [string]::IsNullOrWhiteSpace($fatal))
Write-FaLine ''
Write-FaLine ('checks                   : ' + [string]$script:FaChecks.Count + ' recorded, ' + [string]$failed.Count + ' failed')
Write-FaLine ('source revision          : expected ' + $revision.Expected + '; observed ' +
              $(if ($null -eq $observedRevision) { '<not read>' } else { $observedRevision }))
Write-FaLine ('FINAL ACCEPTANCE ' + $(if ($ok) { 'PASS' } else { 'FAIL' }))
Write-FaLine ('report                   : ' + $script:FaPath)
Write-Host ('The report is at ' + $script:FaPath) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
