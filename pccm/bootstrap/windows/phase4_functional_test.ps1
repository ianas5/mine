<#
.SYNOPSIS
    PCCM Phase-4 Windows functional test harness. Exercises the structural runtime
    against a real Excel instance and produces a human-readable report.

.DESCRIPTION
    Gate-A source review is approved. Gate-B run 1 completed the whole Stage-B
    bootstrap on the target machine -- both Excel instances shut down naturally --
    and then threw in the bootstrap's final reporting block, so scenarios B onward
    have NOT yet run. That defect is fixed; see docs/phase4_gate_b_run1.md.

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
      PRE Collection-shape preflight, pure PowerShell, BEFORE Excel is started:
          the table-row emission contract must give 0/1/N rows with row boundaries
          intact, or the run aborts without creating an Excel process
      A   Stage-B build: .xlsm, FileFormat 52, 14 CodeNames, modules, five buttons,
          natural COM shutdown
      B   Permanent Cost Line IDs: sequence and non-reuse after deletion
      B2  A REAL ListObject reorder: identity travels with its own row data
      C   Permanent Risk IDs, independently sequenced
      D0  Seed a REAL, KEYED Inflation Profile, so the timeline scenarios below
          assert rate preservation against a named profile rather than an unkeyed row
      D   First timeline application
      E   Duration increase
      F   Start-year shift with unchanged duration, and BLANK preservation
      G   Duration decrease: destructive confirmation, cancelled then accepted
      H   Base-Year movement, earlier and later
      I   Combined three-way change
      J   Degenerate inflation span
      K   Profiling synchronisation by permanent ID
      K2  Profiling PERCENTAGES belong to the ID, not the row: with a LIVE timeline,
          distinct values are seeded per ID, the register is physically reordered by
          a real sort, synchronisation is re-run, and every value is checked against
          the identifier it was seeded against. B2 cannot cover this: it runs before
          the first Apply, when no project-year column exists yet.
      L   Apply failure after mutation: logical restore
      M   Growth beyond the 25 reserved rows, with presentation and Data Validation
          proved on the runtime-created row
      N   An Inflation Profile removed from Config: destructive, cancelled then
          accepted, with the timeline unchanged throughout
      O   Non-numeric content in a removed profiling cell counts as data loss
      P   Oversized pasted timeline values rejected without a VBA overflow
      Q   Add failure after row mutation: rows, values, counter and the Phase-3
          input contract restored
      R   Delete failure after row mutation: the deleted ListRow recreated with its
          fills and Data Validation intact
      S   Application state RESTORED to its prior values, on success and on failure
      T   Unkeyed structural data refuses Add and Apply, in all three grids
      U   Counter integrity: an invalid counter refuses allocation, including the
          historical case where every identifier has been deleted
      V   Generated year cells equal the contract input_fill exactly; unkeyed rows
          are model-controlled
      W   The representation ceiling is EXHAUSTED VALID STATE: a counter at
          ID_COUNTER_MAX refuses allocation without overflowing counter + 1, while
          structural revalidation stays clean and Apply Timeline still succeeds

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
# The row-emission idiom
# ===========================================================================
# THE single place a table row is put on the pipeline. Get-TableBody calls it, and
# so does the preflight below -- deliberately the same mechanism, because a probe
# with its own copy of the idiom can pass while the real reader is broken.
#
# -NoEnumerate is the whole point: without it PowerShell unrolls the row array and
# emits its cells as separate objects, so the caller sees cells where it expects
# rows. With it, one row goes out as one object, whatever its length.
function Write-RowObject {
    param([object[]]$Row)
    Write-Output -NoEnumerate $Row
}

# The same loop shape Get-TableBody uses, over fabricated values instead of cells.
# It exists so the preflight can exercise the emission contract with no Excel and
# no COM at all.
function Write-FabricatedRows {
    param([int]$RowCount, [int]$CellCount)
    for ($r = 1; $r -le $RowCount; $r++) {
        $line = @()
        for ($c = 1; $c -le $CellCount; $c++) { $line += ("r{0}c{1}" -f $r, $c) }
        Write-RowObject $line
    }
}

# ===========================================================================
# Preflight: collection shape, BEFORE Excel is started
# ===========================================================================
# Linux cannot execute PowerShell, so the static tests can only read the source.
# The actual pipeline semantics -- how many objects a function emits, and whether a
# row survives as a row -- can only be observed here. Two wrong shapes have already
# shipped past a source review, so this runs first, costs nothing, and stops the
# run before a single Excel process is created if the contract is broken.
try {
    $list = New-Checklist

    $zero = @(Write-FabricatedRows -RowCount 0 -CellCount 3)
    $null = Add-Check $list 'zero rows: the caller collection is empty' `
        ($zero.Count -eq 0) ("Count " + $zero.Count)

    $one = @(Write-FabricatedRows -RowCount 1 -CellCount 3)
    $null = Add-Check $list 'one row: the caller collection holds exactly one row' `
        ($one.Count -eq 1) ("Count " + $one.Count)
    $oneOk = $false
    if ($one.Count -eq 1) {
        $oneOk = (@($one[0]).Count -eq 3)
        $null = Add-Check $list 'one row: that row still has its three cells' `
            $oneOk ("row Count " + @($one[0]).Count)
        $null = Add-Check $list 'one row: the cell values are unchanged' `
            ((@($one[0]) -join ',') -eq 'r1c1,r1c2,r1c3') (@($one[0]) -join ',')
    }

    $two = @(Write-FabricatedRows -RowCount 2 -CellCount 3)
    $null = Add-Check $list 'two rows: the caller collection holds exactly two rows' `
        ($two.Count -eq 2) ("Count " + $two.Count)
    if ($two.Count -eq 2) {
        $null = Add-Check $list 'two rows: each element is one row, boundaries preserved' `
            ((@($two[0]) -join ',') -eq 'r1c1,r1c2,r1c3' -and `
             (@($two[1]) -join ',') -eq 'r2c1,r2c2,r2c3') `
            ("[0] " + (@($two[0]) -join ',') + " / [1] " + (@($two[1]) -join ','))
    }

    # The failure mode this replaced: `return ,$rows` would make Count 1 here, with
    # element 0 being the whole collection. Named so a regression is unmistakable.
    $null = Add-Check $list 'the whole collection is never emitted as one object' `
        ($two.Count -ne 1) "the producer emitted the entire table as a single object"

    $preflightOk = Test-ChecklistOk $list
    Add-Result 'PRE' 'Collection shape preflight (pure PowerShell, no Excel)' `
        $(if ($preflightOk) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    if (-not $preflightOk) {
        Write-Host ''
        Write-Host 'PHASE-4 FUNCTIONAL TEST ABORTED before Excel was started:' -ForegroundColor Red
        Write-Host 'the table-row emission contract is wrong, so every scenario would' -ForegroundColor Red
        Write-Host 'compare the wrong shape. Fix Write-RowObject / Get-TableBody first.' -ForegroundColor Red
        exit 1
    }
} catch {
    Add-Result 'PRE' 'Collection shape preflight' 'FAIL' (Format-Err $_)
    Write-Host ''
    Write-Host 'PHASE-4 FUNCTIONAL TEST ABORTED before Excel was started.' -ForegroundColor Red
    exit 1
}

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

function Set-NamedValueText {
    param($Workbook, [string]$DefinedName, [string]$Text)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $rng.Value2 = $Text
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

# Reads the data body of a table as a sequence of row arrays, so a later comparison
# is a plain value comparison and never a live COM read.
#
# ONE PIPELINE OBJECT PER TABLE ROW. That is the producer contract, and it is the
# only shape that gives the caller stable zero/one/many semantics:
#
#   0 rows -> 0 objects emitted -> caller's @(...) has Count 0
#   1 row  -> 1 row-array object -> Count 1, and [0] is that row
#   N rows -> N row-array objects -> Count N, each element one row
#
# Two wrong shapes were tried before this one, and both are named here because each
# looks correct in isolation:
#
#   `return $rows`   PowerShell enumerates a collection on output, so ONE row is
#                    emitted as its own cells and the caller sees N rows of one
#                    cell each; zero rows emits nothing and lands $null.
#   `return ,$rows`  the unary comma emits the WHOLE jagged array as a SINGLE
#                    object, so the caller's @(...) wraps that one object and ends
#                    up one level too deep: Count 1, and [0] is the entire table.
#                    Every `foreach ($row in ...)` then sees the table, not a row.
#
# Emitting each row through Write-RowObject is what makes the caller's @(...) do
# exactly one job -- normalising 0/1/N -- rather than trying to repair nesting.
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
    $sortField = $null
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
        # SortFields.Add returns a SortField. Assigning it to $null still mints the
        # RCW and then discards ownership of it, which is the same defect as a bare
        # discarded return: it is captured, released once and nulled.
        $sortField = $sortFields.Add($keyRange, 0, $Order)
        Release-Transient $sortField 'SortField'; $sortField = $null
        $sortObj.Apply()
        $sortFields.Clear()
    } finally {
        if ($null -ne $sortField)       { Release-Transient $sortField       'SortField';   $sortField       = $null }
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

# Asserts the Phase-3 input contract on ONE driver row: fills and Data Validation.
# Used after growth and after every rollback, because a row that comes back with the
# right values but no validation and no input language is restored in name only.
function Add-DriverRowContractChecks {
    param($List, $Workbook, $Register, [int]$RowIndex, $Manifest, [string]$Label)
    $idFill = Get-TableCellFill -Workbook $Workbook -SheetName $Register.sheet `
        -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex 1
    $userFill = Get-TableCellFill -Workbook $Workbook -SheetName $Register.sheet `
        -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex 2
    $null = Add-Check $List "$Label ID cell keeps the model-controlled fill" `
        ($idFill -eq $Manifest.presentation.locked_fill) ("fill " + $idFill)
    $null = Add-Check $List "$Label user cells keep the editable input fill" `
        ($userFill -eq $Manifest.presentation.input_fill) ("fill " + $userFill)
    $null = Add-Check $List "$Label ID cell carries NO user Data Validation" `
        (-not (Test-TableCellValidation -Workbook $Workbook -SheetName $Register.sheet `
            -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex 1))

    $validated = $Manifest.driver_validation_columns.($Register.key)
    $ok = $true
    foreach ($key in $validated) {
        $columnIndex = [array]::IndexOf($Register.columns, $key) + 1
        if (-not (Test-TableCellValidation -Workbook $Workbook -SheetName $Register.sheet `
                -TableName $Register.table_name -RowIndex $RowIndex -ColumnIndex $columnIndex)) {
            $ok = $false
        }
    }
    $null = Add-Check $List "$Label every validated user column still has its Data Validation" $ok
}

# Number format and fill of one grid cell, so the snapshot's presentation claims can
# be checked after a rollback rather than taken on trust.
function Get-TableCellFormat {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        return [string]$cell.NumberFormat
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Get-TableColumnWidth {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$ColumnIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null
    $cols = $null; $col = $null; $range = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $col = $cols.Item($ColumnIndex)
        $range = $col.Range
        return [double]$range.ColumnWidth
    } finally {
        if ($null -ne $range)           { Release-Transient $range           'Range(col)';  $range           = $null }
        if ($null -ne $col)             { Release-Transient $col             'ListColumn';  $col             = $null }
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
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
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}

if ($buildOk) {
  $preExisting = @(Get-PreExistingExcelPids)
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
        $afterThree = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'three adds issue CL-001, CL-002, CL-003' `
            (($afterThree -join ',') -eq (($expected[0..2]) -join ',')) ("got " + ($afterThree -join ','))

        $excel.Run('PCCM_DeleteCostLineById', $expected[1]) | Out-Null
        $afterDelete = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'deleting CL-002 leaves CL-001 and CL-003' `
            (($afterDelete -join ',') -eq (($expected[0], $expected[2]) -join ',')) ("got " + ($afterDelete -join ','))

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $afterFourth = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
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
        $before = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
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

        $after = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
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
        $excel.Run('PCCM_DeleteCostLineById', @(Get-IdColumnValues -Workbook $wb -Info $costReg)[-1]) | Out-Null
        $profileIds = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
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
        $afterTwo = @(Get-IdColumnValues -Workbook $wb -Info $riskReg)
        $null = Add-Check $list 'two adds issue R-001 and R-002' `
            (($afterTwo -join ',') -eq (($expected[0], $expected[1]) -join ',')) ("got " + ($afterTwo -join ','))
        $null = Add-Check $list 'the risk sequence is independent of the cost sequence' `
            ($afterTwo[0] -eq 'R-001') 'four cost lines have been added; the risk sequence still starts at 1'

        $excel.Run('PCCM_DeleteRiskById', $expected[1]) | Out-Null
        $excel.Run('PCCM_AddRisk') | Out-Null
        $afterReadd = @(Get-IdColumnValues -Workbook $wb -Info $riskReg)
        $null = Add-Check $list 'R-002 is not reused after deletion' `
            (($afterReadd -join ',') -eq (($expected[0], $expected[2]) -join ',')) ("got " + ($afterReadd -join ','))

        Add-Result 'C' 'Permanent Risk IDs: independent sequence, non-reuse' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'C' 'Permanent Risk IDs' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # D0. Seed a REAL, KEYED Inflation Profile before the timeline scenarios
    # -------------------------------------------------------------------
    # Inflation ownership is keyed by (Profile Name, Calendar Year). Writing a rate
    # into a row whose Profile Name is blank produces nothing the model owns:
    # CountRateLosses skips it, SetYearColumns never captures it, SyncProfileRows
    # clears it. Without a genuine named profile the D-J scenarios could not prove
    # calendar-year preservation at all, and the destructive Base-Year step had no
    # real rate to threaten.
    $testProfile = 'P4 Timeline Test Profile'
    try {
        $list = New-Checklist
        Add-ConfigProfile -Workbook $wb -ProfileName $testProfile -RowIndex 1
        $names = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName 'Config' -TableName 'tblInflationProfiles')) {
            if ($row[0] -ne '') { $names += $row[0] }
        }
        $null = Add-Check $list 'the test profile is in the Config master' ($names -contains $testProfile)
        Add-Result 'D0' 'Seed a keyed Inflation Profile for the timeline scenarios' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'D0' 'Seed a keyed Inflation Profile' 'FAIL' (Format-Err $_)
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
            $costBefore = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
            $inflBefore = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
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
                $costBefore = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
            }
            # Seed DISTINCT rates against the named profile's actual calendar-year
            # headers, so preservation can be asserted per calendar year rather than
            # per column position.
            $inflHeadersBefore = @(Get-TableColumnNames -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
            $inflYearsBefore = @($inflHeadersBefore | Select-Object -Skip $fixedInfl)
            $profileRow = 0
            $rowIdx = 0
            foreach ($row in $inflBefore) { $rowIdx++; if ($row[0] -eq $testProfile) { $profileRow = $rowIdx } }

            $ratesBefore = @{}
            if ($profileRow -gt 0 -and $inflYearsBefore.Count -gt 0) {
                for ($y = 0; $y -lt $inflYearsBefore.Count; $y++) {
                    $rate = 0.01 + (0.001 * $y)
                    Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                        -RowIndex $profileRow -ColumnIndex ($fixedInfl + $y + 1) -Value $rate
                }
                $inflBefore = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
                for ($y = 0; $y -lt $inflYearsBefore.Count; $y++) {
                    # Plain data: (Profile Name, Calendar Year) -> value.
                    $ratesBefore[$inflYearsBefore[$y]] = $inflBefore[$profileRow - 1][$fixedInfl + $y]
                }
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

            $costAfter = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
            $inflAfter = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)

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

            # Inflation rates survive by CALENDAR YEAR, never by column index. The
            # comparison is against the captured (Profile Name, Calendar Year) map.
            if ((-not $step.expect_rejected) -and $step.confirm -and $ratesBefore.Count -gt 0) {
                $afterRow = 0; $rowIdx = 0
                foreach ($row in $inflAfter) { $rowIdx++; if ($row[0] -eq $testProfile) { $afterRow = $rowIdx } }
                $null = Add-Check $list 'the named inflation profile still has its row' ($afterRow -gt 0)

                if ($afterRow -gt 0) {
                    $survivorsOk = $true
                    $survivorCount = 0
                    foreach ($year in $inflYears) {
                        $idx = [array]::IndexOf($inflYears, $year)
                        $actual = $inflAfter[$afterRow - 1][$fixedInfl + $idx]
                        if ($ratesBefore.ContainsKey($year)) {
                            # This calendar year survived the intersection.
                            $survivorCount++
                            if ($actual -ne $ratesBefore[$year]) { $survivorsOk = $false }
                        }
                    }
                    $null = Add-Check $list 'every surviving calendar year keeps EXACTLY its own rate' `
                        $survivorsOk ("checked $survivorCount surviving year(s) by calendar key")

                    $newBlank = $true
                    foreach ($year in @($step.transition.added_inflation_years)) {
                        $idx = [array]::IndexOf($inflYears, $year)
                        if ($idx -ge 0) {
                            if ($inflAfter[$afterRow - 1][$fixedInfl + $idx] -ne '') { $newBlank = $false }
                        }
                    }
                    $null = Add-Check $list 'newly required inflation years arrive BLANK, never zero' $newBlank

                    $goneOk = $true
                    foreach ($year in @($step.transition.removed_inflation_years)) {
                        if ($inflYears -contains $year) { $goneOk = $false }
                    }
                    $null = Add-Check $list 'calendar years leaving the span are gone from the headers' $goneOk
                }
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
        $costIds = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $profileIds = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'every identified cost line has exactly one profiling row' `
            (($costIds -join ',') -eq ($profileIds -join ',')) `
            ("register " + ($costIds -join ',') + " / grid " + ($profileIds -join ','))
        $null = Add-Check $list 'no profiling identifier appears twice' `
            ((@($profileIds | Select-Object -Unique)).Count -eq $profileIds.Count)

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $afterAdd = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $afterAdd += $row[0] }
        }
        $null = Add-Check $list 'adding a driver creates exactly one matching profiling row' `
            ($afterAdd.Count -eq ($profileIds.Count + 1)) ("grid rows " + $afterAdd.Count)

        $victim = $afterAdd[$afterAdd.Count - 1]
        $excel.Run('PCCM_DeleteCostLineById', $victim) | Out-Null
        $afterDelete = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
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
    # K2. Profiling PERCENTAGES survive a real reorder, with a LIVE timeline
    # -------------------------------------------------------------------
    # B2 performs a real ListObject.Sort, but it runs BEFORE the first timeline is
    # applied, when the profiling grids have no project-year columns at all. It
    # therefore proves that identity travels with row data and that profiling rows
    # stay keyed -- but it cannot prove the thing the permanent-ID design exists for:
    # that a PROFILED VALUE belongs to an IDENTIFIER, not to a worksheet row.
    #
    # This runs after D-J, so year columns exist. Distinct percentages are seeded per
    # ID, the register is physically reordered, the real synchronisation pathway is
    # re-run, and every value is checked against the ID it was seeded against.
    try {
        $list = New-Checklist

        # --- 1. at least two identified Cost Lines --------------------------
        while (@(Get-IdColumnValues -Workbook $wb -Info $costReg).Count -lt 2) {
            $excel.Run('PCCM_AddCostLine') | Out-Null
        }
        $ids = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'at least two identified Cost Lines exist' ($ids.Count -ge 2) `
            ($ids -join ',')

        # --- 2. at least one project-year profiling column -------------------
        $headers = @(Get-TableColumnNames -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $yearColumn = $fixedCost + 1
        $null = Add-Check $list 'the profiling grid has at least one project-year column' `
            ($headers.Count -gt $fixedCost) ("headers: " + ($headers -join '|'))

        if ($headers.Count -gt $fixedCost -and $ids.Count -ge 2) {
            # --- 3. a DISTINCT percentage per ID, same project-year index ----
            # Generated, not typed in: only the ID-to-value mapping matters.
            $seeded = @{}
            $step = 0
            foreach ($id in $ids) {
                $step++
                $seeded[$id] = [math]::Round(0.01 + ($step * 0.07), 4)
            }
            $gridRow = 0
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                $gridRow++
                if ($row[0] -ne '' -and $seeded.ContainsKey($row[0])) {
                    Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                        -RowIndex $gridRow -ColumnIndex $yearColumn -Value $seeded[$row[0]]
                }
            }

            # --- 4. capture ID -> percentage as PLAIN DATA -------------------
            # Read back rather than trusting the write, and hold no live COM object:
            # every later comparison is a value comparison.
            $before = @{}
            $positionalBefore = @()
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                if ($row[0] -ne '') {
                    $before[$row[0]] = [string]$row[$yearColumn - 1]
                    $positionalBefore += [string]$row[$yearColumn - 1]
                }
            }
            $distinct = @($before.Values | Select-Object -Unique)
            $null = Add-Check $list 'every seeded percentage is distinct, so a swap cannot pass' `
                ($distinct.Count -eq $before.Count) (($before.Values | Sort-Object) -join ',')

            $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
            $orderBefore = @(Get-IdColumnValues -Workbook $wb -Info $costReg)

            # --- 5. a REAL ListObject.Sort that INVERTS the physical order ---
            # Markers are written so that an ascending sort reverses the register.
            # Nothing is edited in place: whole rows move.
            $rank = $orderBefore.Count
            $registerRow = 0
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
                $registerRow++
                if ($row[0] -ne '') {
                    Set-TableCell -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
                        -RowIndex $registerRow -ColumnIndex 3 -Value ('ORDER-' + $rank.ToString('D4'))
                    $rank--
                }
            }
            Invoke-TableSort -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
                -KeyColumnIndex 3 -Order 1

            # --- 6. the REAL synchronisation pathway, values untouched -------
            # Apply re-runs SetYearColumns and SyncRows against the reordered
            # register. No entered timeline value is changed, so any movement of a
            # percentage is synchronisation behaviour, not a timeline edit.
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $outcome = [string]$excel.Run('PCCM_AutomationResult')
            $null = Add-Check $list 'the synchronisation pathway ran successfully' `
                ($outcome -like 'OK|*') $outcome

            # --- 7. the physical order really changed ------------------------
            $orderAfter = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
            $null = Add-Check $list 'the register row order actually changed' `
                (($orderBefore -join ',') -ne ($orderAfter -join ',')) `
                ("before " + ($orderBefore -join ',') + " / after " + ($orderAfter -join ','))

            # --- 8. every original identifier still exists -------------------
            $missing = @()
            foreach ($id in $orderBefore) { if ($orderAfter -notcontains $id) { $missing += $id } }
            $null = Add-Check $list 'every original permanent ID still exists' `
                ($missing.Count -eq 0) ("missing " + ($missing -join ','))

            # --- 9. every percentage still belongs to ITS OWN ID -------------
            $after = @{}
            $positionalAfter = @()
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                if ($row[0] -ne '') {
                    $after[$row[0]] = [string]$row[$yearColumn - 1]
                    $positionalAfter += [string]$row[$yearColumn - 1]
                }
            }
            $wrong = @()
            foreach ($id in $orderBefore) {
                if ($after[$id] -ne $before[$id]) {
                    $wrong += ("{0}: seeded {1}, now {2}" -f $id, $before[$id], $after[$id])
                }
            }
            $null = Add-Check $list 'every profiling percentage still belongs to its own permanent ID' `
                ($wrong.Count -eq 0) ($wrong -join '; ')

            # --- 10. and did NOT stay behind with the row position -----------
            # If a value had followed worksheet position instead of identity, the
            # positional sequence would be unchanged while the ID order reversed.
            $null = Add-Check $list 'no percentage followed worksheet row position' `
                (($positionalBefore -join ',') -ne ($positionalAfter -join ',')) `
                ("positional before " + ($positionalBefore -join ',') + " / after " + ($positionalAfter -join ','))
            $null = Add-Check $list 'the profiling grid order follows the reordered register' `
                ((@($after.Keys | Sort-Object) -join ',') -eq (@($orderAfter | Sort-Object) -join ','))

            # --- 11. the reorder issued nothing ------------------------------
            $null = Add-Check $list 'the ID counter is unchanged by the reorder' `
                ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $counterBefore)

            # --- 12. and the workbook is structurally sound ------------------
            $report = [string]$excel.Run('PCCM_StructuralReport')
            $null = Add-Check $list 'structural revalidation is clean after the reorder and sync' `
                ([string]::IsNullOrWhiteSpace($report)) $report
        }

        Add-Result 'K2' 'Profiling percentages belong to the ID, not the row, across a real reorder' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'K2' 'Profiling percentage ownership across a reorder' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # L. Runtime failure containment
    # -------------------------------------------------------------------
    try {
        $list = New-Checklist
        $costBefore = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $inflBefore = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
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

        $costAfter = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $inflAfter = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
        $null = Add-Check $list 'the cost profiling grid was logically restored' `
            ((($costBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($costAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the inflation grid was logically restored' `
            ((($inflBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($inflAfter | ForEach-Object { $_ -join '|' }) -join ';'))

        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'the restored workbook still passes structural revalidation' ([string]::IsNullOrWhiteSpace($report)) $report

        # The snapshot claims to preserve number format, width and the input-language
        # treatment of every year cell. After a failed reshape, prove it did.
        if ($costAfter.Count -gt 0 -and $costAfter[0].Count -gt $fixedCost) {
            $keyedRow = 0; $rowIdx = 0
            foreach ($row in $costAfter) { $rowIdx++; if ($row[0] -ne '' -and $keyedRow -eq 0) { $keyedRow = $rowIdx } }
            if ($keyedRow -gt 0) {
                $null = Add-Check $list 'a restored profiling year cell keeps its number format' `
                    ((Get-TableCellFormat -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                        -RowIndex $keyedRow -ColumnIndex ($fixedCost + 1)) -eq $costGrid.year_number_format)
                $null = Add-Check $list 'a restored profiling year cell keeps a positive column width' `
                    ((Get-TableColumnWidth -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                        -ColumnIndex ($fixedCost + 1)) -gt 0)
                $null = Add-Check $list 'a restored profiling year cell on a KEYED row is editable-input styled' `
                    ((Get-TableCellFill -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                        -RowIndex $keyedRow -ColumnIndex ($fixedCost + 1)) -eq $manifest.presentation.input_fill)
            }
        }

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
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'the grown driver has exactly one profiling row' `
            (@($profileIds | Where-Object { $_ -eq $newId }).Count -eq 1)

        # Presentation and validation on the RUNTIME-created row. If Excel Table
        # propagation is what supplies these, this is where that reliance is proved
        # rather than assumed.
        $grownRow = 0
        $rowIdx = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
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
        if (@(Get-TableColumnNames -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name).Count -gt $fixedCost) {
            $gridRow = 0; $rowIdx = 0
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
                $rowIdx++
                if ($row[0] -eq $newId) { $gridRow = $rowIdx }
            }
            if ($gridRow -gt 0) {
                $yearFill = Get-TableCellFill -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name -RowIndex $gridRow -ColumnIndex ($fixedCost + 1)
                $null = Add-Check $list 'a generated profiling year cell equals the contract input_fill' `
                    ($yearFill -eq $manifest.presentation.input_fill) `
                    ("expected " + $manifest.presentation.input_fill + ", got " + $yearFill)
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

        $inflRows = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
        $targetRow = 0; $rowIdx = 0
        foreach ($row in $inflRows) { $rowIdx++; if ($row[0] -eq $profileName) { $targetRow = $rowIdx } }
        $null = Add-Check $list 'the new Config profile gained an inflation row' ($targetRow -gt 0)

        if ($targetRow -gt 0 -and $inflRows[0].Count -gt $fixedInfl) {
            Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                -RowIndex $targetRow -ColumnIndex ($fixedInfl + 1) -Value 0.042
            $inflBefore = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)

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

            $inflAfterCancel = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
            $null = Add-Check $list 'cancelling leaves the inflation row and all its rates unchanged' `
                ((($inflBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($inflAfterCancel | ForEach-Object { $_ -join '|' }) -join ';'))

            # --- accepted --------------------------------------------------
            $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $remaining = @()
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)) {
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
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
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

            $stillThere = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
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
        $registerBefore = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
        $gridBefore = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
        $idsBefore = @(Get-IdColumnValues -Workbook $wb -Info $costReg)

        # Armed AFTER the identifier has been allocated and written, so the failure
        # lands with the register already mutated.
        $excel.Run('PCCM_AutomationBegin', $true, 'add.after_write_id') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected Add failure was reported' ($outcome -like 'FAIL|*') ("outcome '$outcome'")

        $registerAfter = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
        $gridAfter = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
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

        # The row that survived the rollback must still honour the Phase-3 contract.
        Add-DriverRowContractChecks -List $list -Workbook $wb -Register $costReg `
            -RowIndex 1 -Manifest $manifest -Label 'after Add rollback the restored row'

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
        $registerBefore = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
        $gridBefore = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
        $victim = @(Get-IdColumnValues -Workbook $wb -Info $costReg)[0]

        $excel.Run('PCCM_AutomationBegin', $true, 'delete.after_remove') | Out-Null
        $excel.Run('PCCM_DeleteCostLineById', $victim) | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected Delete failure was reported' ($outcome -like 'FAIL|*') ("outcome '$outcome'")

        $registerAfter = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
        $gridAfter = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
        $null = Add-Check $list 'the driver table row count was restored' ($registerAfter.Count -eq $registerBefore.Count) `
            ("before " + $registerBefore.Count + " / after " + $registerAfter.Count)
        $null = Add-Check $list 'the profiling row count was restored' ($gridAfter.Count -eq $gridBefore.Count)
        $null = Add-Check $list 'the deleted identifier is back' `
            (@(Get-IdColumnValues -Workbook $wb -Info $costReg) -contains $victim)
        $null = Add-Check $list 'the register values were restored' `
            ((($registerBefore | ForEach-Object { $_ -join '|' }) -join ';') -eq (($registerAfter | ForEach-Object { $_ -join '|' }) -join ';'))
        $null = Add-Check $list 'the ID counter was restored' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $counterBefore)

        # The recreated row matters most here: Delete removed a real ListRow and the
        # rollback had to put it back, formatting and validation included.
        $restoredRow = 0; $rowIdx = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
            $rowIdx++; if ($row[0] -eq $victim) { $restoredRow = $rowIdx }
        }
        if ($restoredRow -gt 0) {
            Add-DriverRowContractChecks -List $list -Workbook $wb -Register $costReg `
                -RowIndex $restoredRow -Manifest $manifest -Label 'after Delete rollback the recreated row'
        }

        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'the restored workbook still passes structural revalidation' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        Add-Result 'R' 'Delete failure after row mutation: full logical restore' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'R' 'Delete failure after row mutation' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # S. Application state is RESTORED, not forced to a convenient default
    # -------------------------------------------------------------------
    # The requirement is restoration of the caller's prior state. Asserting
    # ScreenUpdating = True proves nothing if the caller had it False to begin with.
    try {
        $list = New-Checklist

        # A deliberately unusual pre-operation state, none of it the default.
        $excel.ScreenUpdating = $false
        $excel.EnableEvents   = $false
        $excel.DisplayAlerts  = $false
        $excel.Calculation    = -4135          # xlCalculationManual
        $excel.StatusBar      = 'PCCM harness sentinel'

        $before = [pscustomobject]@{
            ScreenUpdating = [bool]$excel.ScreenUpdating
            EnableEvents   = [bool]$excel.EnableEvents
            DisplayAlerts  = [bool]$excel.DisplayAlerts
            Calculation    = [int]$excel.Calculation
            StatusBar      = [string]$excel.StatusBar
        }

        # --- a SUCCESSFUL structural operation ---------------------------
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the successful operation reported success' ($outcome -like 'OK|*') $outcome

        $null = Add-Check $list 'ScreenUpdating restored to its prior value' ([bool]$excel.ScreenUpdating -eq $before.ScreenUpdating)
        $null = Add-Check $list 'EnableEvents restored to its prior value'   ([bool]$excel.EnableEvents   -eq $before.EnableEvents)
        $null = Add-Check $list 'DisplayAlerts restored to its prior value'  ([bool]$excel.DisplayAlerts  -eq $before.DisplayAlerts)
        $null = Add-Check $list 'Calculation restored to its prior value'    ([int]$excel.Calculation     -eq $before.Calculation) `
            ("before " + $before.Calculation + " / after " + [int]$excel.Calculation)
        $null = Add-Check $list 'StatusBar restored to its prior value'      ([string]$excel.StatusBar    -eq $before.StatusBar) `
            ("read '" + [string]$excel.StatusBar + "'")

        # --- an INJECTED-FAILURE operation --------------------------------
        $excel.ScreenUpdating = $false
        $excel.EnableEvents   = $false
        $excel.StatusBar      = 'PCCM harness sentinel 2'
        $before2 = [pscustomobject]@{
            ScreenUpdating = [bool]$excel.ScreenUpdating
            EnableEvents   = [bool]$excel.EnableEvents
            DisplayAlerts  = [bool]$excel.DisplayAlerts
            Calculation    = [int]$excel.Calculation
            StatusBar      = [string]$excel.StatusBar
        }
        $excel.Run('PCCM_AutomationBegin', $true, 'add.after_write_id') | Out-Null
        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'the injected failure was reported' ($outcome -like 'FAIL|*') $outcome

        $null = Add-Check $list 'ScreenUpdating restored after failure' ([bool]$excel.ScreenUpdating -eq $before2.ScreenUpdating)
        $null = Add-Check $list 'EnableEvents restored after failure'   ([bool]$excel.EnableEvents   -eq $before2.EnableEvents)
        $null = Add-Check $list 'DisplayAlerts restored after failure'  ([bool]$excel.DisplayAlerts  -eq $before2.DisplayAlerts)
        $null = Add-Check $list 'Calculation restored after failure'    ([int]$excel.Calculation     -eq $before2.Calculation)
        $null = Add-Check $list 'StatusBar restored after failure'      ([string]$excel.StatusBar    -eq $before2.StatusBar)

        # Hand the workbook back in a normal state for the sections that follow.
        $excel.Run('PCCM_AutomationEnd') | Out-Null
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $excel.ScreenUpdating = $true
        $excel.EnableEvents   = $true
        $excel.DisplayAlerts  = $false
        $excel.Calculation    = -4105          # xlCalculationAutomatic
        $excel.StatusBar      = $false

        Add-Result 'S' 'Application state is restored to its prior values, on success and on failure' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'S' 'Application state restoration' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # T. Unkeyed structural data blocks every mutating operation
    # -------------------------------------------------------------------
    # An unkeyed row is invisible to every destructive assessment, because all of
    # them are keyed. Synchronisation would erase it with no warning, so the
    # operation is refused instead.
    try {
        $list = New-Checklist

        # --- T1: a driver row with content but no ID ----------------------
        $freeRow = 0
        $rowIdx = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
            $rowIdx++
            if ($row[0] -eq '' -and $freeRow -eq 0) { $freeRow = $rowIdx }
        }
        if ($freeRow -eq 0) {
            # Every reserved row is identified; grow one so an orphan can be placed.
            $excel.Run('PCCM_AddCostLine') | Out-Null
            $lastId = @(Get-IdColumnValues -Workbook $wb -Info $costReg)[-1]
            $excel.Run('PCCM_DeleteCostLineById', $lastId) | Out-Null
            $rowIdx = 0
            foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)) {
                $rowIdx++
                if ($row[0] -eq '' -and $freeRow -eq 0) { $freeRow = $rowIdx }
            }
        }
        $countBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine'
        Set-TableCell -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
            -RowIndex $freeRow -ColumnIndex 3 -Value 'ORPHAN DESCRIPTION'

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'Add is refused while a driver orphan exists' ($outcome -like 'FAIL|*') $outcome
        $null = Add-Check $list 'no identifier was allocated' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterCostLine') -eq $countBefore)
        $body = @(Get-TableBody -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name)
        $null = Add-Check $list 'the orphan row is untouched' ($body[$freeRow - 1][2] -eq 'ORPHAN DESCRIPTION')
        $null = Add-Check $list 'the orphan row still has no ID' ($body[$freeRow - 1][0] -eq '')

        Set-TableCell -Workbook $wb -SheetName $costReg.sheet -TableName $costReg.table_name `
            -RowIndex $freeRow -ColumnIndex 3 -Value $null

        # --- T2: a profiling row with a percentage but no ID --------------
        $gridFree = 0; $rowIdx = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            $rowIdx++
            if ($row[0] -eq '' -and $gridFree -eq 0) { $gridFree = $rowIdx }
        }
        if ($gridFree -gt 0 -and @(Get-TableColumnNames -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name).Count -gt $fixedCost) {
            Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                -RowIndex $gridFree -ColumnIndex ($fixedCost + 1) -Value 0.25
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $outcome = [string]$excel.Run('PCCM_AutomationResult')
            $null = Add-Check $list 'Apply is refused while a profiling orphan exists' ($outcome -like 'FAIL|*') $outcome
            $after = @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)
            $null = Add-Check $list 'the orphan profiling value is untouched' `
                ([double]$after[$gridFree - 1][$fixedCost] -eq 0.25)
            Set-TableCell -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name `
                -RowIndex $gridFree -ColumnIndex ($fixedCost + 1) -Value $null
        }

        # --- T3: an inflation row with a rate but no profile name ---------
        $inflFree = 0; $rowIdx = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)) {
            $rowIdx++
            if ($row[0] -eq '' -and $inflFree -eq 0) { $inflFree = $rowIdx }
        }
        if ($inflFree -gt 0 -and @(Get-TableColumnNames -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name).Count -gt $fixedInfl) {
            Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                -RowIndex $inflFree -ColumnIndex ($fixedInfl + 1) -Value 0.077
            $excel.Run('PCCM_ApplyTimeline') | Out-Null
            $outcome = [string]$excel.Run('PCCM_AutomationResult')
            $null = Add-Check $list 'Apply is refused while an inflation orphan exists' ($outcome -like 'FAIL|*') $outcome
            $after = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
            $null = Add-Check $list 'the orphan inflation rate is untouched' `
                ([double]$after[$inflFree - 1][$fixedInfl] -eq 0.077)
            Set-TableCell -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                -RowIndex $inflFree -ColumnIndex ($fixedInfl + 1) -Value $null
        }

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'once the orphans are cleared, Add succeeds again' ($outcome -like 'OK|*') $outcome

        Add-Result 'T' 'Unkeyed structural data refuses every mutating operation' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'T' 'Unkeyed structural data' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # U. A corrupt ID counter must never allow reuse
    # -------------------------------------------------------------------
    # The dangerous case is HISTORICAL, not "counter below a current row": if every
    # identified Risk is deleted and the counter is then corrupted, current rows say
    # nothing, and a silent fallback to zero would reissue R-001.
    try {
        $list = New-Checklist
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk'

        foreach ($id in @(Get-IdColumnValues -Workbook $wb -Info $riskReg)) {
            $excel.Run('PCCM_DeleteRiskById', $id) | Out-Null
        }
        $remaining = @(Get-IdColumnValues -Workbook $wb -Info $riskReg)
        $null = Add-Check $list 'every identified Risk was deleted' ($remaining.Count -eq 0)
        $null = Add-Check $list 'the counter survived the deletions' `
            ((Get-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk') -eq $counterBefore)

        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'a valid counter with zero rows is not a fault' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        # Corrupt the counter to text.
        Set-NamedValueText -Workbook $wb -DefinedName 'nmCounterRisk' -Text 'corrupt'
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'structural revalidation reports the invalid counter' `
            ($report -match 'counter_integrity') $report

        $excel.Run('PCCM_AddRisk') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'Add Risk is refused while the counter is invalid' ($outcome -like 'FAIL|*') $outcome
        $null = Add-Check $list 'no identifier was allocated' `
            (@(Get-IdColumnValues -Workbook $wb -Info $riskReg).Count -eq 0)
        $profileIds = @()
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $riskGrid.sheet -TableName $riskGrid.table_name)) {
            if ($row[0] -ne '') { $profileIds += $row[0] }
        }
        $null = Add-Check $list 'no profiling row was created' ($profileIds.Count -eq 0)
        # THE ROLLBACK MUST NOT LAUNDER CORRUPTION INTO ZERO. The failed Add rolls the
        # driver operation back, and the rollback restores the counter it snapshotted.
        # If that snapshot were taken through a lossy numeric accessor, the corrupt text
        # would come back as a VALID 0 -- and R-001 would be reissuable on the next Add.
        # The counter must still read as exactly the corrupt text that was written.
        $afterCorrupt = [string](Get-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk')
        $null = Add-Check $list 'the corrupt counter is STILL the same corrupt text after the failed Add' `
            ($afterCorrupt -ceq 'corrupt') ("counter now reads '" + $afterCorrupt + "'")
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'and structural revalidation still reports it as invalid' `
            ($report -match 'counter_integrity') $report

        # Blank is equally invalid, and equally must not read as zero.
        Set-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk' -Value $null
        $excel.Run('PCCM_AddRisk') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'a BLANK counter is refused too, never treated as zero' ($outcome -like 'FAIL|*') $outcome
        $afterBlank = [string](Get-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk')
        $null = Add-Check $list 'the blank counter is STILL blank after the failed Add, not restored as 0' `
            ($afterBlank -eq '') ("counter now reads '" + $afterBlank + "'")

        # Restore, and prove the sequence continues from history rather than restarting.
        Set-NamedValue -Workbook $wb -DefinedName 'nmCounterRisk' -Value ([double]$counterBefore)
        $excel.Run('PCCM_AddRisk') | Out-Null
        $issued = @(Get-IdColumnValues -Workbook $wb -Info $riskReg)
        $null = Add-Check $list 'the restored counter continues history and does not reuse R-001' `
            ($issued.Count -eq 1 -and $issued[0] -ne 'R-001') ("issued " + ($issued -join ','))

        Add-Result 'U' 'Counter integrity: an invalid counter refuses allocation and blocks reuse' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'U' 'Counter integrity' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # V. Generated year cells carry the EXACT editable-input treatment
    # -------------------------------------------------------------------
    # "not the locked fill" is not an assertion. Equality against the contract's
    # input_fill is.
    try {
        $list = New-Checklist
        $inputFill  = $manifest.presentation.input_fill
        $lockedFill = $manifest.presentation.locked_fill

        foreach ($grid in @($costGrid, $riskGrid)) {
            $headers = @(Get-TableColumnNames -Workbook $wb -SheetName $grid.sheet -TableName $grid.table_name)
            $fixed = $grid.fixed_columns.Count
            if ($headers.Count -gt $fixed) {
                $rows = @(Get-TableBody -Workbook $wb -SheetName $grid.sheet -TableName $grid.table_name)
                $keyed = 0; $unkeyed = 0; $rowIdx = 0
                foreach ($row in $rows) {
                    $rowIdx++
                    if ($row[0] -ne '' -and $keyed -eq 0) { $keyed = $rowIdx }
                    if ($row[0] -eq '' -and $unkeyed -eq 0) { $unkeyed = $rowIdx }
                }
                if ($keyed -gt 0) {
                    $fill = Get-TableCellFill -Workbook $wb -SheetName $grid.sheet -TableName $grid.table_name `
                        -RowIndex $keyed -ColumnIndex ($fixed + 1)
                    $null = Add-Check $list ($grid.table_name + ': a project-year cell on an IDENTIFIED row equals input_fill') `
                        ($fill -eq $inputFill) ("expected " + $inputFill + ", got " + $fill)
                }
                if ($unkeyed -gt 0) {
                    $fill = Get-TableCellFill -Workbook $wb -SheetName $grid.sheet -TableName $grid.table_name `
                        -RowIndex $unkeyed -ColumnIndex ($fixed + 1)
                    $null = Add-Check $list ($grid.table_name + ': a year cell on an UNKEYED reserved row is model-controlled') `
                        ($fill -eq $lockedFill) ("expected " + $lockedFill + ", got " + $fill)
                }
            }
        }

        $inflHeaders = @(Get-TableColumnNames -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
        if ($inflHeaders.Count -gt $fixedInfl) {
            $rows = @(Get-TableBody -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name)
            $named = 0; $unnamed = 0; $rowIdx = 0
            foreach ($row in $rows) {
                $rowIdx++
                if ($row[0] -ne '' -and $named -eq 0) { $named = $rowIdx }
                if ($row[0] -eq '' -and $unnamed -eq 0) { $unnamed = $rowIdx }
            }
            if ($named -gt 0) {
                $fill = Get-TableCellFill -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                    -RowIndex $named -ColumnIndex ($fixedInfl + 1)
                $null = Add-Check $list 'tblInflation: a calendar-year cell on a NAMED profile row equals input_fill' `
                    ($fill -eq $inputFill) ("expected " + $inputFill + ", got " + $fill)
            }
            if ($unnamed -gt 0) {
                $fill = Get-TableCellFill -Workbook $wb -SheetName $inflGrid.sheet -TableName $inflGrid.table_name `
                    -RowIndex $unnamed -ColumnIndex ($fixedInfl + 1)
                $null = Add-Check $list 'tblInflation: a year cell on an UNNAMED row is model-controlled' `
                    ($fill -eq $lockedFill) ("expected " + $lockedFill + ", got " + $fill)
            }
        }

        Add-Result 'V' 'Generated year cells carry the exact contract input-language treatment' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'V' 'Generated year cell presentation' 'FAIL' (Format-Err $_)
    }

    # -------------------------------------------------------------------
    # W. The representation ceiling is EXHAUSTED VALID STATE, not corruption
    # -------------------------------------------------------------------
    # Two things have to be true at once, and they pull in opposite directions:
    #
    #   * allocation must refuse, and must refuse BEFORE evaluating counter + 1,
    #     because ID_COUNTER_MAX + 1 overflows a VBA Long at runtime;
    #   * the workbook is NOT broken. A counter at the ceiling says the sequence is
    #     spent, not that the structure is incoherent -- so structural revalidation
    #     must stay clean, or Apply and Delete would be rolled back for every
    #     unrelated structural operation merely because the sequence ran out.
    try {
        $list = New-Checklist
        $ceiling = [double]$manifest.limits.id_counter_max
        $costCounter = $null
        foreach ($c in $manifest.counters) { if ($c.key -eq 'cost_line') { $costCounter = $c } }
        $counterName = [string]$costCounter.defined_name
        $counterBefore = Get-NamedValue -Workbook $wb -DefinedName $counterName

        # Composed from the manifest prefix and pad width, not typed in: the harness
        # restates no identifier format of its own.
        $digits = [string][long]$ceiling
        while ($digits.Length -lt [int]$costCounter.pad_width) { $digits = '0' + $digits }
        $ceilingId = [string]$costCounter.prefix + $digits

        # The ceiling identifier must not already be present, or a refusal could be
        # attributed to a duplicate rather than to the ceiling itself.
        $idsBefore = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'the ceiling identifier is not already in the register' `
            ($idsBefore -notcontains $ceilingId) $ceilingId

        Set-NamedValue -Workbook $wb -DefinedName $counterName -Value $ceiling
        $report = [string]$excel.Run('PCCM_StructuralReport')
        $null = Add-Check $list 'a counter AT the ceiling is not reported as a structural fault' `
            ([string]::IsNullOrWhiteSpace($report)) $report

        $profBefore = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profBefore++ }
        }

        $excel.Run('PCCM_AddCostLine') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'Add is refused cleanly at the ceiling, with no overflow' `
            ($outcome -like 'FAIL|*') $outcome
        $null = Add-Check $list 'the refusal names the ceiling as a representation limit' `
            ($outcome -match 'representation') $outcome

        $null = Add-Check $list 'the counter is unchanged: nothing was allocated' `
            ([double](Get-NamedValue -Workbook $wb -DefinedName $counterName) -eq $ceiling)
        $idsAfter = @(Get-IdColumnValues -Workbook $wb -Info $costReg)
        $null = Add-Check $list 'no register row was keyed' `
            ($idsAfter.Count -eq $idsBefore.Count) ("before " + $idsBefore.Count + ", after " + $idsAfter.Count)
        $null = Add-Check $list 'the ceiling identifier was never issued' `
            ($idsAfter -notcontains $ceilingId)
        $profAfter = 0
        foreach ($row in @(Get-TableBody -Workbook $wb -SheetName $costGrid.sheet -TableName $costGrid.table_name)) {
            if ($row[0] -ne '') { $profAfter++ }
        }
        $null = Add-Check $list 'no profiling row was created' ($profAfter -eq $profBefore)

        # An unrelated structural operation must still work while the counter is spent.
        $excel.Run('PCCM_ApplyTimeline') | Out-Null
        $outcome = [string]$excel.Run('PCCM_AutomationResult')
        $null = Add-Check $list 'Apply Timeline still succeeds while the sequence is exhausted' `
            ($outcome -like 'OK|*') $outcome

        Set-NamedValue -Workbook $wb -DefinedName $counterName -Value ([double]$counterBefore)
        $null = Add-Check $list 'the counter was restored for the remaining scenarios' `
            ([double](Get-NamedValue -Workbook $wb -DefinedName $counterName) -eq [double]$counterBefore)

        Add-Result 'W' 'Representation ceiling: refused allocation, valid structure' `
            $(if (Test-ChecklistOk $list) { 'PASS' } else { 'FAIL' }) (Format-Checklist $list)
    } catch {
        Add-Result 'W' 'Representation ceiling' 'FAIL' (Format-Err $_)
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
# @(...) at the caller: an empty collection returned from a function emits zero
# pipeline objects, so this lands $null, and Set-StrictMode turns .Count on $null
# into a PropertyNotFoundException. Same defect as build_stage_b.ps1, and it would
# have fired here too -- after every scenario had already passed.
$transient = @(Get-TransientFailures)
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
