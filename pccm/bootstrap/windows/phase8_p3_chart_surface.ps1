<#
.SYNOPSIS
    PCCM Phase 8 - the MINIMAL P8-3 runner: the analytical charts, live.

.DESCRIPTION
    P8-3 DREW FOUR CHARTS. This runs them.

    THE CLAIM UNDER TEST HAS TWO HALVES, and neither is arithmetic:

      THE OBJECTS ARE WHAT THE PROJECTION SAYS. Four chart objects on the
      Dashboard, of the projected types, at the projected anchors, each series
      and each category range pointed at the projected bridge cells and at
      nothing else. That is read out of Excel's own ChartObjects and
      SeriesCollection, not inferred from the file the builder wrote.

      AND WHAT THEY PLOT IS WHAT THE SHEET SAYS. Every plotted value is compared
      against a bridge cell frozen BEFORE the chart was inspected. There is no
      second histogram, no reconstructed cumulative series and no re-ranked
      tornado anywhere in this file: a runner that recomputed what it was
      checking would agree with itself about a chart that had drifted.

    THE ABSENCE IS AS IMPORTANT AS THE PRESENCE. Excel plots a zero-length
    string as ZERO, so the bridge emits NA() where there is no point. A series
    over the 200-row annual window must therefore come back as four numbers and
    196 #N/A - not as four numbers and 196 zeros, which is what a fabricated
    project costing nothing looks like. `Series.Values` returns the error code
    for #N/A, so this is checked directly rather than believed.

    SEVEN PARTS, ONE SESSION, ONE DISPOSABLE WORKBOOK:

      0   EMPTY      nothing produced: the charts exist, the bridge evaluates,
                     and every plotted point is #N/A rather than a zero
      A1  CURRENT    the accepted W4 fixture, calculate, simulate, annual at
                     P80 - and DELIBERATELY NO SENSITIVITY, so the tornado has
                     to say it has nothing rather than draw nothing convincingly
      A2  RANKED     the sensitivity endpoint invoked explicitly, as an
                     acceptance action; the tornado then plots the accepted
                     ranking in the accepted order with its signs intact
      B   SELECTOR   P80 -> P50, recalculation only, NO endpoint: the annual
                     charts keep the preserved P80 profile and say OTHER Px,
                     while the histogram stays CURRENT because a reporting
                     selector is not a simulation fingerprint input
      C   RERUN      the annual step alone at P50: the annual charts follow, the
                     histogram and the tornado do not move
      D   REQUEST    iterations 1000 -> 1001, recalculation only, NO endpoint:
                     the live simulation state goes STALE and every chart that
                     depends on it is qualified STALE - while the persisted
                     `(last evaluated)` row may still read CURRENT, which is
                     exactly why nothing here is qualified by that row
      E   INVALID    a cost line bound below its own minimum: the calculation
                     refuses, the live state reads INVALID, and the preserved
                     payload stays visible and stays qualified

    THE STATE ORACLE IS THE WORKSHEET, NOT THIS FILE. Not one state word is
    typed here: every expectation is read out of the accepted projections, and
    every comparison is against a frozen Results cell. There is no PowerShell
    reimplementation of CURRENT, STALE, INVALID, HISTORICAL or OTHER Px.

    WHAT THIS RUNNER IS NOT. It re-tests no Phase-7 mathematics, re-proves no
    P8-1 or P8-2 assertion, recomputes no percentile, no bin, no cumulative sum
    and no rank, and it starts no analysis to make a chart look better.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.NOTES
    WINDOWS POWERSHELL 5.1 is the target shell. `Join-Path a b c` - a child per
    positional argument - is PowerShell 6+ only, and the roots below are
    resolved the way the harnesses that have actually run on 5.1 resolve them.
#>

[CmdletBinding()]
param(
    [string]$BuildDir
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ALL THREE ARE DEFINITION-ONLY AT TOP LEVEL. Dot-sourcing them defines
# functions and script variables and runs no scenario.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# ===========================================================================
# THE TEN HELPERS COPIED FROM THE ACCEPTED PHASE-7 TIMING HARNESS
# ===========================================================================
# THE ACCEPTED PHASE-5 FIXTURE CHOREOGRAPHY CALLS THEM and its own file does not
# define them. Copied BYTE FOR BYTE from `phase7_timing_scenarios.ps1` - the
# same bytes P8-1 and P8-2 copied - so the tree carries one behaviour rather
# than four. A source control pins them to it.
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

function Get-IdColumnValues {
    param($Workbook, $Info)
    $out = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}
$script:P83Lines = New-Object System.Collections.ArrayList
$script:P83Path = ''
$script:P83Checks = New-Object System.Collections.ArrayList
$script:P83Residual = New-Object System.Collections.ArrayList

function Write-P83Line {
    param([string]$Text = '')
    $null = $script:P83Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:P83Path)) {
        try {
            Set-Content -LiteralPath $script:P83Path `
                -Value ($script:P83Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-P83Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:P83Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-P83Line $line
    return $Ok
}

function Invoke-P83Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:P83Residual.Add(
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

function Get-P83SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a P83 result could not be attributed to a source revision')
    }
    $dirty = @()
    foreach ($pathspec in @('pccm/src', 'pccm/spec', 'pccm/builder')) {
        $lines = @()
        try { $lines = @(& git -C $RepoRoot status --porcelain -- $pathspec 2>$null) } catch { $lines = @() }
        foreach ($line in $lines) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $dirty += [string]$line }
        }
    }
    return [pscustomobject]@{ Head = $head; Dirty = $dirty }
}


# THE ENDPOINT WRAPPER AND THE NAMED-TEXT WRITER, from the accepted P8-1
# runner. Both are TRANSITIONS - step 1 of the observation rule - and never
# appear inside an observation.
function Invoke-P83Endpoint {
    param($Excel, [string]$Endpoint)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run($Endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

function Set-P83NamedText {
    param($Workbook, [string]$DefinedName, [string]$Value)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $rng.Value2 = [string]$Value
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}


# THE REGISTER HELPERS, from the accepted P8-1 runner. Part E edits one cost
# line through the projected register rather than by address.
function Get-P83Register {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-P83RegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}

function Get-P83RegisterRowIndex {
    param($Workbook, $Register, [string]$PermanentId)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName ([string]$Register.sheet) `
        -TableName ([string]$Register.table_name))
    for ($index = 0; $index -lt $body.Count; $index++) {
        $cells = @($body[$index])
        if ($cells.Count -lt 1) { continue }
        if (([string]$cells[0]) -ceq $PermanentId) { return ($index + 1) }
    }
    return 0
}

# ===========================================================================
# READING A CELL ON EITHER SHEET, AND KNOWING WHEN IT DID NOT EVALUATE
# ===========================================================================
# THE ONE REAL DIFFERENCE FROM P8-1'S READER: a sheet name. P8-1 read one
# surface and could take the sheet from its projection; this runner exists to
# compare two, so the sheet is an argument and every call site says which one it
# means. The three facts per cell are the same three, and they are still not
# interchangeable:
#
#   Value2   the stored value. THE ONLY AUTHORITY for agreement. An Excel error
#            arrives as an Int32 in the CVErr band, which is why the band is
#            named below rather than inferred from how the text looks.
#   Formula  what the cell was built from. Agreement of values is not enough
#            here: two cells can hold the same number for different reasons, and
#            the claim under test is that one is a MIRROR of the other.
#   Text     what a reader sees. Captured for the report and never compared
#            against a number: `#,##0` turns 4.5E-13 into `0`, and a check that
#            read that string would call a mismatch agreement.
$script:P83ErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}

function Get-P83Cell {
    param($Workbook, [string]$SheetName, [string]$Address)
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($SheetName)
        $range = $sheet.Range($Address)
        $value = $range.Value2
        $text = ''
        try { $text = [string]$range.Text } catch { $text = '<unreadable>' }
        $formula = ''
        try { $formula = [string]$range.Formula } catch { $formula = '<unreadable>' }
        $format = ''
        try { $format = [string]$range.NumberFormat } catch { $format = '<unreadable>' }
        $errorName = ''
        if (($value -is [int]) -or ($value -is [long])) {
            $code = [int]$value
            if ($script:P83ErrorCodes.ContainsKey($code)) {
                $errorName = [string]$script:P83ErrorCodes[$code]
            }
        }
        # A CELL WHOSE TEXT IS AN ERROR NAME IS AN ERROR whatever its Value2
        # arrived as. Both routes are kept because neither is guaranteed alone.
        if ([string]::IsNullOrEmpty($errorName)) {
            foreach ($name in $script:P83ErrorCodes.Values) {
                if (([string]$text) -ceq [string]$name) { $errorName = [string]$name }
            }
        }
        return [pscustomobject]@{
            Sheet = $SheetName; Address = $Address; Value = $value; Text = $text
            Formula = $formula; NumberFormat = $format; ErrorName = $errorName
            IsError = (-not [string]::IsNullOrEmpty($errorName))
        }
    } finally {
        if ($null -ne $range)  { Release-Transient $range  ('Range(' + $SheetName + ')');      $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  ('Worksheet(' + $SheetName + ')');  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';                       $sheets = $null }
    }
}

function Format-P83Cell {
    param($Cell)
    if ($null -eq $Cell) { return '<not read>' }
    if ($Cell.IsError) { return ([string]$Cell.ErrorName + ' (error)') }
    return ([string]$Cell.Sheet + '!' + [string]$Cell.Address + ' = ' +
            (Format-SimValue $Cell.Value) + ' shown as ' + [char]39 + [string]$Cell.Text + [char]39)
}

function Test-P83Blank {
    param($Cell)
    if ($null -eq $Cell) { return $false }
    if ($Cell.IsError) { return $false }
    return (Test-SimBlank -Value $Cell.Value)
}

# EXACT AGREEMENT, AND BLANK COUNTS AS A VALUE. A mirror of an empty cell must
# be empty; `Test-SimSameValue` settles the ordinary cases and the two blanks are
# named first because "" and $null and a missing value all arrive here.
function Test-P83SameCellValue {
    param($Left, $Right)
    if (($null -eq $Left) -or ($null -eq $Right)) { return $false }
    if ($Left.IsError -or $Right.IsError) { return $false }
    $leftBlank = Test-SimBlank -Value $Left.Value
    $rightBlank = Test-SimBlank -Value $Right.Value
    if ($leftBlank -or $rightBlank) { return ($leftBlank -and $rightBlank) }
    return (Test-SimSameValue -A $Left.Value -B $Right.Value)
}

# ORDINARY RECALCULATION, AND WHAT THE APPLICATION WAS SET TO WHEN IT HAPPENED.
# The endpoints put calculation into manual and restore it; a mirror that failed
# to update because the workbook was left in manual is a different finding from
# one that failed because the formula is wrong, and the report must be able to
# tell them apart.
function Invoke-P83Recalculate {
    param($Excel, [string]$Stage)
    $mode = '<unreadable>'
    try { $mode = [string]$Excel.Calculation } catch { $mode = '<unreadable>' }
    $failure = ''
    try { $Excel.Calculate() } catch { $failure = (Format-Err $_) }
    Write-P83Line ('    recalculated at ' + $Stage + ' (Application.Calculation = ' + $mode + ')')
    return (Add-P83Check ($Stage + ': the workbook recalculated') `
        ([string]::IsNullOrWhiteSpace($failure)) ($failure + ' calculation mode ' + $mode) `
        'PREREQUISITE')
}


# ===========================================================================
# THE PROJECTION, READ THE ONE WAY IT IS MEANT TO BE READ
# ===========================================================================
# NOT ONE CHART IDENTITY, ANCHOR, TYPE OR RANGE IS SPELLED IN THIS FILE.
# `phase8_charts_inspection.json` carries all of it, and a runner that typed any
# of it would be a second declaration of something the manifest owns - which is
# the P7-4 failure mode, where a hand-written address went stale in silence and
# reported "Not produced for this run" forever.

# EXCEL'S OWN CHART TYPE NUMBERS, and only the three this project permits. A
# third dimension carries no data here and distorts the comparison a chart
# exists to make, so the 3-D members are named as things to REFUSE rather than
# left out and hoped about.
$script:P83ChartTypes = @{ 'line' = 4; 'column' = 51; 'bar' = 57 }
$script:P83ThreeD = @{
    -4100 = 'xl3DColumn'; -4101 = 'xl3DLine'; -4102 = 'xl3DPie'
    54 = 'xl3DColumnClustered'; 55 = 'xl3DColumnStacked'; 56 = 'xl3DColumnStacked100'
    60 = 'xl3DBarClustered';    61 = 'xl3DBarStacked';    62 = 'xl3DBarStacked100'
    -4151 = 'xlSurface'; 21 = 'xl3DPieExploded'; 70 = 'xl3DArea'
}

# A COLUMN LETTER AS A NUMBER, copied from the accepted P8-1 runner. The
# projection spells an anchor as `B65`; Excel reports it as a column index.
function ConvertTo-P83ColumnNumber {
    param([string]$Letters)
    $number = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - 64)
    }
    return $number
}

function ConvertFrom-P83ColumnNumber {
    param([int]$Number)
    $letters = ''
    $remaining = $Number
    while ($remaining -gt 0) {
        $remainder = ($remaining - 1) % 26
        $letters = ([string][char](65 + $remainder)) + $letters
        $remaining = [int](($remaining - $remainder - 1) / 26)
    }
    return $letters
}

# POINTS TO CENTIMETRES. Excel sizes a shape in points; the manifest declares
# centimetres because that is what a person laying out a page thinks in.
function ConvertTo-P83Centimetres {
    param([double]$Points)
    return ($Points / 72.0 * 2.54)
}

# ===========================================================================
# READING A CHART OUT OF EXCEL
# ===========================================================================
# EVERY FACT BELOW COMES FROM THE LIVE OBJECT, not from the file the builder
# wrote. `ChartObjects` is what Excel actually loaded; `SeriesCollection` is what
# it actually plots; and `Series.Formula` is the SERIES() expression Excel
# itself resolved, which is the only thing that can prove a range survived the
# round trip through the .xlsm.
function Get-P83SeriesParts {
    param([string]$Formula)
    # =SERIES(name, categories, values, order). Split at the TOP level only: a
    # range can carry no comma, but a quoted name can, and a naive split would
    # tear one in half.
    $inner = $Formula
    $open = $inner.IndexOf('(')
    if ($open -ge 0) { $inner = $inner.Substring($open + 1) }
    if ($inner.EndsWith(')')) { $inner = $inner.Substring(0, $inner.Length - 1) }
    $parts = New-Object System.Collections.ArrayList
    $depth = 0
    $quoted = $false
    $token = ''
    foreach ($character in $inner.ToCharArray()) {
        if ($character -eq '"') { $quoted = -not $quoted; $token = $token + [string]$character; continue }
        if (-not $quoted) {
            if ($character -eq '(') { $depth = $depth + 1 }
            elseif ($character -eq ')') { $depth = $depth - 1 }
            elseif (($character -eq ',') -and ($depth -eq 0)) {
                $null = $parts.Add($token); $token = ''; continue
            }
        }
        $token = $token + [string]$character
    }
    $null = $parts.Add($token)
    $all = @($parts)
    $name = ''; $categories = ''; $values = ''
    if ($all.Count -ge 1) { $name = ([string]$all[0]).Trim() }
    if ($all.Count -ge 2) { $categories = ([string]$all[1]).Trim() }
    if ($all.Count -ge 3) { $values = ([string]$all[2]).Trim() }
    return [pscustomobject]@{ Name = $name; Categories = $categories; Values = $values }
}

# A RANGE AS EXCEL SPELLS IT VERSUS AS THE PROJECTION SPELLS IT. Excel quotes a
# sheet name when it feels like it and always writes the dollars; the projection
# writes `Results!$H$279:$H$478`. Normalising both to the same shape is the only
# honest way to compare them, and it is done by removing what carries no meaning
# rather than by matching loosely.
function ConvertTo-P83NormalRange {
    param([string]$Reference)
    $out = [string]$Reference
    $out = $out.Replace("'", '')
    $out = $out.Replace('=', '')
    $out = $out.Trim()
    return $out
}

function Get-P83Charts {
    param($Workbook, [string]$SheetName)
    $sheets = $null; $sheet = $null; $objects = $null
    $out = New-Object System.Collections.ArrayList
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($SheetName)
        $objects = $sheet.ChartObjects()
        $count = [int]$objects.Count
        for ($index = 1; $index -le $count; $index++) {
            $object = $null; $chart = $null; $series = $null; $anchor = $null
            try {
                $object = $objects.Item($index)
                $chart = $object.Chart
                $title = ''
                try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = '<unreadable>' }
                $chartType = -1
                try { $chartType = [int]$chart.ChartType } catch { $chartType = -1 }
                $anchor = $object.TopLeftCell
                $anchorRow = [int]$anchor.Row
                $anchorColumn = [int]$anchor.Column
                $widthCm = ConvertTo-P83Centimetres -Points ([double]$object.Width)
                $heightCm = ConvertTo-P83Centimetres -Points ([double]$object.Height)
                $hasLegend = $false
                try { $hasLegend = [bool]$chart.HasLegend } catch { $hasLegend = $false }
                $series = $chart.SeriesCollection()
                $seriesCount = [int]$series.Count
                $collected = New-Object System.Collections.ArrayList
                for ($s = 1; $s -le $seriesCount; $s++) {
                    $item = $null
                    try {
                        $item = $series.Item($s)
                        $formula = [string]$item.Formula
                        $parts = Get-P83SeriesParts -Formula $formula
                        # THE PLOTTED VALUES THEMSELVES. A #N/A arrives as the
                        # CVErr code, which is how "no point" is told apart from
                        # "a point at zero" - the whole reason the bridge exists.
                        $values = @()
                        try { $values = @($item.Values) } catch { $values = @() }
                        $null = $collected.Add([pscustomobject]@{
                            Index = $s; Formula = $formula; Name = $parts.Name
                            Categories = (ConvertTo-P83NormalRange -Reference $parts.Categories)
                            ValuesRange = (ConvertTo-P83NormalRange -Reference $parts.Values)
                            Values = $values
                        })
                    } finally {
                        if ($null -ne $item) { Release-Transient $item 'Series'; $item = $null }
                    }
                }
                $null = $out.Add([pscustomobject]@{
                    Index = $index; Title = $title; ChartType = $chartType
                    AnchorRow = $anchorRow; AnchorColumn = $anchorColumn
                    WidthCm = $widthCm; HeightCm = $heightCm; HasLegend = $hasLegend
                    SeriesCount = $seriesCount; Series = @($collected)
                })
            } finally {
                if ($null -ne $anchor)  { Release-Transient $anchor  'TopLeftCell';       $anchor  = $null }
                if ($null -ne $series)  { Release-Transient $series  'SeriesCollection';  $series  = $null }
                if ($null -ne $chart)   { Release-Transient $chart   'Chart';             $chart   = $null }
                if ($null -ne $object)  { Release-Transient $object  'ChartObject';       $object  = $null }
            }
        }
    } finally {
        if ($null -ne $objects) { Release-Transient $objects 'ChartObjects'; $objects = $null }
        if ($null -ne $sheet)   { Release-Transient $sheet   'Worksheet';    $sheet   = $null }
        if ($null -ne $sheets)  { Release-Transient $sheets  'Worksheets';   $sheets  = $null }
    }
    return @($out)
}

# THE PLOTTED POINT, CLASSIFIED. Three outcomes and they are not the same thing:
# a number, a deliberate no-point, or something else entirely.
function Test-P83NoPoint {
    param($Value)
    if ($null -eq $Value) { return $true }
    if (($Value -is [int]) -or ($Value -is [long])) {
        return ($script:P83ErrorCodes.ContainsKey([int]$Value))
    }
    return $false
}

function Format-P83Point {
    param($Value)
    if ($null -eq $Value) { return '<empty>' }
    if (Test-P83NoPoint -Value $Value) {
        $code = 0
        if (($Value -is [int]) -or ($Value -is [long])) { $code = [int]$Value }
        if ($script:P83ErrorCodes.ContainsKey($code)) { return [string]$script:P83ErrorCodes[$code] }
        return '<no point>'
    }
    return (Format-SimValue $Value)
}

# ===========================================================================
# STEPS 3 AND 4 - FREEZING WHAT THE CHART WILL BE COMPARED AGAINST
# ===========================================================================
# THE ORDER IS THE MEASUREMENT. The authoritative state cells and the bridge
# cells are read BEFORE any chart object is touched, so what a series is
# compared against is a value that existed before anything asked Excel to draw.

# THE STATE CELLS, FROM THE TWO ACCEPTED SURFACES. The four annual lines are
# P8-1's, on Results; the two chart-status lines are P8-3's, in the bridge. Not
# one of them is derived here.
function Get-P83StateFrozen {
    param($Workbook, $P8, $Charts)
    $sheet = [string]$P8.sheet
    $column = [string]$P8.columns.nominal
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $P8.state.PSObject.Properties.Name) {
        $out.Add([string]$key, (Get-P83Cell -Workbook $Workbook -SheetName $sheet `
            -Address ($column + [string]([int]$P8.state.$key.row))))
    }
    foreach ($entry in @($Charts.bridge.status.rows)) {
        $out.Add([string]$entry.key, (Get-P83Cell -Workbook $Workbook `
            -SheetName ([string]$Charts.bridge_sheet) `
            -Address ($column + [string]([int]$entry.row))))
    }
    # THE PERSISTED ROW IS FROZEN TOO, and never used as an expectation. It is
    # here so the report can show what it said at the moment a chart was
    # qualified by the LIVE state instead - which is the whole point of the
    # pre-Windows correction, and is unarguable only if both are recorded.
    foreach ($field in @($P8.run_stamp.fields)) {
        if (([string]$field.key -ceq 'simulation_status') -or
            ([string]$field.key -ceq 'iterations_run') -or
            ([string]$field.key -ceq 'run_id')) {
            $out.Add(('run_stamp.' + [string]$field.key), (Get-P83Cell -Workbook $Workbook `
                -SheetName $sheet -Address ($column + [string]([int]$field.row))))
        }
    }
    return $out
}

# EVERY BRIDGE CELL A CHART CAN PLOT, frozen in one pass. The ranges come from
# the projection; this walks them rather than naming a single address.
function Get-P83BridgeFrozen {
    param($Workbook, $Charts)
    $sheet = [string]$Charts.bridge_sheet
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($name in @('annual', 'distribution', 'drivers')) {
        $block = $Charts.bridge.$name
        $first = [int]$block.first_row
        $last = [int]$block.last_row
        foreach ($column in @($block.columns)) {
            $letter = [string]$column.column
            $values = New-Object System.Collections.ArrayList
            for ($row = $first; $row -le $last; $row++) {
                $null = $values.Add((Get-P83Cell -Workbook $Workbook -SheetName $sheet `
                    -Address ($letter + [string]$row)))
            }
            $out.Add(($name + '.' + [string]$column.key), @($values))
        }
    }
    return $out
}

# ===========================================================================
# STEPS 5 AND 6 - THE CHART OBJECTS, AND THE COMPARISON
# ===========================================================================
# WHAT IS ASSERTED PER CHART, and each fails separately so a failure can be
# attributed:
#
#   identity   the projected chart exists, once, with the projected title
#   type       Excel's own ChartType is the projected one, and is not 3-D
#   geometry   the anchor is the projected cell and the size is within a
#              tolerance that allows for Excel rounding a shape to whole pixels
#   wiring     every series formula names the projected bridge range, and the
#              category range is the projected one - read out of SERIES(), so a
#              range that did not survive the round trip is caught
#   provenance no series names the machine sheets
#   payload    every plotted point equals the bridge cell frozen before the
#              chart was read, INCLUDING the no-points: a #N/A must stay a #N/A
#              and must never arrive as a zero
function Invoke-P83ChartChecks {
    param($Charts, $Frozen, $Bridge, [string]$Stage)
    $expected = @($Charts.charts)
    $missing = New-Object System.Collections.ArrayList
    $wrongType = New-Object System.Collections.ArrayList
    $threeD = New-Object System.Collections.ArrayList
    $wrongGeometry = New-Object System.Collections.ArrayList
    $wrongWiring = New-Object System.Collections.ArrayList
    $machine = New-Object System.Collections.ArrayList
    $wrongPayload = New-Object System.Collections.ArrayList
    $fabricated = New-Object System.Collections.ArrayList

    $byTitle = @{}
    foreach ($chart in @($Frozen)) {
        if (-not $byTitle.ContainsKey([string]$chart.Title)) {
            $byTitle[[string]$chart.Title] = $chart
        }
    }

    foreach ($spec in $expected) {
        $title = [string]$spec.title
        if (-not $byTitle.ContainsKey($title)) {
            $null = $missing.Add([string]$spec.key + ' (' + $title + ')')
            continue
        }
        $actual = $byTitle[$title]

        $wantType = [int]$script:P83ChartTypes[[string]$spec.kind]
        if ([int]$actual.ChartType -ne $wantType) {
            $null = $wrongType.Add($spec.key + ': ChartType ' + [string]$actual.ChartType +
                                   ', expected ' + [string]$wantType + ' for ' + [string]$spec.kind)
        }
        if ($script:P83ThreeD.ContainsKey([int]$actual.ChartType)) {
            $null = $threeD.Add($spec.key + ' is ' + [string]$script:P83ThreeD[[int]$actual.ChartType])
        }

        $anchorColumn = ''
        $anchorRow = ''
        foreach ($character in ([string]$spec.anchor).ToCharArray()) {
            if ([char]::IsDigit($character)) { $anchorRow = $anchorRow + [string]$character }
            else { $anchorColumn = $anchorColumn + [string]$character }
        }
        $wantColumn = ConvertTo-P83ColumnNumber -Letters $anchorColumn
        if (([int]$actual.AnchorRow -ne [int]$anchorRow) -or
            ([int]$actual.AnchorColumn -ne [int]$wantColumn)) {
            $null = $wrongGeometry.Add($spec.key + ': anchored at r' + [string]$actual.AnchorRow +
                                       'c' + [string]$actual.AnchorColumn + ', projected ' +
                                       [string]$spec.anchor)
        }
        # EXCEL ROUNDS A SHAPE TO WHOLE PIXELS, so the size is checked to a
        # tolerance rather than to the millimetre. Half a centimetre is far
        # tighter than any layout mistake and far looser than that rounding.
        if (([Math]::Abs([double]$actual.WidthCm - [double]$spec.width_cm) -gt 0.5) -or
            ([Math]::Abs([double]$actual.HeightCm - [double]$spec.height_cm) -gt 0.5)) {
            $null = $wrongGeometry.Add($spec.key + ': ' +
                ('{0:N2}' -f $actual.WidthCm) + 'x' + ('{0:N2}' -f $actual.HeightCm) +
                'cm, projected ' + [string]$spec.width_cm + 'x' + [string]$spec.height_cm)
        }

        $wantSeries = @($spec.series)
        if ([int]$actual.SeriesCount -ne $wantSeries.Count) {
            $null = $wrongWiring.Add($spec.key + ': ' + [string]$actual.SeriesCount +
                                     ' series, projected ' + [string]$wantSeries.Count)
            continue
        }
        $wantCategories = ConvertTo-P83NormalRange -Reference ([string]$spec.categories.range)
        for ($index = 0; $index -lt $wantSeries.Count; $index++) {
            $series = @($actual.Series)[$index]
            $want = $wantSeries[$index]
            $wantValues = ConvertTo-P83NormalRange -Reference ([string]$want.range)
            if (([string]$series.ValuesRange) -cne $wantValues) {
                $null = $wrongWiring.Add($spec.key + ' series ' + [string]($index + 1) +
                                         ': plots ' + [string]$series.ValuesRange +
                                         ', projected ' + $wantValues)
            }
            if (([string]$series.Categories) -cne $wantCategories) {
                $null = $wrongWiring.Add($spec.key + ' series ' + [string]($index + 1) +
                                         ': categories ' + [string]$series.Categories +
                                         ', projected ' + $wantCategories)
            }
            foreach ($sheet in @('_SimData', '_Calc')) {
                if (([string]$series.Formula).Contains($sheet)) {
                    $null = $machine.Add($spec.key + ' series ' + [string]($index + 1) +
                                         ' reads ' + $sheet)
                }
            }

            # THE PAYLOAD, AGAINST THE ALREADY-FROZEN BRIDGE.
            $key = [string]$spec.source_block + '.' + [string]$want.key
            if (-not $Bridge.Contains($key)) {
                $null = $wrongPayload.Add($spec.key + ': nothing frozen for ' + $key)
                continue
            }
            $cells = @($Bridge[$key])
            $points = @($series.Values)
            if ($points.Count -ne $cells.Count) {
                $null = $wrongPayload.Add($spec.key + ' series ' + [string]($index + 1) +
                                          ': ' + [string]$points.Count + ' points against ' +
                                          [string]$cells.Count + ' bridge cells')
                continue
            }
            for ($row = 0; $row -lt $points.Count; $row++) {
                $cell = $cells[$row]
                $point = $points[$row]
                $cellBlank = ((Test-P83Blank -Cell $cell) -or $cell.IsError)
                $pointBlank = (Test-P83NoPoint -Value $point)
                if ($cellBlank -ne $pointBlank) {
                    # A BRIDGE CELL SAYING "no point" AND A CHART DRAWING ONE is
                    # the fabricated zero this whole layer exists to prevent.
                    $null = $fabricated.Add($spec.key + ' row ' + [string]($row + 1) + ': cell ' +
                                            (Format-P83Cell $cell) + ', plotted ' +
                                            (Format-P83Point $point))
                    continue
                }
                if ($pointBlank) { continue }
                if (-not (Test-SimSameValue -A $cell.Value -B $point)) {
                    $null = $wrongPayload.Add($spec.key + ' row ' + [string]($row + 1) + ': cell ' +
                                              (Format-SimValue $cell.Value) + ', plotted ' +
                                              (Format-SimValue $point))
                }
                if ($wrongPayload.Count -gt 8) { break }
            }
        }
    }

    $extra = @($Frozen).Count - $expected.Count
    $null = Add-P83Check ($Stage + ': the Dashboard carries exactly the projected charts') `
        (($missing.Count -eq 0) -and ($extra -eq 0)) `
        ('missing: ' + ($missing -join ', ') + '; surplus: ' + [string]$extra)
    $null = Add-P83Check ($Stage + ': every chart is the projected Excel chart type') `
        ($wrongType.Count -eq 0) (($wrongType -join '; '))
    $null = Add-P83Check ($Stage + ': no chart is three-dimensional') `
        ($threeD.Count -eq 0) (($threeD -join '; '))
    $null = Add-P83Check ($Stage + ': every chart is anchored and sized as projected') `
        ($wrongGeometry.Count -eq 0) (($wrongGeometry -join '; '))
    $null = Add-P83Check ($Stage + ': every series and category range is the projected bridge range') `
        ($wrongWiring.Count -eq 0) (($wrongWiring -join '; '))
    $null = Add-P83Check ($Stage + ': no chart series reads the machine sheets') `
        ($machine.Count -eq 0) (($machine -join '; '))
    $null = Add-P83Check ($Stage + ': every plotted point equals the frozen bridge cell') `
        ($wrongPayload.Count -eq 0) (($wrongPayload -join '; '))
    $null = Add-P83Check ($Stage + ': no chart draws a point where the bridge says there is none') `
        ($fabricated.Count -eq 0) (($fabricated -join '; '))
    return [pscustomobject]@{
        Missing = $missing.Count; WrongType = $wrongType.Count; ThreeD = $threeD.Count
        WrongGeometry = $wrongGeometry.Count; WrongWiring = $wrongWiring.Count
        Machine = $machine.Count; WrongPayload = $wrongPayload.Count
        Fabricated = $fabricated.Count
    }
}

# THE ONE ORCHESTRATOR, so the order cannot be got wrong at a call site. Steps
# 2 through 6 happen here, in this sequence, with nothing between them. A caller
# performs step 1 before calling and step 7 after it returns.
function Invoke-P83Observation {
    param($Excel, $Workbook, $P8, $Charts, [string]$Stage)
    # (2)
    $null = Invoke-P83Recalculate -Excel $Excel -Stage $Stage
    # (3) THE AUTHORITATIVE STATE, BEFORE ANY CHART IS TOUCHED.
    $state = Get-P83StateFrozen -Workbook $Workbook -P8 $P8 -Charts $Charts
    # (4) THE BRIDGE CELLS, IMMEDIATELY AFTERWARDS.
    $bridge = Get-P83BridgeFrozen -Workbook $Workbook -Charts $Charts
    # (5) ONLY NOW ARE THE CHART OBJECTS READ.
    $objects = Get-P83Charts -Workbook $Workbook -SheetName ([string]$Charts.chart_sheet)
    # (6)
    $verdict = Invoke-P83ChartChecks -Charts $Charts -Frozen $objects -Bridge $bridge -Stage $Stage
    Write-P83Line ('    ' + $Stage + ': ' + [string]@($objects).Count + ' chart objects, ' +
                   [string]$verdict.WrongPayload + ' payload mismatches, ' +
                   [string]$verdict.Fabricated + ' fabricated points, ' +
                   [string]$verdict.WrongWiring + ' wiring faults')
    return [pscustomobject]@{
        State = $state; Bridge = $bridge; Objects = $objects; Verdict = $verdict
    }
}

# ===========================================================================
# THE STATE, READ OFF THE FROZEN CELLS AND NEVER DERIVED HERE
# ===========================================================================
# NOT ONE STATE WORD IS TYPED IN THIS FILE. Every expectation below arrives as
# a parameter, read out of the accepted projections at preflight, and every
# comparison is against a cell frozen before the charts were looked at. There is
# no PowerShell reimplementation of CURRENT, STALE, INVALID, HISTORICAL or
# OTHER Px, because a second state oracle is the one thing that would make all
# of this agree with itself and nothing else.
function Get-P83Frozen {
    param($Observation, [string]$Key)
    if (-not $Observation.State.Contains($Key)) {
        throw ('nothing was frozen under ' + [char]39 + $Key + [char]39)
    }
    return $Observation.State[$Key]
}

function Invoke-P83StateChecks {
    param($Observation, [string]$Stage, [string]$Simulation, [string]$Distribution,
          [string]$Profile, $ExpectedPx, $ExpectedYears)
    $live = Get-P83Frozen -Observation $Observation -Key 'simulation_state'
    $distributionCell = Get-P83Frozen -Observation $Observation -Key 'distribution_state'
    $profileCell = Get-P83Frozen -Observation $Observation -Key 'profile_state'
    $pxCell = Get-P83Frozen -Observation $Observation -Key 'profile_px'
    $yearsCell = Get-P83Frozen -Observation $Observation -Key 'year_count'

    $null = Add-P83Check ($Stage + ': the LIVE simulation state is ' + $Simulation) `
        (Test-SimExactText -Actual $live.Value -Expected $Simulation) (Format-P83Cell $live)
    $null = Add-P83Check ($Stage + ': the annual distributions are ' + $Distribution) `
        (Test-SimExactText -Actual $distributionCell.Value -Expected $Distribution) `
        (Format-P83Cell $distributionCell)
    $null = Add-P83Check ($Stage + ': the annual profile is ' + $Profile) `
        (Test-SimExactText -Actual $profileCell.Value -Expected $Profile) `
        (Format-P83Cell $profileCell)
    if ($null -eq $ExpectedPx) {
        $null = Add-P83Check ($Stage + ': the published profile Px is blank') `
            (Test-P83Blank -Cell $pxCell) (Format-P83Cell $pxCell)
    } else {
        $null = Add-P83Check ($Stage + ': the published profile Px is ' + [string]$ExpectedPx) `
            (Test-SimExactText -Actual $pxCell.Value -Expected ([string]$ExpectedPx)) `
            (Format-P83Cell $pxCell)
    }
    $null = Add-P83Check ($Stage + ': the year count is ' + [string]$ExpectedYears) `
        (Test-SimExactDouble -Actual $yearsCell.Value -Expected ([double]$ExpectedYears)) `
        (Format-P83Cell $yearsCell)
}

# THE HISTOGRAM'S QUALIFICATION IS THE LIVE STATE, NOT THE PERSISTED ROW - and
# this records both so the distinction is evidence rather than a claim. P8-1
# proved in live Excel that `Simulation Status (last evaluated)` can still read
# CURRENT after a request change; if that happens here, the two cells disagree
# and the report shows exactly that, with the chart qualified by the live one.
function Invoke-P83QualificationChecks {
    param($Observation, [string]$Stage, [string]$Simulation, [string]$CurrentWord)
    $live = Get-P83Frozen -Observation $Observation -Key 'simulation_state'
    $persisted = Get-P83Frozen -Observation $Observation -Key 'run_stamp.simulation_status'
    Write-P83Line ('    ' + $Stage + ': live simulation state ' + (Format-P83Cell $live))
    Write-P83Line ('    ' + $Stage + ': persisted (last evaluated) ' + (Format-P83Cell $persisted))
    if ($Simulation -cne $CurrentWord) {
        # THE CHART MUST NOT READ CURRENT WHEN THE MODEL HAS MOVED ON, whatever
        # the persisted row says.
        $null = Add-P83Check ($Stage + ': the histogram is qualified ' + $Simulation +
                              ', not by the persisted row') `
            ((Test-SimExactText -Actual $live.Value -Expected $Simulation) -and
             (-not (Test-SimExactText -Actual $live.Value -Expected $CurrentWord))) `
            ('live ' + (Format-P83Cell $live) + ' / persisted ' + (Format-P83Cell $persisted))
    }
}

# THE TORNADO NEEDS BOTH CONDITIONS, and they are two different questions.
function Invoke-P83TornadoQualification {
    param($Observation, [string]$Stage, [string]$Simulation, [string]$CurrentWord,
          [switch]$ExpectAvailable, [string]$UnavailablePhrase)
    $availability = Get-P83Frozen -Observation $Observation -Key 'sensitivity_availability'
    $live = Get-P83Frozen -Observation $Observation -Key 'simulation_state'
    $text = [string]$availability.Value
    Write-P83Line ('    ' + $Stage + ': sensitivity availability ' + (Format-P83Cell $availability))
    if ($ExpectAvailable) {
        $null = Add-P83Check ($Stage + ': the sensitivity ranking is published for this run') `
            (-not $text.Contains($UnavailablePhrase)) (Format-P83Cell $availability)
    } else {
        $null = Add-P83Check ($Stage + ': the sensitivity ranking says it is not produced') `
            ($text.Contains($UnavailablePhrase)) (Format-P83Cell $availability)
    }
    # AND THE SECOND CONDITION, WHICH THE FIRST CANNOT SEE. The availability
    # sentence compares two persisted records; it is blind to a model that has
    # moved since, so a tornado is only current when the live state agrees.
    $null = Add-P83Check ($Stage + ': the tornado carries the live simulation state as well') `
        (Test-SimExactText -Actual $live.Value -Expected $Simulation) (Format-P83Cell $live)
    if ($Simulation -cne $CurrentWord) {
        $null = Add-P83Check ($Stage + ': the tornado is qualified ' + $Simulation +
                              ' even though its ranking still names the published run') `
            (-not (Test-SimExactText -Actual $live.Value -Expected $CurrentWord)) `
            ('availability ' + (Format-P83Cell $availability) + ', live ' + (Format-P83Cell $live))
    }
}

# ===========================================================================
# WHAT EACH CHART IS PLOTTING, EXPRESSED AS A COUNT OF REAL POINTS
# ===========================================================================
# NOT A RECOMPUTATION. This counts how many of a bridge column's frozen cells
# carry a number and how many are the contracted no-point, and compares that to
# what the sheet itself published as the year count or the iteration count. It
# builds no series, sums no profile and bins nothing.
function Measure-P83Points {
    param($Observation, [string]$Key)
    $cells = @($Observation.Bridge[$Key])
    $numbers = 0
    $absent = 0
    foreach ($cell in $cells) {
        if ($cell.IsError -or (Test-P83Blank -Cell $cell)) { $absent = $absent + 1 }
        else { $numbers = $numbers + 1 }
    }
    return [pscustomobject]@{ Numbers = $numbers; Absent = $absent; Total = $cells.Count }
}

function Invoke-P83AnnualSeriesChecks {
    param($Observation, [string]$Stage, [int]$ExpectedYears)
    foreach ($key in @('annual.project_index', 'annual.calendar_year', 'annual.annual_nominal',
                       'annual.cumulative_nominal', 'annual.cumulative_pv')) {
        $measure = Measure-P83Points -Observation $Observation -Key $key
        $null = Add-P83Check ($Stage + ': ' + $key + ' carries exactly ' +
                              [string]$ExpectedYears + ' points and no more') `
            ($measure.Numbers -eq $ExpectedYears) `
            ([string]$measure.Numbers + ' numbers, ' + [string]$measure.Absent +
             ' absent, of ' + [string]$measure.Total)
    }
    # AND THE POINTS ARE THE FIRST ROWS, IN ORDER. A gap would mean a year was
    # dropped and the ones after it slid up.
    $cells = @($Observation.Bridge['annual.calendar_year'])
    $outOfOrder = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $cells.Count; $index++) {
        $present = -not ($cells[$index].IsError -or (Test-P83Blank -Cell $cells[$index]))
        $shouldBePresent = ($index -lt $ExpectedYears)
        if ($present -ne $shouldBePresent) {
            $null = $outOfOrder.Add('row ' + [string]($index + 1) + ' ' +
                                    (Format-P83Cell $cells[$index]))
        }
        if ($outOfOrder.Count -gt 6) { break }
    }
    $null = Add-P83Check ($Stage + ': the annual points are the first ' + [string]$ExpectedYears +
                          ' rows, in order, with nothing after them') `
        ($outOfOrder.Count -eq 0) (($outOfOrder -join '; '))
}

# COUNT CONSERVATION, TAKEN AGAINST WHAT THE SHEET PUBLISHED. Every iteration
# lands in exactly one bin, so the bins must sum to the published iteration
# count. This sums the frozen counts; it does not bin anything, and it takes the
# expected total off Results rather than out of a variable in this file.
function Invoke-P83HistogramChecks {
    param($Observation, [string]$Stage, [int]$ExpectedBins, [switch]$Published)
    $cells = @($Observation.Bridge['distribution.count'])
    $lower = @($Observation.Bridge['distribution.lower'])
    $upper = @($Observation.Bridge['distribution.upper'])
    $null = Add-P83Check ($Stage + ': the distribution block carries the projected ' +
                          [string]$ExpectedBins + ' bins') `
        ($cells.Count -eq $ExpectedBins) ([string]$cells.Count + ' rows')

    if (-not $Published) {
        $fabricated = New-Object System.Collections.ArrayList
        foreach ($cell in $cells) {
            if (-not ($cell.IsError -or (Test-P83Blank -Cell $cell))) {
                $null = $fabricated.Add((Format-P83Cell $cell))
            }
        }
        $null = Add-P83Check ($Stage + ': no bin carries a count when nothing is published') `
            ($fabricated.Count -eq 0) (($fabricated -join '; '))
        return
    }

    $errors = New-Object System.Collections.ArrayList
    $total = 0.0
    foreach ($cell in $cells) {
        if ($cell.IsError) { $null = $errors.Add((Format-P83Cell $cell)); continue }
        if (Test-P83Blank -Cell $cell) { continue }
        $total = $total + [double]$cell.Value
    }
    $null = Add-P83Check ($Stage + ': every bin evaluated without an Excel error') `
        ($errors.Count -eq 0) (($errors -join '; '))
    $iterations = Get-P83Frozen -Observation $Observation -Key 'run_stamp.iterations_run'
    $null = Add-P83Check ($Stage + ': the bin counts sum to the published iteration count') `
        ((Test-SimExactDouble -Actual $total -Expected ([double]$iterations.Value))) `
        ([string]$total + ' against ' + (Format-P83Cell $iterations))
    # THE EDGES SPAN THE PUBLISHED EXTREMES, and the extremes are read off the
    # Results summary rather than recomputed from the iteration column.
    $null = Add-P83Check ($Stage + ': the bin edges ascend and the last closes above the first') `
        ((-not $lower[0].IsError) -and (-not $upper[$upper.Count - 1].IsError) -and
         ([double]$upper[$upper.Count - 1].Value -gt [double]$lower[0].Value)) `
        ((Format-P83Cell $lower[0]) + ' .. ' + (Format-P83Cell $upper[$upper.Count - 1]))
}

# THE TORNADO'S ROWS AGAINST THE SHEET THAT RANKED THEM. The comparison is
# positional: bridge row k must be Sensitivity row (first + k), which is what
# "took the first N in the published order" means. Nothing here sorts, ranks or
# takes an absolute value.
function Invoke-P83TornadoRowChecks {
    param($Workbook, $Observation, $Sensitivity, [string]$SheetName, [string]$Stage,
          [switch]$Published)
    $names = @($Observation.Bridge['drivers.driver_name'])
    $rhos = @($Observation.Bridge['drivers.rho'])
    if (-not $Published) {
        $fabricated = New-Object System.Collections.ArrayList
        foreach ($cell in ($names + $rhos)) {
            if (-not ($cell.IsError -or (Test-P83Blank -Cell $cell))) {
                $null = $fabricated.Add((Format-P83Cell $cell))
            }
        }
        $null = Add-P83Check ($Stage + ': the tornado shows no driver when none is published') `
            ($fabricated.Count -eq 0) (($fabricated -join '; '))
        return 0
    }

    $columns = @{}
    foreach ($column in @($Sensitivity.columns)) { $columns[[string]$column.key] = [string]$column.column }
    $first = [int]$Sensitivity.first_row
    $mismatched = New-Object System.Collections.ArrayList
    $plotted = 0
    for ($index = 0; $index -lt $names.Count; $index++) {
        $sourceRow = $first + $index
        $sourceName = Get-P83Cell -Workbook $Workbook -SheetName $SheetName `
            -Address ($columns['driver_name'] + [string]$sourceRow)
        $sourceRho = Get-P83Cell -Workbook $Workbook -SheetName $SheetName `
            -Address ($columns['rho'] + [string]$sourceRow)
        $bridgeName = $names[$index]
        $bridgeRho = $rhos[$index]
        $sourcePresent = -not ((Test-P83Blank -Cell $sourceName) -or $sourceName.IsError)
        $bridgePresent = -not ((Test-P83Blank -Cell $bridgeName) -or $bridgeName.IsError)
        if ($sourcePresent -ne $bridgePresent) {
            $null = $mismatched.Add('row ' + [string]($index + 1) + ': sheet ' +
                                    (Format-P83Cell $sourceName) + ', bridge ' +
                                    (Format-P83Cell $bridgeName))
            continue
        }
        if (-not $sourcePresent) { continue }
        $plotted = $plotted + 1
        if (-not (Test-SimExactText -Actual $bridgeName.Value -Expected ([string]$sourceName.Value))) {
            $null = $mismatched.Add('row ' + [string]($index + 1) + ' name: ' +
                                    (Format-P83Cell $bridgeName) + ' vs ' + (Format-P83Cell $sourceName))
        }
        # THE SIGN IS THE POINT OF A TORNADO. An exact double comparison keeps a
        # negative rho negative; a magnitude test would not.
        if (-not (Test-SimExactDouble -Actual $bridgeRho.Value -Expected ([double]$sourceRho.Value))) {
            $null = $mismatched.Add('row ' + [string]($index + 1) + ' rho: ' +
                                    (Format-P83Cell $bridgeRho) + ' vs ' + (Format-P83Cell $sourceRho))
        }
    }
    $null = Add-P83Check ($Stage + ': every plotted driver is the Sensitivity row of the same rank, ' +
                          'name and signed rho') `
        ($mismatched.Count -eq 0) (($mismatched -join '; '))
    $null = Add-P83Check ($Stage + ': the tornado plots at most the projected top N') `
        ($plotted -le $names.Count) ([string]$plotted + ' of ' + [string]$names.Count)
    return $plotted
}

# TWO FROZEN BRIDGES, COMPARED. Parts that must change nothing say so by naming
# the columns that must be identical; parts that must change something name the
# same columns and require movement. Both read the captures already taken, so
# neither disturbs what was measured.
function Compare-P83Bridge {
    param($Before, $After, [string[]]$Keys, [string]$Stage, [string]$What,
          [switch]$ExpectMoved)
    $moved = New-Object System.Collections.ArrayList
    foreach ($key in $Keys) {
        $left = @($Before.Bridge[$key])
        $right = @($After.Bridge[$key])
        if ($left.Count -ne $right.Count) {
            $null = $moved.Add($key + ': ' + [string]$left.Count + ' -> ' + [string]$right.Count)
            continue
        }
        for ($index = 0; $index -lt $left.Count; $index++) {
            if (-not (Test-P83SameCellValue -Left $left[$index] -Right $right[$index])) {
                $null = $moved.Add($key + ' row ' + [string]($index + 1) + ': ' +
                                   (Format-SimValue $left[$index].Value) + ' -> ' +
                                   (Format-SimValue $right[$index].Value))
            }
            if ($moved.Count -gt 8) { break }
        }
        if ($moved.Count -gt 8) { break }
    }
    if ($ExpectMoved) {
        $null = Add-P83Check ($Stage + ': ' + $What) ($moved.Count -gt 0) `
            ('changed: ' + ($moved -join '; '))
    } else {
        $null = Add-P83Check ($Stage + ': ' + $What) ($moved.Count -eq 0) (($moved -join '; '))
    }
    return $moved.Count
}

# ===========================================================================
# THE LAYOUT THAT MAKES THE QUALIFICATION VISIBLE
# ===========================================================================
# A STATE LINE FORTY ROWS ABOVE A CHART IS NOT A QUALIFICATION. Two things fix
# that and both are checked live: the executive status block is frozen on
# screen, and the two chart-specific lines sit inside the region immediately
# above the first plot.
function Test-P83Layout {
    param($Workbook, $Dashboard, $Charts, [string]$Stage)
    $sheetName = [string]$Dashboard.sheet
    $sheets = $null; $sheet = $null; $windows = $null; $window = $null
    $split = ''
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($sheetName)
        $null = $sheet.Activate()
        $windows = $Workbook.Windows
        $window = $windows.Item(1)
        $frozen = [bool]$window.FreezePanes
        $rows = 0
        $columns = 0
        try { $rows = [int]$window.SplitRow } catch { $rows = -1 }
        try { $columns = [int]$window.SplitColumn } catch { $columns = -1 }
        $split = 'FreezePanes=' + [string]$frozen + ' SplitRow=' + [string]$rows +
                 ' SplitColumn=' + [string]$columns
        # THE MANIFEST DECLARES A CELL; EXCEL REPORTS A SPLIT. `A17` freezes the
        # sixteen rows above it and no columns, which is what is asserted.
        $wantRow = 0
        foreach ($character in ([string]$Dashboard.freeze_panes).ToCharArray()) {
            if ([char]::IsDigit($character)) { $wantRow = $wantRow * 10 + [int]::Parse([string]$character) }
        }
        $null = Add-P83Check ($Stage + ': the Dashboard freezes the rows above ' +
                              [string]$Dashboard.freeze_panes) `
            ($frozen -and ($rows -eq ($wantRow - 1)) -and ($columns -eq 0)) $split
        # AND THE FROZEN BAND REALLY COVERS THE STATUS BLOCK, so a reader
        # scrolled to a chart still sees the annual state.
        $statusLast = 0
        foreach ($section in @($Dashboard.sections)) {
            if (([string]$section.key) -cne 'status') { continue }
            foreach ($entry in @($section.rows)) {
                if ([int]$entry.row -gt $statusLast) { $statusLast = [int]$entry.row }
            }
        }
        $null = Add-P83Check ($Stage + ': the frozen band keeps the whole Result Status block on screen') `
            (($statusLast -gt 0) -and ($statusLast -le $rows)) `
            ('status ends at row ' + [string]$statusLast + ', frozen band is ' + [string]$rows + ' rows')
    } finally {
        if ($null -ne $window)  { Release-Transient $window  'Window';     $window  = $null }
        if ($null -ne $windows) { Release-Transient $windows 'Windows';    $windows = $null }
        if ($null -ne $sheet)   { Release-Transient $sheet   'Worksheet';  $sheet   = $null }
        if ($null -ne $sheets)  { Release-Transient $sheets  'Worksheets'; $sheets  = $null }
    }

    # THE TWO CHART-STATUS LINES ARE WHERE THE PROJECTION PUTS THEM, inside the
    # region and above the first plot.
    $region = $Dashboard.chart_region
    $statusSection = $null
    foreach ($section in @($Dashboard.sections)) {
        if (([string]$section.key) -ceq 'chart_status') { $statusSection = $section }
    }
    if ($null -eq $statusSection) { throw 'the Dashboard projection carries no chart-status section' }
    $misplaced = New-Object System.Collections.ArrayList
    $firstChart = 0
    foreach ($chart in @($Charts.charts)) {
        $row = 0
        foreach ($character in ([string]$chart.anchor).ToCharArray()) {
            if ([char]::IsDigit($character)) { $row = $row * 10 + [int]::Parse([string]$character) }
        }
        if (($firstChart -eq 0) -or ($row -lt $firstChart)) { $firstChart = $row }
    }
    foreach ($entry in @($statusSection.rows)) {
        $cell = Get-P83Cell -Workbook $Workbook -SheetName $sheetName `
            -Address ([string]$Dashboard.columns.nominal + [string]([int]$entry.row))
        if ($cell.IsError) { $null = $misplaced.Add([string]$entry.key + ' ' + (Format-P83Cell $cell)) }
        if (([int]$entry.row -le [int]$region.heading_row) -or
            ([int]$entry.row -ge $firstChart)) {
            $null = $misplaced.Add([string]$entry.key + ' at row ' + [string]$entry.row +
                                   ', outside the band between the region heading and the first chart')
        }
    }
    $null = Add-P83Check ($Stage + ': the chart-status lines sit inside the region, above every plot') `
        ($misplaced.Count -eq 0) (($misplaced -join '; '))
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 8 P8-3 (the analytical charts, live)' -ForegroundColor Cyan
Write-Host '================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$p8InspectPath  = Join-Path $BuildDir 'phase8_results_inspection.json'
$dashPath       = Join-Path $BuildDir 'phase8_dashboard_inspection.json'
$chartPath      = Join-Path $BuildDir 'phase8_charts_inspection.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath, $dashPath, $chartPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ($required + ' not found. Run the Stage-A build first: ' +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest      = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json
$gateBCases    = Get-Content -LiteralPath $gateBCasePath  -Raw | ConvertFrom-Json
$p7            = Get-Content -LiteralPath $p7InspectPath  -Raw | ConvertFrom-Json
$cases         = Get-Content -LiteralPath $casesPath      -Raw | ConvertFrom-Json
$p8            = Get-Content -LiteralPath $p8InspectPath  -Raw | ConvertFrom-Json
$dash          = Get-Content -LiteralPath $dashPath       -Raw | ConvertFrom-Json
$charts        = Get-Content -LiteralPath $chartPath      -Raw | ConvertFrom-Json

# THE ACCEPTED BEHAVIOURAL FIXTURE, UNCHANGED. Reusing W4's model is what makes
# this runner comparable to the accepted P8-1 and P8-2 evidence rather than a
# new experiment.
$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no behavioural scenario for P8-3 to reuse.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-P83SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a P8-3 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase8-p3-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath, $dashPath, $chartPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:P83Path = Join-Path $tempRoot 'phase8_p3_chart_surface.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$model = $case.model
$yearCount = [int]$model.timeline.duration
$iterations = [int]$case.iterations
$staleIterations = [int]$gateBCases.bounds.business_minimum_iterations + 1
$suppliedSeed = [double]$case.supplied_seed
$firstLabel = [string]$case.selected_confidence_level
$secondLabel = [string]$case.second_confidence_level
$binCount = [int]$charts.bin_contract.bin_count

# EVERY STATE WORD, READ OUT OF THE ACCEPTED PROJECTIONS. Not one is typed.
$notProduced = [string]$p7.handoff.distribution_states[0]
$notProducedProfile = [string]$p7.handoff.profile_states[0]
$annualCurrent = [string]$p7.handoff.distribution_states[1]
$profileCurrent = [string]$p7.handoff.profile_states[1]
$otherPx = [string]$p7.handoff.inconsistent_stamp_state
$historical = [string]$p7.handoff.distribution_states[$p7.handoff.distribution_states.Count - 1]
$calcNotCalculated = [string]$p7.model_states.derived_status[0]
$calcCurrent = [string]$p7.model_states.derived_status[1]
$calcStale = [string]$p7.model_states.derived_status[2]
$calcInvalid = [string]$p7.model_states.derived_status[3]
# THE SENSITIVITY SHEET'S OWN UNAVAILABLE WORDING, taken from the manifest the
# sheet was built from rather than retyped here.
$notProducedPhrase = 'Not produced for this run'

$costRegister = Get-P83Register -Manifest $manifest -Key 'cost_lines'
$victim = @($model.cost_lines)[0]
$victimId = [string]$victim.permanent_id
$invalidMaximum = [double]$victim.min_value - 1.0
$maximumOrdinal = 0
try { $maximumOrdinal = Get-P83RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_max' }
catch { $maximumOrdinal = 0 }

Write-P83Line 'PCCM - PHASE 8 P8-3: THE ANALYTICAL CHARTS, LIVE'
Write-P83Line '================================================'
Write-P83Line ''
Write-P83Line ('source revision        : ' + $revision.Head)
Write-P83Line ('charts projected       : ' + [string]@($charts.charts).Count)
Write-P83Line ('histogram bins         : ' + [string]$binCount)
Write-P83Line ('behavioural fixture    : ' + [string]$case.id + ' (' + [string]$yearCount +
               ' years, ' + [string]$iterations + ' iterations, ' + $firstLabel + ' -> ' +
               $secondLabel + ')')
Write-P83Line ''
Write-P83Line 'THE ORACLE IS THE WORKSHEET. Every chart series is compared against a bridge'
Write-P83Line 'cell frozen BEFORE the chart was read, and every state word comes from the'
Write-P83Line 'accepted projections. Nothing here bins, sums, ranks or re-simulates.'
Write-P83Line ''

$null = Add-P83Check 'the Stage-B workbook was bootstrapped' $bootstrapOk `
    ('build_stage_b.ps1 exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-P83Line ''
    Write-P83Line 'STOP. The Stage-B workbook was not produced; nothing below could run.'
    Set-Content -LiteralPath $script:P83Path -Value ($script:P83Lines -join [Environment]::NewLine)
    Write-Host ('The report is at ' + $script:P83Path) -ForegroundColor Cyan
    exit 1
}

# ===========================================================================
# THE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 8 P8-3 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-P83Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                   [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-P83Line 'ever terminated, and the workbook is never saved.'
    Write-P83Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-P83Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }
    # THE LOCKED FX SEED, CAPTURED ONCE, ON THE UNTOUCHED STAGE-B WORKBOOK, at
    # the lifecycle point every accepted runner captures it: after the compile
    # prerequisite and before the first Phase-5 mutation. The helper READS, so
    # Part 0 below still observes a workbook nothing has touched.
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # ===================================================================
    # PART 0 - NOTHING HAS BEEN PRODUCED
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART 0 - NOTHING HAS RUN'
    Write-P83Line '------------------------'
    $observation0 = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part 0'
    Invoke-P83StateChecks -Observation $observation0 -Stage 'part 0' `
        -Simulation $calcNotCalculated -Distribution $notProduced -Profile $notProducedProfile `
        -ExpectedPx $null -ExpectedYears 0
    # NOTHING IS PLOTTED, AND THAT IS DIFFERENT FROM PLOTTING ZEROS. The
    # comparison inside the observation already refuses a point where the bridge
    # says there is none; these count what the bridge itself carries.
    Invoke-P83AnnualSeriesChecks -Observation $observation0 -Stage 'part 0' -ExpectedYears 0
    Invoke-P83HistogramChecks -Observation $observation0 -Stage 'part 0' -ExpectedBins $binCount
    $null = Invoke-P83TornadoRowChecks -Workbook $wb -Observation $observation0 `
        -Sensitivity $simInspection -SheetName ([string]$charts.sensitivity_sheet) -Stage 'part 0'
    Invoke-P83TornadoQualification -Observation $observation0 -Stage 'part 0' `
        -Simulation $calcNotCalculated -CurrentWord $calcCurrent `
        -UnavailablePhrase 'No simulation has been published'
    Test-P83Layout -Workbook $wb -Dashboard $dash -Charts $charts -Stage 'part 0'

    # ===================================================================
    # PART A1 - A CURRENT RESULT, WITH SENSITIVITY DELIBERATELY NOT RUN
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART A1 - CALCULATE, SIMULATE, ANNUAL - AND NO SENSITIVITY'
    Write-P83Line '----------------------------------------------------------'
    Write-P83Line 'SENSITIVITY IS NOT INVOKED HERE ON PURPOSE. A tornado with nothing to plot'
    Write-P83Line 'must say so; an empty chart that looks like a finished analysis is the'
    Write-P83Line 'failure this part exists to catch.'
    $applied = ''
    $fixtureFailure = ''
    try {
        $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
            -Inspection $inspection -Model $model)
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name) `
            -Value ([double]$iterations)
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.random_seed.defined_name) `
            -Value $suppliedSeed
        Set-P83NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
            -Value $firstLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-P83Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the P8-3 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'P8-3 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-P83Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-P83Check 'PCCM_RunSimulation succeeded' `
        ($simResult -like 'OK|*') $simResult 'PREREQUISITE'
    if (-not ($simResult -like 'OK|*')) { throw ('the simulation did not commit: ' + $simResult) }

    $annualResult = Invoke-P83Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P83Check ([string]$p7.command_surface.annual_endpoint + ' succeeded') `
        ($annualResult -like 'OK|*') $annualResult 'PREREQUISITE'
    if (-not ($annualResult -like 'OK|*')) { throw ('the annual step did not commit: ' + $annualResult) }

    $observationA1 = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part A1'
    Invoke-P83StateChecks -Observation $observationA1 -Stage 'part A1' `
        -Simulation $calcCurrent -Distribution $annualCurrent -Profile $profileCurrent `
        -ExpectedPx $firstLabel -ExpectedYears $yearCount
    Invoke-P83AnnualSeriesChecks -Observation $observationA1 -Stage 'part A1' -ExpectedYears $yearCount
    Invoke-P83HistogramChecks -Observation $observationA1 -Stage 'part A1' -ExpectedBins $binCount -Published
    Invoke-P83QualificationChecks -Observation $observationA1 -Stage 'part A1' `
        -Simulation $calcCurrent -CurrentWord $calcCurrent
    # THE TORNADO HAS NOTHING, AND SAYS SO.
    $null = Invoke-P83TornadoRowChecks -Workbook $wb -Observation $observationA1 `
        -Sensitivity $simInspection -SheetName ([string]$charts.sensitivity_sheet) -Stage 'part A1'
    Invoke-P83TornadoQualification -Observation $observationA1 -Stage 'part A1' `
        -Simulation $calcCurrent -CurrentWord $calcCurrent `
        -UnavailablePhrase $notProducedPhrase

    # ===================================================================
    # PART A2 - THE SENSITIVITY ENDPOINT, INVOKED EXPLICITLY
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART A2 - THE SENSITIVITY ENDPOINT, INVOKED BY THIS RUNNER'
    Write-P83Line '----------------------------------------------------------'
    Write-P83Line 'AN ACCEPTANCE ACTION, NOT WORKBOOK BEHAVIOUR. Nothing on the sheet starts an'
    Write-P83Line 'analysis to populate a chart; an operator asks, and this runner is asking.'
    $sensitivityResult = Invoke-P83Endpoint -Excel $excel `
        -Endpoint ([string]$charts.sensitivity_endpoint)
    $null = Add-P83Check ([string]$charts.sensitivity_endpoint + ' succeeded') `
        ($sensitivityResult -like 'OK|*') $sensitivityResult 'PREREQUISITE'
    if (-not ($sensitivityResult -like 'OK|*')) {
        # A LEGITIMATE REFUSAL IS A FINDING, NOT SOMETHING TO WORK AROUND.
        Write-P83Line ''
        Write-P83Line ('STOP. ' + [string]$charts.sensitivity_endpoint + ' refused: ' +
                       $sensitivityResult)
        Write-P83Line 'The tornado cannot be accepted without a published ranking, and this'
        Write-P83Line 'runner does not patch around a refusal it did not expect.'
        throw ('the sensitivity endpoint refused: ' + $sensitivityResult)
    }
    $observationA2 = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part A2'
    Invoke-P83StateChecks -Observation $observationA2 -Stage 'part A2' `
        -Simulation $calcCurrent -Distribution $annualCurrent -Profile $profileCurrent `
        -ExpectedPx $firstLabel -ExpectedYears $yearCount
    $plottedA2 = Invoke-P83TornadoRowChecks -Workbook $wb -Observation $observationA2 `
        -Sensitivity $simInspection -SheetName ([string]$charts.sensitivity_sheet) `
        -Stage 'part A2' -Published
    $null = Add-P83Check 'part A2: the tornado plots at least one ranked driver' `
        ($plottedA2 -ge 1) ([string]$plottedA2 + ' drivers plotted')
    Invoke-P83TornadoQualification -Observation $observationA2 -Stage 'part A2' `
        -Simulation $calcCurrent -CurrentWord $calcCurrent -ExpectAvailable `
        -UnavailablePhrase $notProducedPhrase
    # THE OTHER THREE CHARTS DID NOT MOVE. Sensitivity publishes its own block;
    # it does not touch the annual answer or the distribution.
    $movedA2 = Compare-P83Bridge -Before $observationA1 -After $observationA2 `
        -Keys @('annual.annual_nominal', 'annual.cumulative_nominal', 'annual.cumulative_pv',
                'distribution.count') -Stage 'part A2' `
        -What 'the annual and distribution bridges did not move when sensitivity ran'
    $null = $movedA2

    # ===================================================================
    # PART B - THE SELECTOR MOVES AND NOTHING ELSE DOES
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART B - THE SELECTOR MOVES, RECALCULATION ONLY'
    Write-P83Line '-----------------------------------------------'
    Write-P83Line 'NO ENDPOINT IS INVOKED IN THIS PART. The only thing that happens between the'
    Write-P83Line 'observation above and the one below is a selector write and a recalculation.'
    # STEP 1 - A CELL WRITE, NOT A COMMAND.
    Set-P83NamedText -Workbook $wb `
        -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
        -Value $secondLabel
    $observationB = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part B'
    # THE ANNUAL PRODUCT IS RETIRED; THE DISTRIBUTION IS NOT. A reporting
    # selector is not a simulation fingerprint input, so the histogram's state
    # must stay exactly where it was while the profile's moves to OTHER Px.
    Invoke-P83StateChecks -Observation $observationB -Stage 'part B' `
        -Simulation $calcCurrent -Distribution $annualCurrent -Profile $otherPx `
        -ExpectedPx $firstLabel -ExpectedYears $yearCount
    Invoke-P83QualificationChecks -Observation $observationB -Stage 'part B' `
        -Simulation $calcCurrent -CurrentWord $calcCurrent
    # THE PRESERVED P80 PROFILE IS STILL WHAT THE ANNUAL CHARTS PLOT.
    $null = Compare-P83Bridge -Before $observationA2 -After $observationB `
        -Keys @('annual.project_index', 'annual.calendar_year', 'annual.annual_nominal',
                'annual.cumulative_nominal', 'annual.cumulative_pv') -Stage 'part B' `
        -What 'the annual and cumulative series still carry the preserved profile'
    $null = Compare-P83Bridge -Before $observationA2 -After $observationB `
        -Keys @('distribution.lower', 'distribution.upper', 'distribution.count') `
        -Stage 'part B' -What 'the histogram is untouched by a reporting selector'
    $null = Compare-P83Bridge -Before $observationA2 -After $observationB `
        -Keys @('drivers.driver_name', 'drivers.rho') -Stage 'part B' `
        -What 'the tornado is untouched by a reporting selector'
    Invoke-P83TornadoQualification -Observation $observationB -Stage 'part B' `
        -Simulation $calcCurrent -CurrentWord $calcCurrent -ExpectAvailable `
        -UnavailablePhrase $notProducedPhrase
    # AND THE RUN IDENTITY DID NOT MOVE.
    $runB = Get-P83Frozen -Observation $observationB -Key 'run_stamp.run_id'
    $runA2 = Get-P83Frozen -Observation $observationA2 -Key 'run_stamp.run_id'
    $null = Add-P83Check 'part B: no run identity moved from a presentation recalculation' `
        (Test-P83SameCellValue -Left $runA2 -Right $runB) `
        ((Format-P83Cell $runA2) + ' -> ' + (Format-P83Cell $runB))

    # ===================================================================
    # PART C - THE ANNUAL STEP ALONE
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART C - THE ANNUAL ENDPOINT ALONE'
    Write-P83Line '----------------------------------'
    $rerun = Invoke-P83Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P83Check ('part C: ' + [string]$p7.command_surface.annual_endpoint +
                          ' succeeded on the moved selector') ($rerun -like 'OK|*') $rerun
    $observationC = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part C'
    Invoke-P83StateChecks -Observation $observationC -Stage 'part C' `
        -Simulation $calcCurrent -Distribution $annualCurrent -Profile $profileCurrent `
        -ExpectedPx $secondLabel -ExpectedYears $yearCount
    Invoke-P83AnnualSeriesChecks -Observation $observationC -Stage 'part C' -ExpectedYears $yearCount
    # THE ANNUAL CHARTS FOLLOWED THE REPUBLISHED PROFILE.
    $null = Compare-P83Bridge -Before $observationB -After $observationC `
        -Keys @('annual.annual_nominal', 'annual.cumulative_nominal', 'annual.cumulative_pv') `
        -Stage 'part C' -What 'the annual and cumulative series followed the republished profile' `
        -ExpectMoved
    # THE OTHER TWO DID NOT.
    $null = Compare-P83Bridge -Before $observationB -After $observationC `
        -Keys @('distribution.lower', 'distribution.upper', 'distribution.count') `
        -Stage 'part C' -What 'the histogram did not move for an annual rerun'
    $null = Compare-P83Bridge -Before $observationB -After $observationC `
        -Keys @('drivers.driver_name', 'drivers.rho') -Stage 'part C' `
        -What 'the tornado did not move for an annual rerun'
    $runC = Get-P83Frozen -Observation $observationC -Key 'run_stamp.run_id'
    $iterationsC = Get-P83Frozen -Observation $observationC -Key 'run_stamp.iterations_run'
    $iterationsA2 = Get-P83Frozen -Observation $observationA2 -Key 'run_stamp.iterations_run'
    $null = Add-P83Check 'part C: the simulation identity and iteration count are unchanged' `
        ((Test-P83SameCellValue -Left $runA2 -Right $runC) -and
         (Test-P83SameCellValue -Left $iterationsA2 -Right $iterationsC)) `
        ((Format-P83Cell $runA2) + ' -> ' + (Format-P83Cell $runC) + ', iterations ' +
         (Format-P83Cell $iterationsA2) + ' -> ' + (Format-P83Cell $iterationsC))

    # ===================================================================
    # PART D - THE REQUEST MOVES, AND ONLY A RECALCULATION FOLLOWS
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART D - ITERATIONS CHANGE, RECALCULATION ONLY'
    Write-P83Line '----------------------------------------------'
    Write-P83Line 'NO ENDPOINT IS INVOKED IN THIS PART EITHER. This is the transition the live'
    Write-P83Line 'simulation state exists for: the persisted (last evaluated) row may still'
    Write-P83Line 'read CURRENT, and no chart may be qualified by it.'
    # STEP 1 - A CELL WRITE, AND NOTHING ELSE.
    $iterationsControl = [string]$simInspection.controls.monte_carlo_iterations.defined_name
    $requestBefore = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    Set-NamedValue -Workbook $wb -DefinedName $iterationsControl -Value ([double]$staleIterations)
    $requestAfter = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    $null = Add-P83Check 'part D: the projected iteration control really moved' `
        ((([string]$requestAfter) -ceq ([string]$staleIterations)) -and
         (([string]$requestBefore) -cne ([string]$requestAfter))) `
        ($iterationsControl + ': ' + [string]$requestBefore + ' -> ' + [string]$requestAfter) `
        'PREREQUISITE'
    $observationD = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part D'
    Invoke-P83StateChecks -Observation $observationD -Stage 'part D' `
        -Simulation $calcStale -Distribution $historical -Profile $historical `
        -ExpectedPx $secondLabel -ExpectedYears $yearCount
    # THE QUALIFICATION IS THE LIVE STATE, AND THE PERSISTED ROW IS RECORDED
    # BESIDE IT SO THE DISTINCTION IS EVIDENCE.
    Invoke-P83QualificationChecks -Observation $observationD -Stage 'part D' `
        -Simulation $calcStale -CurrentWord $calcCurrent
    Invoke-P83TornadoQualification -Observation $observationD -Stage 'part D' `
        -Simulation $calcStale -CurrentWord $calcCurrent -ExpectAvailable `
        -UnavailablePhrase $notProducedPhrase
    # THE PRESERVED EVIDENCE IS STILL THERE. A stale answer is qualified, not
    # erased: the histogram still describes the successful publication and the
    # annual charts still carry the P50 profile.
    Invoke-P83HistogramChecks -Observation $observationD -Stage 'part D' -ExpectedBins $binCount -Published
    Invoke-P83AnnualSeriesChecks -Observation $observationD -Stage 'part D' -ExpectedYears $yearCount
    $null = Compare-P83Bridge -Before $observationC -After $observationD `
        -Keys @('annual.annual_nominal', 'annual.cumulative_nominal', 'annual.cumulative_pv',
                'distribution.lower', 'distribution.upper', 'distribution.count',
                'drivers.driver_name', 'drivers.rho') -Stage 'part D' `
        -What 'no chart bridge moved from a presentation recalculation'
    $runD = Get-P83Frozen -Observation $observationD -Key 'run_stamp.run_id'
    $null = Add-P83Check 'part D: no run identity moved from a presentation recalculation' `
        (Test-P83SameCellValue -Left $runC -Right $runD) `
        ((Format-P83Cell $runC) + ' -> ' + (Format-P83Cell $runD))
    # STEP 7 - THE ONLY OUT-OF-CELL CALL IN THIS PART, AND IT IS LAST.
    $calcStatusD = ''
    try { $calcStatusD = [string]$excel.Run('PCCM_CalculationStatus') } catch { $calcStatusD = (Format-Err $_) }
    $null = Add-P83Check 'part D: the deterministic calculation is still CURRENT, so the state moved for the request alone' `
        ($calcStatusD -ceq $calcCurrent) $calcStatusD

    # ===================================================================
    # PART E - THE MODEL STOPS BEING VALID
    # ===================================================================
    Write-P83Line ''
    Write-P83Line 'PART E - AN INVALID MODEL'
    Write-P83Line '-------------------------'
    $victimRow = Get-P83RegisterRowIndex -Workbook $wb -Register $costRegister -PermanentId $victimId
    $null = Add-P83Check 'part E: the cost line the corpus names is in the register' `
        ($victimRow -ge 1) ($victimId + ' at body row ' + [string]$victimRow) 'PREREQUISITE'
    if ($victimRow -lt 1) { throw ('the P8-3 register edit could not find ' + $victimId) }
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
        -TableName ([string]$costRegister.table_name) -RowIndex $victimRow `
        -ColumnIndex $maximumOrdinal -Value $invalidMaximum
    $calcAnnouncement = Invoke-P83Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'
    $null = Add-P83Check 'part E: PCCM_Calculate refuses the edited model' `
        ($calcAnnouncement -like 'FAIL|*') $calcAnnouncement
    $observationE = Invoke-P83Observation -Excel $excel -Workbook $wb -P8 $p8 -Charts $charts `
        -Stage 'part E'
    Invoke-P83StateChecks -Observation $observationE -Stage 'part E' `
        -Simulation $calcInvalid -Distribution $historical -Profile $historical `
        -ExpectedPx $secondLabel -ExpectedYears $yearCount
    Invoke-P83QualificationChecks -Observation $observationE -Stage 'part E' `
        -Simulation $calcInvalid -CurrentWord $calcCurrent
    Invoke-P83TornadoQualification -Observation $observationE -Stage 'part E' `
        -Simulation $calcInvalid -CurrentWord $calcCurrent -ExpectAvailable `
        -UnavailablePhrase $notProducedPhrase
    # THE SUCCESSFUL HISTORICAL PAYLOAD IS INTACT.
    Invoke-P83HistogramChecks -Observation $observationE -Stage 'part E' -ExpectedBins $binCount -Published
    Invoke-P83AnnualSeriesChecks -Observation $observationE -Stage 'part E' -ExpectedYears $yearCount
    $null = Compare-P83Bridge -Before $observationD -After $observationE `
        -Keys @('annual.annual_nominal', 'annual.cumulative_nominal', 'annual.cumulative_pv',
                'distribution.count', 'drivers.driver_name', 'drivers.rho') -Stage 'part E' `
        -What 'the preserved payload survived the invalid model'
    Test-P83Layout -Workbook $wb -Dashboard $dash -Charts $charts -Stage 'part E'
    # STEP 7 AGAIN, AFTER THE FROZEN OBSERVATION.
    $calcStatusE = ''
    try { $calcStatusE = [string]$excel.Run('PCCM_CalculationStatus') } catch { $calcStatusE = (Format-Err $_) }
    $null = Add-P83Check 'part E: the model is no longer CURRENT' `
        ($calcStatusE -cne $calcCurrent) $calcStatusE

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-P83Line ''
    Write-P83Line ('THE P8-3 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-P83Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-P83Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-P83Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-P83Line ''
        Write-P83Line 'EXCEL SHUTDOWN'
        Write-P83Line '--------------'
        if ($naturalExit) {
            Write-P83Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                           [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'P8-3'
            Write-P83Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-P83Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:P83Residual)) {
            Write-P83Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-P83Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# VERDICT
# ===========================================================================
$null = Add-P83Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-P83Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-P83Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-P83Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-P83Check 'every COM release left 0 outstanding references' `
    (@($script:P83Residual).Count -eq 0) ((@($script:P83Residual)) -join '; ')
$null = Add-P83Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:P83Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:P83Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-P83Line ''
Write-P83Line 'VERDICT'
Write-P83Line '-------'
Write-P83Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
               [string]$failedPrereqs.Count + ' failed')
Write-P83Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
               [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-P83Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-P83Line ''
if ($ok) {
    Write-P83Line 'P8-3: PASS'
} else {
    Write-P83Line 'P8-3: FAIL'
    Write-P83Line ''
    Write-P83Line 'STOP AND REVIEW. Do not start Phase 9 until this is understood.'
}
Write-P83Line ''
Write-P83Line ('report                 : ' + $script:P83Path)
Set-Content -LiteralPath $script:P83Path -Value ($script:P83Lines -join [Environment]::NewLine)
Write-Host ''
Write-Host ('The report is at ' + $script:P83Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
