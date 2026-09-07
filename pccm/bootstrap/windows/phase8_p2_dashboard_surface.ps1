<#
.SYNOPSIS
    PCCM Phase 8 - the MINIMAL P8-2 runner: the Dashboard executive summary, live.

.DESCRIPTION
    P8-2 BUILT A MIRROR. This runs it.

    THE CLAIM UNDER TEST IS ONE SENTENCE: every analytical cell on the Dashboard
    holds `=IF(Results!$D$nn="","",Results!$D$nn)` and therefore says exactly
    what Results says, including when Results says nothing. That is a claim about
    TWO SHEETS AGREEING, and nothing else, so this runner asserts nothing about
    how a total, a contingency, a percentile or a reconciliation was arrived at.
    Those belong to Phase 7 and to P8-1, both already accepted live.

    SO THE ORACLE IS RESULTS, NOT ARITHMETIC. Every comparison below is
    Dashboard-against-frozen-Results. There is no second total, no re-derived
    state word, no recomputed identity and no independent percentile anywhere in
    this file. A runner that recomputed what it was checking would agree with
    itself about a sheet that had drifted.

    AND THE ORDER IS PART OF THE MEASUREMENT. For every part:

      1  make the state transition
      2  recalculate the workbook normally
      3  FREEZE the Results cells the projection names
      4  FREEZE the Dashboard cells the projection names
      5  compare the frozen Dashboard against the frozen Results
      6  only then, if a diagnostic genuinely needs it, call anything out of cell

    Steps 3 and 4 happen with no recalculation between them and no out-of-cell
    call before them. P8-1 learned this the expensive way: a direct accessor
    invoked before the worksheet observation manufactures the agreement the
    observation was supposed to discover. Nothing here may call an endpoint or an
    accessor to make the Dashboard right before looking at it.

    SIX PARTS, ONE SESSION, ONE DISPOSABLE WORKBOOK:

      0  EMPTY       nothing has run: the mirrors evaluate, the state words are
                     NOT PRODUCED, the money is BLANK rather than a fabricated
                     zero, and the reserved chart region is empty
      A  CURRENT     the accepted P8-1 fixture through the accepted workflow:
                     CURRENT / CURRENT / P80 / 4, and every mirrored cell equal
                     to its Results source
      B  SELECTOR    P80 -> P50, recalculation only, NO endpoint: the selected
                     level follows, the profile does not, and the reconciliation
                     verdict arrives with its whole qualifier attached
      C  RERUN       the annual step alone at P50: the mirror follows Results
                     rather than caching what it showed a moment ago
      D  REQUEST     iterations 1000 -> 1001, recalculation only, NO endpoint:
                     HISTORICAL on both state lines, the payload still visible
                     and still qualified
      E  INVALID     a cost line bound below its own minimum: the calculation
                     refuses, the state lines stay HISTORICAL, and this runner
                     asserts NO "INVALID" indicator on the Dashboard - P8-2
                     reported that Results publishes no live deterministic
                     status, and inventing one here would test a sheet nobody
                     built

    TWO SHEETS, AND ONLY TWO. Every live assertion reads Results or Dashboard.
    `_SimData` and `_Calc` are never read: this runner proves a PRESENTATION
    layer, and the persisted payload's own integrity is P8-1's evidence, taken
    here through the surface that presents it rather than around it.

    NUMBERS ARE READ, NOT LOOKED AT. Every comparison uses `Value2` and
    `Formula`; the formatted `Text` is captured for the report and is never the
    authority for anything.

    WHAT THIS RUNNER IS NOT. It re-tests no Phase-7 scenario, re-proves no P8-1
    assertion, closes nothing, and draws no chart.

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
# define them, because in a Gate-B run the Phase-4 driver dot-sources the
# scenarios file and the helpers are already in scope. Reaching the scenarios
# file directly leaves that dependency unmet, and W1 died on exactly that, one
# helper at a time.
#
# They are copied BYTE FOR BYTE from `phase7_timing_scenarios.ps1`, the same
# bytes P8-1 copied, so there is one behaviour rather than three. A source
# control pins them to it.
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
$script:P82Lines = New-Object System.Collections.ArrayList
$script:P82Path = ''
$script:P82Checks = New-Object System.Collections.ArrayList
$script:P82Residual = New-Object System.Collections.ArrayList

function Write-P82Line {
    param([string]$Text = '')
    $null = $script:P82Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:P82Path)) {
        try {
            Set-Content -LiteralPath $script:P82Path `
                -Value ($script:P82Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-P82Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:P82Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-P82Line $line
    return $Ok
}

function Invoke-P82Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:P82Residual.Add(
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

function Get-P82SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a P82 result could not be attributed to a source revision')
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
function Invoke-P82Endpoint {
    param($Excel, [string]$Endpoint)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run($Endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

function Set-P82NamedText {
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
function Get-P82Register {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-P82RegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}

function Get-P82RegisterRowIndex {
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
$script:P82ErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}

function Get-P82Cell {
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
            if ($script:P82ErrorCodes.ContainsKey($code)) {
                $errorName = [string]$script:P82ErrorCodes[$code]
            }
        }
        # A CELL WHOSE TEXT IS AN ERROR NAME IS AN ERROR whatever its Value2
        # arrived as. Both routes are kept because neither is guaranteed alone.
        if ([string]::IsNullOrEmpty($errorName)) {
            foreach ($name in $script:P82ErrorCodes.Values) {
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

function Format-P82Cell {
    param($Cell)
    if ($null -eq $Cell) { return '<not read>' }
    if ($Cell.IsError) { return ([string]$Cell.ErrorName + ' (error)') }
    return ([string]$Cell.Sheet + '!' + [string]$Cell.Address + ' = ' +
            (Format-SimValue $Cell.Value) + ' shown as ' + [char]39 + [string]$Cell.Text + [char]39)
}

function Test-P82Blank {
    param($Cell)
    if ($null -eq $Cell) { return $false }
    if ($Cell.IsError) { return $false }
    return (Test-SimBlank -Value $Cell.Value)
}

# EXACT AGREEMENT, AND BLANK COUNTS AS A VALUE. A mirror of an empty cell must
# be empty; `Test-SimSameValue` settles the ordinary cases and the two blanks are
# named first because "" and $null and a missing value all arrive here.
function Test-P82SameCellValue {
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
function Invoke-P82Recalculate {
    param($Excel, [string]$Stage)
    $mode = '<unreadable>'
    try { $mode = [string]$Excel.Calculation } catch { $mode = '<unreadable>' }
    $failure = ''
    try { $Excel.Calculate() } catch { $failure = (Format-Err $_) }
    Write-P82Line ('    recalculated at ' + $Stage + ' (Application.Calculation = ' + $mode + ')')
    return (Add-P82Check ($Stage + ': the workbook recalculated') `
        ([string]::IsNullOrWhiteSpace($failure)) ($failure + ' calculation mode ' + $mode) `
        'PREREQUISITE')
}

# ===========================================================================
# THE PROJECTION, READ THE ONE WAY IT IS MEANT TO BE READ
# ===========================================================================
# EVERY PAIR THIS RUNNER COMPARES COMES FROM HERE, and none of it is spelled in
# this file. `phase8_dashboard_inspection.json` carries, for each mirrored row,
# the Dashboard cell and the Results cell it must equal; a runner that typed
# either address would be a second declaration of something the manifest owns,
# and P7-4 is what a second declaration costs when the first one moves.
function Get-P82MirrorPairs {
    param($Dash)
    $pairs = New-Object System.Collections.ArrayList
    foreach ($section in @($Dash.sections)) {
        foreach ($entry in @($section.rows)) {
            foreach ($measure in $entry.cells.PSObject.Properties.Name) {
                $cells = $entry.cells.$measure
                $null = $pairs.Add([pscustomobject]@{
                    Section = [string]$section.key
                    Key = [string]$entry.key
                    Measure = [string]$measure
                    Label = [string]$entry.label
                    Format = [string]$entry.format
                    SourceBlock = [string]$entry.source_block
                    SourceKey = [string]$entry.source_key
                    Dashboard = [string]$cells.dashboard
                    # `Results!D47` in the projection; the sheet and the address
                    # are split here rather than parsed at each call site.
                    ResultsSheet = ([string]$cells.results).Split('!')[0]
                    ResultsAddress = ([string]$cells.results).Split('!')[1]
                })
            }
        }
    }
    return @($pairs)
}

# STEP 3 OF THE OBSERVATION RULE. Results first, and on its own, so that what
# the Dashboard is compared against is a value that existed BEFORE the Dashboard
# was looked at - not one re-read afterwards and not one this runner computed.
function Get-P82ResultsFrozen {
    param($Workbook, $Pairs)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($pair in $Pairs) {
        $key = $pair.Dashboard
        if ($out.Contains($key)) { continue }
        $out.Add($key, (Get-P82Cell -Workbook $Workbook -SheetName $pair.ResultsSheet `
            -Address $pair.ResultsAddress))
    }
    return $out
}

# STEP 4. The Dashboard, immediately afterwards, with nothing in between.
function Get-P82DashboardFrozen {
    param($Workbook, [string]$SheetName, $Pairs)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($pair in $Pairs) {
        $key = $pair.Dashboard
        if ($out.Contains($key)) { continue }
        $out.Add($key, (Get-P82Cell -Workbook $Workbook -SheetName $SheetName -Address $key))
    }
    return $out
}

# ===========================================================================
# STEP 5 - THE COMPARISON, AND IT IS THE WHOLE POINT
# ===========================================================================
# FOUR QUESTIONS PER MIRRORED CELL, and they fail separately on purpose:
#
#   evaluated   neither cell is an Excel error. A #VALUE! on either side makes
#               every other question about that pair unanswerable, and P8-1's
#               first Windows run is why this is asked first.
#   formula     the Dashboard cell really is the projected mirror of the
#               projected Results cell. Value agreement alone would pass for a
#               hard-coded constant that happened to match today.
#   value       the two stored values are equal, with blank counting as a value.
#               This is the claim; the other three exist so a failure of it can
#               be attributed.
#   format      the Dashboard renders the same kind of thing Results does, so a
#               4.5e-13 difference is not displayed as 0 on one sheet and not
#               the other.
function Invoke-P82MirrorChecks {
    param($Pairs, $ResultsFrozen, $DashboardFrozen, $Dash, [string]$Stage)
    $errored = New-Object System.Collections.ArrayList
    $notMirror = New-Object System.Collections.ArrayList
    $disagreed = New-Object System.Collections.ArrayList
    $misformatted = New-Object System.Collections.ArrayList
    $template = [string]$Dash.mirror_formula
    $formats = $Dash.number_formats

    foreach ($pair in $Pairs) {
        $results = $ResultsFrozen[$pair.Dashboard]
        $dashboard = $DashboardFrozen[$pair.Dashboard]

        if ($results.IsError -or $dashboard.IsError) {
            $null = $errored.Add($pair.Key + '.' + $pair.Measure + ': Results ' +
                                 (Format-P82Cell $results) + ', Dashboard ' +
                                 (Format-P82Cell $dashboard))
            continue
        }

        # THE REFERENCE THE PROJECTION NAMES, spelled absolutely the way the
        # builder spells it. Nothing here reconstructs a row number.
        $column = ''
        $row = ''
        foreach ($character in ([string]$pair.ResultsAddress).ToCharArray()) {
            if ([char]::IsDigit($character)) { $row = $row + [string]$character }
            else { $column = $column + [string]$character }
        }
        $reference = $pair.ResultsSheet + '!$' + $column + '$' + $row
        $expectedFormula = $template.Replace('{ref}', $reference)
        if (([string]$dashboard.Formula) -cne $expectedFormula) {
            $null = $notMirror.Add($pair.Dashboard + ': ' + [string]$dashboard.Formula +
                                   ' (expected ' + $expectedFormula + ')')
        }

        if (-not (Test-P82SameCellValue -Left $results -Right $dashboard)) {
            $null = $disagreed.Add($pair.Key + '.' + $pair.Measure + ': Results ' +
                                   (Format-SimValue $results.Value) + ' vs Dashboard ' +
                                   (Format-SimValue $dashboard.Value))
        }

        $expectedFormat = [string]$formats.($pair.Format)
        if (([string]$dashboard.NumberFormat) -cne $expectedFormat) {
            $null = $misformatted.Add($pair.Dashboard + ': ' + [string]$dashboard.NumberFormat +
                                      ' (expected ' + $expectedFormat + ')')
        }
    }

    $null = Add-P82Check ($Stage + ': every mirrored cell evaluated without an Excel error') `
        ($errored.Count -eq 0) (($errored -join '; '))
    $null = Add-P82Check ($Stage + ': every Dashboard cell is the projected mirror of its Results cell') `
        ($notMirror.Count -eq 0) (($notMirror -join '; '))
    $null = Add-P82Check ($Stage + ': every Dashboard value equals the frozen Results value') `
        ($disagreed.Count -eq 0) (($disagreed -join '; '))
    $null = Add-P82Check ($Stage + ': every Dashboard cell carries the projected number format') `
        ($misformatted.Count -eq 0) (($misformatted -join '; '))
    return [pscustomobject]@{
        Errored = $errored.Count; NotMirror = $notMirror.Count
        Disagreed = $disagreed.Count; Misformatted = $misformatted.Count
    }
}

# THE ONE ORCHESTRATOR, so the order cannot be got wrong at a call site. Steps
# 2, 3, 4 and 5 happen here, in this sequence, with nothing between them. A
# caller performs step 1 before calling and step 6 after it returns.
function Invoke-P82Observation {
    param($Excel, $Workbook, $Dash, $Pairs, [string]$Stage)
    # (2)
    $null = Invoke-P82Recalculate -Excel $Excel -Stage $Stage
    # (3) RESULTS FIRST, AND BEFORE ANYTHING IS ASKED OUT OF A CELL.
    $results = Get-P82ResultsFrozen -Workbook $Workbook -Pairs $Pairs
    # (4) THE DASHBOARD, IMMEDIATELY AFTERWARDS.
    $dashboard = Get-P82DashboardFrozen -Workbook $Workbook `
        -SheetName ([string]$Dash.sheet) -Pairs $Pairs
    # (5)
    $verdict = Invoke-P82MirrorChecks -Pairs $Pairs -ResultsFrozen $results `
        -DashboardFrozen $dashboard -Dash $Dash -Stage $Stage
    Write-P82Line ('    ' + $Stage + ': ' + [string]@($Pairs).Count + ' mirrored cells, ' +
                   [string]$verdict.Disagreed + ' disagreed, ' + [string]$verdict.NotMirror +
                   ' not a mirror, ' + [string]$verdict.Errored + ' errored')
    return [pscustomobject]@{
        Results = $results; Dashboard = $dashboard; Verdict = $verdict
    }
}

# A NAMED ROW OUT OF A FROZEN OBSERVATION. The state assertions below read the
# frozen capture rather than the sheet, so nothing they do can disturb what was
# measured.
function Get-P82Frozen {
    param($Observation, $Pairs, [string]$Key, [string]$Measure = 'nominal')
    foreach ($pair in $Pairs) {
        if (($pair.Key -ceq $Key) -and ($pair.Measure -ceq $Measure)) {
            return $Observation.Dashboard[$pair.Dashboard]
        }
    }
    throw ('the projection carries no mirrored row ' + [char]39 + $Key + '.' + $Measure + [char]39)
}

# THE STATE WORDS, ASSERTED ON THE DASHBOARD AND ONLY FROM THE FROZEN CAPTURE.
# The words themselves are the accepted Phase-7 vocabulary, read out of the
# acceptance projection; not one of them is typed in this file.
function Invoke-P82StateChecks {
    param($Observation, $Pairs, [string]$Stage, [string]$Distribution, [string]$Profile,
          $ExpectedPx, $ExpectedYears)
    $distributionCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'distribution_state'
    $profileCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'profile_state'
    $pxCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'profile_px'
    $yearsCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'year_count'

    $null = Add-P82Check ($Stage + ': the Dashboard shows the annual distributions as ' + $Distribution) `
        (Test-SimExactText -Actual $distributionCell.Value -Expected $Distribution) `
        (Format-P82Cell $distributionCell)
    $null = Add-P82Check ($Stage + ': the Dashboard shows the annual profile as ' + $Profile) `
        (Test-SimExactText -Actual $profileCell.Value -Expected $Profile) `
        (Format-P82Cell $profileCell)
    # THE TWO STATE LINES ARE TWO ANSWERS. A moved selector retires the profile
    # and must not retire the ladders, so a Dashboard that showed one word twice
    # would have destroyed the distinction one sheet after it was made.
    if ($Distribution -cne $Profile) {
        $null = Add-P82Check ($Stage + ': the two state lines are not the same word') `
            ((([string]$distributionCell.Value)) -cne (([string]$profileCell.Value))) `
            ((Format-P82Cell $distributionCell) + ' / ' + (Format-P82Cell $profileCell))
    }
    if ($null -eq $ExpectedPx) {
        $null = Add-P82Check ($Stage + ': the published profile Px is blank') `
            (Test-P82Blank -Cell $pxCell) (Format-P82Cell $pxCell)
    } else {
        $null = Add-P82Check ($Stage + ': the published profile Px is ' + [string]$ExpectedPx) `
            (Test-SimExactText -Actual $pxCell.Value -Expected ([string]$ExpectedPx)) `
            (Format-P82Cell $pxCell)
    }
    $null = Add-P82Check ($Stage + ': the year count is ' + [string]$ExpectedYears) `
        (Test-SimExactDouble -Actual $yearsCell.Value -Expected ([double]$ExpectedYears)) `
        (Format-P82Cell $yearsCell)
}

# THE SELECTED LEVEL AND THE PUBLISHED PROFILE'S Px ARE TWO DIFFERENT QUESTIONS,
# and part B is the part where they differ. Comparing them as one field is the
# specific mistake this check exists to make impossible.
function Invoke-P82SelectorChecks {
    param($Observation, $Pairs, [string]$Stage, [string]$SelectedLabel, $ProfilePx)
    $selectedCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'selected_confidence_level'
    $pxCell = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'profile_px'
    $null = Add-P82Check ($Stage + ': the Dashboard selected confidence level is ' + $SelectedLabel) `
        (Test-SimExactText -Actual $selectedCell.Value -Expected $SelectedLabel) `
        (Format-P82Cell $selectedCell)
    if ($null -ne $ProfilePx) {
        $expectSame = ($SelectedLabel -ceq [string]$ProfilePx)
        $areSame = ((([string]$selectedCell.Value)) -ceq (([string]$pxCell.Value)))
        $null = Add-P82Check ($Stage + ': the selected level and the published profile Px are ' +
                              $(if ($expectSame) { 'the same' } else { 'different' }) + ' fields, read separately') `
            ($areSame -eq $expectSame) `
            ((Format-P82Cell $selectedCell) + ' / ' + (Format-P82Cell $pxCell))
    }
}

# THE TWO LADDERS STAY TWO. W5 read one as the other and failed four
# reconciliations by exactly the deterministic base; nothing on the Dashboard
# recomputes the difference that would expose the substitution, so it is
# asserted directly - on both measures, from the frozen capture.
function Invoke-P82HeadlineChecks {
    param($Observation, $Pairs, [string]$Stage, [switch]$Published)
    foreach ($measure in @('nominal', 'pv')) {
        $total = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'selected_total' -Measure $measure
        $contingency = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'contingency' -Measure $measure
        $base = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'deterministic_base' -Measure $measure
        if (-not $Published) {
            $null = Add-P82Check ($Stage + ': the ' + $measure +
                                  ' headline cells are blank, not fabricated zeros') `
                ((Test-P82Blank -Cell $total) -and (Test-P82Blank -Cell $contingency) -and
                 (Test-P82Blank -Cell $base)) `
                ((Format-P82Cell $total) + ' / ' + (Format-P82Cell $contingency) + ' / ' +
                 (Format-P82Cell $base))
            continue
        }
        $null = Add-P82Check ($Stage + ': the ' + $measure + ' TOTAL and contingency are numbers') `
            (($total.Value -is [double]) -and ($contingency.Value -is [double])) `
            ((Format-P82Cell $total) + ' / ' + (Format-P82Cell $contingency))
        $null = Add-P82Check ($Stage + ': the ' + $measure +
                              ' TOTAL is not the contingency wearing its name') `
            (-not (Test-P82SameCellValue -Left $total -Right $contingency)) `
            ((Format-P82Cell $total) + ' vs ' + (Format-P82Cell $contingency))
    }
}

# THE VERDICT ARRIVES WHOLE OR IT DOES NOT ARRIVE. Results appends its own
# qualifier - " - for the OTHER Px profile, not the current model answer" - and
# a Dashboard that trimmed it would turn a reconciliation of somebody else's
# profile into what looks like the current answer. The mirror check above
# already proves the strings are equal; this says what the string must CONTAIN,
# so a Results that stopped qualifying is caught here rather than passing
# because both sheets agree about the wrong thing.
function Invoke-P82QualifierChecks {
    param($Observation, $Pairs, $Dash, $P8, [string]$Stage, [string]$ProfileState,
          [string]$CurrentProfile)
    $qualifier = [string]$P8.reconciliation.verdicts.qualifier_suffix
    foreach ($measure in @('nominal', 'pv')) {
        $status = Get-P82Frozen -Observation $Observation -Pairs $Pairs -Key 'status' -Measure $measure
        $text = [string]$status.Value
        if ($ProfileState -ceq $CurrentProfile) {
            $null = Add-P82Check ($Stage + ': the ' + $measure +
                                  ' reconciliation carries no historical qualifier') `
                (-not $text.Contains($qualifier)) (Format-P82Cell $status)
        } else {
            $null = Add-P82Check ($Stage + ': the ' + $measure + ' reconciliation says it is the ' +
                                  $ProfileState + ' profile, not the current answer') `
                ($text.Contains($qualifier) -and $text.Contains($ProfileState)) `
                (Format-P82Cell $status)
        }
    }
}

# ===========================================================================
# THE RESERVED CHART REGION - LIVE PROOF THAT NOTHING IS DRAWN
# ===========================================================================
# THREE SEPARATE ABSENCES, because they fail for different reasons. A chart
# object is P8-3 started early; a picture is a chart pasted rather than drawn;
# and a value in the reserved rows is the summary having grown down into space
# the next step needs. The heading and the note are the only two cells the
# projection permits inside the region.
function Test-P82ChartRegion {
    param($Workbook, $Dash, [string]$Stage)
    $region = $Dash.chart_region
    $sheetName = [string]$Dash.sheet
    $sheets = $null; $sheet = $null; $charts = $null; $shapes = $null
    $chartCount = -1; $shapeCount = -1
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($sheetName)
        $charts = $sheet.ChartObjects()
        $chartCount = [int]$charts.Count
        $shapes = $sheet.Shapes
        $shapeCount = [int]$shapes.Count
    } finally {
        if ($null -ne $shapes) { Release-Transient $shapes 'Shapes';       $shapes = $null }
        if ($null -ne $charts) { Release-Transient $charts 'ChartObjects'; $charts = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(Dashboard)'; $sheet = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';   $sheets = $null }
    }
    $null = Add-P82Check ($Stage + ': the Dashboard carries no chart object') `
        ($chartCount -eq 0) ([string]$chartCount + ' chart object(s)')
    $null = Add-P82Check ($Stage + ': the Dashboard carries no picture or shape') `
        ($shapeCount -eq 0) ([string]$shapeCount + ' shape(s)')

    # THE HEADING AND THE NOTE ARE PERMITTED; EVERYTHING BELOW THEM IS NOT.
    $occupied = New-Object System.Collections.ArrayList
    $columns = @([string]$Dash.columns.label, [string]$Dash.columns.nominal, [string]$Dash.columns.pv)
    for ($row = [int]$region.first_row; $row -le [int]$region.last_row; $row++) {
        foreach ($column in $columns) {
            $cell = Get-P82Cell -Workbook $Workbook -SheetName $sheetName -Address ($column + [string]$row)
            if (-not (Test-P82Blank -Cell $cell)) {
                $null = $occupied.Add((Format-P82Cell $cell))
            }
            if ($occupied.Count -gt 8) { break }
        }
        if ($occupied.Count -gt 8) { break }
    }
    $null = Add-P82Check ($Stage + ': the reserved chart rows are empty below the heading and note') `
        ($occupied.Count -eq 0) (($occupied -join '; '))
    # AND THE TWO CELLS THAT ARE SUPPOSED TO BE THERE, ARE.
    $heading = Get-P82Cell -Workbook $Workbook -SheetName $sheetName `
        -Address ([string]$Dash.columns.label + [string]([int]$region.heading_row))
    $note = Get-P82Cell -Workbook $Workbook -SheetName $sheetName `
        -Address ([string]$Dash.columns.label + [string]([int]$region.note_row))
    $null = Add-P82Check ($Stage + ': the reserved region is announced by a heading and a note') `
        ((-not (Test-P82Blank -Cell $heading)) -and (-not (Test-P82Blank -Cell $note))) `
        ((Format-P82Cell $heading) + ' / ' + (Format-P82Cell $note))
}

# ===========================================================================
# THE PAYLOAD, WITNESSED THROUGH RESULTS RATHER THAN AROUND IT
# ===========================================================================
# PARTS B AND D CHANGE NOTHING AND MUST BE SEEN TO CHANGE NOTHING. The obvious
# way to prove that is to read the persisted `_SimData` block before and after -
# and this runner deliberately does not, because a PRESENTATION acceptance that
# reached into the machine sheet would be proving the wrong layer's property and
# would need a second oracle to do it.
#
# SO THE WITNESS IS RESULTS ITSELF: the run stamp, the summary ladder, the
# selected block and a bounded sample of the annual table, frozen through the
# same reader everything else uses. If the persisted payload moved, these move;
# if they did not move, nothing a recalculation did reached the stored answer.
function Get-P82ResultsWitness {
    param($Workbook, $P8, [int]$Sample = 6)
    $sheet = [string]$P8.sheet
    $nominal = [string]$P8.columns.nominal
    $pv = [string]$P8.columns.pv
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($field in @($P8.run_stamp.fields)) {
        $address = $nominal + [string]([int]$field.row)
        $out.Add(('run_stamp.' + [string]$field.key),
                 (Get-P82Cell -Workbook $Workbook -SheetName $sheet -Address $address))
    }
    foreach ($metric in @($P8.summary.metrics)) {
        foreach ($column in @($nominal, $pv)) {
            $address = $column + [string]([int]$metric.row)
            $out.Add(('summary.' + [string]$metric.key + '.' + $column),
                     (Get-P82Cell -Workbook $Workbook -SheetName $sheet -Address $address))
        }
    }
    # THE ANNUAL TABLE, SAMPLED RATHER THAN SWEPT. Two hundred rows read twice a
    # part is a COM cost this runner does not need to pay: the rows the answer
    # occupies are the rows that would move.
    for ($offset = 0; $offset -lt $Sample; $offset++) {
        $row = [int]$P8.annual.first_row + $offset
        foreach ($column in @($P8.annual.columns)) {
            $address = [string]$column.column + [string]$row
            $out.Add(('annual.' + [string]($offset + 1) + '.' + [string]$column.key),
                     (Get-P82Cell -Workbook $Workbook -SheetName $sheet -Address $address))
        }
    }
    return $out
}

function Compare-P82Witness {
    param($Before, $After, [string]$Stage, [string]$What)
    $moved = New-Object System.Collections.ArrayList
    foreach ($key in $Before.Keys) {
        if (-not (Test-P82SameCellValue -Left $Before[$key] -Right $After[$key])) {
            $null = $moved.Add([string]$key + ': ' + (Format-SimValue $Before[$key].Value) +
                               ' -> ' + (Format-SimValue $After[$key].Value))
        }
        if ($moved.Count -gt 10) { break }
    }
    $null = Add-P82Check ($Stage + ': ' + $What) ($moved.Count -eq 0) (($moved -join '; '))
    return $moved.Count
}

# THE ONE ABSENCE PART E EXISTS TO ASSERT. P8-2 reported that Results publishes
# no live deterministic-model status; the Dashboard therefore shows none, and
# this runner must not quietly start expecting one. It checks the opposite: that
# no Dashboard cell has acquired a calculation-state word Results never gave it.
function Test-P82NoInventedStatus {
    param($Observation, $Pairs, $Vocabulary, [string]$Stage)
    $invented = New-Object System.Collections.ArrayList
    foreach ($pair in $Pairs) {
        $cell = $Observation.Dashboard[$pair.Dashboard]
        $mirror = $Observation.Results[$pair.Dashboard]
        $text = [string]$cell.Value
        foreach ($word in @($Vocabulary)) {
            if ([string]::IsNullOrEmpty([string]$word)) { continue }
            if ($text.Contains([string]$word) -and
                (-not ([string]$mirror.Value).Contains([string]$word))) {
                $null = $invented.Add($pair.Key + ' says ' + [char]39 + [string]$word +
                                      [char]39 + ' and Results does not')
            }
        }
    }
    $null = Add-P82Check ($Stage + ': the Dashboard invents no state word Results did not give it') `
        ($invented.Count -eq 0) (($invented -join '; '))
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 8 P8-2 (the Dashboard executive summary, live)' -ForegroundColor Cyan
Write-Host '==========================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$p8InspectPath  = Join-Path $BuildDir 'phase8_results_inspection.json'
$dashPath       = Join-Path $BuildDir 'phase8_dashboard_inspection.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath, $dashPath)) {
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

# THE ACCEPTED BEHAVIOURAL FIXTURE, UNCHANGED. P8-2 presents an answer; it does
# not get to choose a friendlier one, and reusing W4's model is what makes this
# runner comparable to the accepted P8-1 evidence rather than a new experiment.
$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no behavioural scenario for P8-2 to reuse.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-P82SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a P8-2 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase8-p2-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath, $dashPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:P82Path = Join-Path $tempRoot 'phase8_p2_dashboard_surface.txt'
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
$notProduced = [string]$p7.handoff.distribution_states[0]
$annualCurrent = [string]$p7.handoff.distribution_states[1]
$profileCurrent = [string]$p7.handoff.profile_states[1]
$otherPx = [string]$p7.handoff.inconsistent_stamp_state
$historical = [string]$p7.handoff.distribution_states[$p7.handoff.distribution_states.Count - 1]
$notProducedProfile = [string]$p7.handoff.profile_states[0]
# EVERY STATE WORD THE PROJECT HAS, for the invented-status sweep. Read from the
# accepted projections; not one of them is typed in this file.
$vocabulary = @()
foreach ($word in @($p7.handoff.distribution_states)) { $vocabulary += [string]$word }
foreach ($word in @($p7.handoff.profile_states)) { $vocabulary += [string]$word }
$vocabulary += [string]$p7.handoff.inconsistent_stamp_state
foreach ($word in @($p7.model_states.PSObject.Properties.Name)) {
    $vocabulary += [string]$p7.model_states.$word
}

$pairs = Get-P82MirrorPairs -Dash $dash
# PART E'S EDIT, RESOLVED THE WAY P8-1 RESOLVES IT: the first cost line of the
# accepted model, with its maximum driven below its own minimum. The corpus
# carries no `invalid_edit` block and inventing one here would be a second
# fixture authority.
$calcCurrent = [string]$p7.model_states.derived_status[1]
$costRegister = Get-P82Register -Manifest $manifest -Key 'cost_lines'
$victim = @($model.cost_lines)[0]
$victimId = [string]$victim.permanent_id
$invalidMaximum = [double]$victim.min_value - 1.0
$maximumOrdinal = 0
try { $maximumOrdinal = Get-P82RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_max' }
catch { $maximumOrdinal = 0 }

Write-P82Line 'PCCM - PHASE 8 P8-2: THE DASHBOARD EXECUTIVE SUMMARY, LIVE'
Write-P82Line '=========================================================='
Write-P82Line ''
Write-P82Line ('source revision        : ' + $revision.Head)
Write-P82Line ('mirrored cells         : ' + [string]@($pairs).Count)
Write-P82Line ('behavioural fixture    : ' + [string]$case.id + ' (' + [string]$yearCount +
               ' years, ' + [string]$iterations + ' iterations, ' + $firstLabel + ' -> ' +
               $secondLabel + ')')
Write-P82Line ''
Write-P82Line 'THE ORACLE IS RESULTS. Every comparison below is Dashboard against a Results'
Write-P82Line 'value frozen BEFORE the Dashboard was read. Nothing here recomputes a total, a'
Write-P82Line 'contingency, a percentile, a state word or a reconciliation.'
Write-P82Line ''

$null = Add-P82Check 'the Stage-B workbook was bootstrapped' $bootstrapOk `
    ('build_stage_b.ps1 exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-P82Line ''
    Write-P82Line 'STOP. The Stage-B workbook was not produced; nothing below could run.'
    Set-Content -LiteralPath $script:P82Path -Value ($script:P82Lines -join [Environment]::NewLine)
    Write-Host ('The report is at ' + $script:P82Path) -ForegroundColor Cyan
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

$rel = New-ReleaseLedger 'phase 8 P8-2 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-P82Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                   [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-P82Line 'ever terminated, and the workbook is never saved.'
    Write-P82Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-P82Check 'the current VBAProject compiles in real Excel' `
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
    Write-P82Line ''
    Write-P82Line 'PART 0 - NOTHING HAS RUN'
    Write-P82Line '------------------------'
    $observation0 = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part 0'
    if ($observation0.Verdict.Errored -gt 0) {
        Write-P82Line ''
        Write-P82Line 'STOP. A MIRRORED CELL DID NOT EVALUATE.'
        Write-P82Line 'Nothing below is attempted: a Dashboard that cannot evaluate before a'
        Write-P82Line 'single run is a finding, and working around it would destroy it.'
        throw 'a mirrored cell did not evaluate in part 0'
    }
    Invoke-P82StateChecks -Observation $observation0 -Pairs $pairs -Stage 'part 0' `
        -Distribution $notProduced -Profile $notProducedProfile -ExpectedPx $null -ExpectedYears 0
    # THE FABRICATED ZERO IS THE WHOLE HAZARD HERE. Excel reads a blank
    # reference back as 0, so an unguarded mirror would print a confident 0
    # beside the words NOT PRODUCED.
    Invoke-P82HeadlineChecks -Observation $observation0 -Pairs $pairs -Stage 'part 0'
    $fabricated = New-Object System.Collections.ArrayList
    foreach ($pair in $pairs) {
        if (($pair.Format -cne 'money') -and ($pair.Format -cne 'delta')) { continue }
        $cell = $observation0.Dashboard[$pair.Dashboard]
        if (-not (Test-P82Blank -Cell $cell)) {
            $null = $fabricated.Add($pair.Key + '.' + $pair.Measure + ' = ' + (Format-P82Cell $cell))
        }
    }
    $null = Add-P82Check 'part 0: every monetary mirror is blank, not a fabricated zero' `
        ($fabricated.Count -eq 0) (($fabricated -join '; '))
    Test-P82ChartRegion -Workbook $wb -Dash $dash -Stage 'part 0'
    # THE LABELS COME FROM RESULTS, and the projection records which. A sheet
    # that had drifted into typing its own would show it here.
    $labels = New-Object System.Collections.ArrayList
    foreach ($section in @($dash.sections)) {
        foreach ($entry in @($section.rows)) {
            $labelCell = Get-P82Cell -Workbook $wb -SheetName ([string]$dash.sheet) `
                -Address ([string]$dash.columns.label + [string]([int]$entry.row))
            if (-not (Test-SimExactText -Actual $labelCell.Value -Expected ([string]$entry.label))) {
                $null = $labels.Add([string]$entry.key + ' ' + (Format-P82Cell $labelCell))
            }
        }
    }
    $null = Add-P82Check 'part 0: every Dashboard label is the label Results shows' `
        ($labels.Count -eq 0) (($labels -join '; '))

    # ===================================================================
    # PART A - A SUCCESSFUL CURRENT RESULT
    # ===================================================================
    Write-P82Line ''
    Write-P82Line 'PART A - THE ACCEPTED FIXTURE, THROUGH THE ACCEPTED WORKFLOW'
    Write-P82Line '-------------------------------------------------------------'
    Write-P82Line ('    at the first reporting level, ' + $firstLabel)
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
        Set-P82NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
            -Value $firstLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-P82Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the P8-2 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'P8-2 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-P82Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-P82Check 'PCCM_RunSimulation succeeded' `
        ($simResult -like 'OK|*') $simResult 'PREREQUISITE'
    if (-not ($simResult -like 'OK|*')) { throw ('the simulation did not commit: ' + $simResult) }

    $annualResult = Invoke-P82Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P82Check ([string]$p7.command_surface.annual_endpoint + ' succeeded') `
        ($annualResult -like 'OK|*') $annualResult 'PREREQUISITE'
    if (-not ($annualResult -like 'OK|*')) { throw ('the annual step did not commit: ' + $annualResult) }

    $observationA = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part A'
    Invoke-P82StateChecks -Observation $observationA -Pairs $pairs -Stage 'part A' `
        -Distribution $annualCurrent -Profile $profileCurrent -ExpectedPx $firstLabel `
        -ExpectedYears $yearCount
    Invoke-P82SelectorChecks -Observation $observationA -Pairs $pairs -Stage 'part A' `
        -SelectedLabel $firstLabel -ProfilePx $firstLabel
    Invoke-P82HeadlineChecks -Observation $observationA -Pairs $pairs -Stage 'part A' -Published
    Invoke-P82QualifierChecks -Observation $observationA -Pairs $pairs -Dash $dash -P8 $p8 `
        -Stage 'part A' -ProfileState $profileCurrent -CurrentProfile $profileCurrent
    Test-P82NoInventedStatus -Observation $observationA -Pairs $pairs -Vocabulary $vocabulary `
        -Stage 'part A'
    Test-P82ChartRegion -Workbook $wb -Dash $dash -Stage 'part A'
    $witnessA = Get-P82ResultsWitness -Workbook $wb -P8 $p8

    # ===================================================================
    # PART B - THE SELECTOR MOVES AND NOTHING ELSE DOES
    # ===================================================================
    Write-P82Line ''
    Write-P82Line 'PART B - THE SELECTOR MOVES, RECALCULATION ONLY'
    Write-P82Line '-----------------------------------------------'
    Write-P82Line ('    ' + $firstLabel + ' -> ' + $secondLabel)
    Write-P82Line 'NO ENDPOINT IS INVOKED IN THIS PART. The only thing that happens between the'
    Write-P82Line 'observation above and the one below is a selector write and a recalculation.'
    # STEP 1 - THE TRANSITION, AND IT IS A CELL WRITE, NOT A COMMAND.
    Set-P82NamedText -Workbook $wb `
        -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
        -Value $secondLabel
    $observationB = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part B'
    # THE DISTRIBUTIONS ARE STILL CURRENT AND THE PROFILE IS NOT. A moved
    # selector retires the blend at one Px; it cannot make a ladder taken across
    # every iteration wrong, and the two state lines are where that is visible.
    Invoke-P82StateChecks -Observation $observationB -Pairs $pairs -Stage 'part B' `
        -Distribution $annualCurrent -Profile $otherPx -ExpectedPx $firstLabel `
        -ExpectedYears $yearCount
    # THE SELECTED LEVEL FOLLOWED; THE PUBLISHED PROFILE'S Px DID NOT. This is
    # the part where the two fields differ, which is the only reason a reader
    # can tell they are two fields.
    Invoke-P82SelectorChecks -Observation $observationB -Pairs $pairs -Stage 'part B' `
        -SelectedLabel $secondLabel -ProfilePx $firstLabel
    Invoke-P82HeadlineChecks -Observation $observationB -Pairs $pairs -Stage 'part B' -Published
    # THE TOTAL FOLLOWED THE SELECTOR, and the mirror check has already proved it
    # equals Results. What is asserted here is that it MOVED - a Dashboard that
    # had cached part A's total would still equal a Results that had also cached
    # it, and only the movement distinguishes the two.
    $movedTotal = New-Object System.Collections.ArrayList
    foreach ($measure in @('nominal', 'pv')) {
        $before = Get-P82Frozen -Observation $observationA -Pairs $pairs -Key 'selected_total' -Measure $measure
        $after = Get-P82Frozen -Observation $observationB -Pairs $pairs -Key 'selected_total' -Measure $measure
        if (Test-P82SameCellValue -Left $before -Right $after) {
            $null = $movedTotal.Add($measure + ' stayed at ' + (Format-SimValue $before.Value))
        }
    }
    $null = Add-P82Check ('part B: the displayed TOTAL followed the selector to ' + $secondLabel) `
        ($movedTotal.Count -eq 0) (($movedTotal -join '; '))
    Invoke-P82QualifierChecks -Observation $observationB -Pairs $pairs -Dash $dash -P8 $p8 `
        -Stage 'part B' -ProfileState $otherPx -CurrentProfile $profileCurrent
    Test-P82NoInventedStatus -Observation $observationB -Pairs $pairs -Vocabulary $vocabulary `
        -Stage 'part B'
    # NOTHING WAS PUBLISHED. The persisted payload and the run identity are
    # witnessed through Results, which is the surface that presents them.
    $witnessB = Get-P82ResultsWitness -Workbook $wb -P8 $p8
    $null = Compare-P82Witness -Before $witnessA -After $witnessB -Stage 'part B' `
        -What 'the persisted run stamp and annual payload did not move'

    # ===================================================================
    # PART C - THE ANNUAL STEP ALONE
    # ===================================================================
    Write-P82Line ''
    Write-P82Line 'PART C - THE ANNUAL ENDPOINT ALONE'
    Write-P82Line '----------------------------------'
    Write-P82Line ('    at the moved reporting level, ' + $secondLabel)
    $rerun = Invoke-P82Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P82Check ('part C: ' + [string]$p7.command_surface.annual_endpoint +
                          ' succeeded on the moved selector') ($rerun -like 'OK|*') $rerun
    $observationC = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part C'
    Invoke-P82StateChecks -Observation $observationC -Pairs $pairs -Stage 'part C' `
        -Distribution $annualCurrent -Profile $profileCurrent -ExpectedPx $secondLabel `
        -ExpectedYears $yearCount
    Invoke-P82SelectorChecks -Observation $observationC -Pairs $pairs -Stage 'part C' `
        -SelectedLabel $secondLabel -ProfilePx $secondLabel
    Invoke-P82HeadlineChecks -Observation $observationC -Pairs $pairs -Stage 'part C' -Published
    Invoke-P82QualifierChecks -Observation $observationC -Pairs $pairs -Dash $dash -P8 $p8 `
        -Stage 'part C' -ProfileState $profileCurrent -CurrentProfile $profileCurrent
    # THE MIRROR FOLLOWED A REFRESHED RESULTS rather than caching what it showed
    # a moment ago. The reconciliation sum is the cell that must have moved: the
    # profile was republished at a different Px.
    $refreshed = New-Object System.Collections.ArrayList
    foreach ($measure in @('nominal', 'pv')) {
        $before = Get-P82Frozen -Observation $observationB -Pairs $pairs -Key 'profile_sum' -Measure $measure
        $after = Get-P82Frozen -Observation $observationC -Pairs $pairs -Key 'profile_sum' -Measure $measure
        if (Test-P82SameCellValue -Left $before -Right $after) {
            $null = $refreshed.Add($measure + ' profile sum stayed at ' + (Format-SimValue $before.Value))
        }
    }
    $null = Add-P82Check 'part C: the mirrored profile sum followed the republished profile' `
        ($refreshed.Count -eq 0) (($refreshed -join '; '))
    # THE SIMULATION ITSELF DID NOT RERUN. The annual step consumes a published
    # distribution; it does not produce one.
    $iterationsCell = Get-P82Frozen -Observation $observationC -Pairs $pairs -Key 'iterations_run'
    $iterationsBefore = Get-P82Frozen -Observation $observationA -Pairs $pairs -Key 'iterations_run'
    $runIdCell = Get-P82Frozen -Observation $observationC -Pairs $pairs -Key 'run_id'
    $runIdBefore = Get-P82Frozen -Observation $observationA -Pairs $pairs -Key 'run_id'
    $null = Add-P82Check 'part C: the simulation identity and iteration count are unchanged' `
        ((Test-P82SameCellValue -Left $iterationsBefore -Right $iterationsCell) -and
         (Test-P82SameCellValue -Left $runIdBefore -Right $runIdCell)) `
        ((Format-P82Cell $runIdBefore) + ' -> ' + (Format-P82Cell $runIdCell) + ', iterations ' +
         (Format-P82Cell $iterationsBefore) + ' -> ' + (Format-P82Cell $iterationsCell))
    Test-P82NoInventedStatus -Observation $observationC -Pairs $pairs -Vocabulary $vocabulary `
        -Stage 'part C'
    $witnessC = Get-P82ResultsWitness -Workbook $wb -P8 $p8

    # ===================================================================
    # PART D - THE REQUEST MOVES, AND ONLY A RECALCULATION FOLLOWS
    # ===================================================================
    Write-P82Line ''
    Write-P82Line 'PART D - ITERATIONS CHANGE, RECALCULATION ONLY'
    Write-P82Line '----------------------------------------------'
    Write-P82Line 'NO ENDPOINT IS INVOKED IN THIS PART EITHER. The volatile adapters below the'
    Write-P82Line 'Results state cells exist for exactly this transition; if the Dashboard still'
    Write-P82Line 'says CURRENT here, either they did not fire or the mirror is stale.'
    # STEP 1 - A CELL WRITE, AND NOTHING ELSE.
    $iterationsControl = [string]$simInspection.controls.monte_carlo_iterations.defined_name
    $requestBefore = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    Set-NamedValue -Workbook $wb -DefinedName $iterationsControl -Value ([double]$staleIterations)
    $requestAfter = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    $null = Add-P82Check 'part D: the projected iteration control really moved' `
        ((([string]$requestAfter) -ceq ([string]$staleIterations)) -and
         (([string]$requestBefore) -cne ([string]$requestAfter))) `
        ($iterationsControl + ': ' + [string]$requestBefore + ' -> ' + [string]$requestAfter) `
        'PREREQUISITE'
    $observationD = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part D'
    Invoke-P82StateChecks -Observation $observationD -Pairs $pairs -Stage 'part D' `
        -Distribution $historical -Profile $historical -ExpectedPx $secondLabel `
        -ExpectedYears $yearCount
    Invoke-P82SelectorChecks -Observation $observationD -Pairs $pairs -Stage 'part D' `
        -SelectedLabel $secondLabel -ProfilePx $secondLabel
    # THE HISTORICAL ANSWER IS STILL VISIBLE. Preserving it is the accepted
    # behaviour; presenting it without saying whose it is would not be.
    Invoke-P82HeadlineChecks -Observation $observationD -Pairs $pairs -Stage 'part D' -Published
    Invoke-P82QualifierChecks -Observation $observationD -Pairs $pairs -Dash $dash -P8 $p8 `
        -Stage 'part D' -ProfileState $historical -CurrentProfile $profileCurrent
    Test-P82NoInventedStatus -Observation $observationD -Pairs $pairs -Vocabulary $vocabulary `
        -Stage 'part D'
    # A RECALCULATION PUBLISHED NOTHING. Witnessed against part C, the last state
    # anything was legitimately written in.
    $witnessD = Get-P82ResultsWitness -Workbook $wb -P8 $p8
    $null = Compare-P82Witness -Before $witnessC -After $witnessD -Stage 'part D' `
        -What 'the persisted run stamp and annual payload survived the recalculation'
    # STEP 6 - THE ONLY OUT-OF-CELL CALL IN THIS PART, AND IT IS AFTER THE
    # OBSERVATION IS FROZEN. Part D's whole question is what the worksheet said
    # before this runner asked anything, so the diagnostic comes last.
    $calcStatusD = ''
    try { $calcStatusD = [string]$excel.Run('PCCM_CalculationStatus') } catch { $calcStatusD = (Format-Err $_) }
    $null = Add-P82Check 'part D: the deterministic calculation is still CURRENT, so the state moved for the request alone' `
        ($calcStatusD -ceq $calcCurrent) $calcStatusD

    # ===================================================================
    # PART E - THE MODEL STOPS BEING VALID
    # ===================================================================
    Write-P82Line ''
    Write-P82Line 'PART E - AN INVALID MODEL'
    Write-P82Line '-------------------------'
    Write-P82Line 'NO INVALID INDICATOR IS EXPECTED ON THE DASHBOARD. P8-2 reported that Results'
    Write-P82Line 'publishes no live deterministic-model status; this part asserts the ABSENCE of'
    Write-P82Line 'one rather than testing a row nobody built.'
    $victimRow = Get-P82RegisterRowIndex -Workbook $wb -Register $costRegister -PermanentId $victimId
    $null = Add-P82Check 'part E: the cost line the corpus names is in the register' `
        ($victimRow -ge 1) ($victimId + ' at body row ' + [string]$victimRow) 'PREREQUISITE'
    if ($victimRow -lt 1) { throw ('the P8-2 register edit could not find ' + $victimId) }
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
        -TableName ([string]$costRegister.table_name) -RowIndex $victimRow `
        -ColumnIndex $maximumOrdinal -Value $invalidMaximum
    $calcAnnouncement = Invoke-P82Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'
    $null = Add-P82Check 'part E: PCCM_Calculate refuses the edited model' `
        ($calcAnnouncement -like 'FAIL|*') $calcAnnouncement
    $observationE = Invoke-P82Observation -Excel $excel -Workbook $wb -Dash $dash `
        -Pairs $pairs -Stage 'part E'
    Invoke-P82StateChecks -Observation $observationE -Pairs $pairs -Stage 'part E' `
        -Distribution $historical -Profile $historical -ExpectedPx $secondLabel `
        -ExpectedYears $yearCount
    Invoke-P82HeadlineChecks -Observation $observationE -Pairs $pairs -Stage 'part E' -Published
    Invoke-P82QualifierChecks -Observation $observationE -Pairs $pairs -Dash $dash -P8 $p8 `
        -Stage 'part E' -ProfileState $historical -CurrentProfile $profileCurrent
    # THE ABSENCE PART E EXISTS FOR. The refusal is real and the calculation is
    # no longer CURRENT - the diagnostic below says so - and NOT ONE Dashboard
    # cell has acquired a word Results did not give it. A sheet that started
    # showing INVALID here would be showing a state its source does not publish.
    Test-P82NoInventedStatus -Observation $observationE -Pairs $pairs -Vocabulary $vocabulary `
        -Stage 'part E'
    $witnessE = Get-P82ResultsWitness -Workbook $wb -P8 $p8
    $null = Compare-P82Witness -Before $witnessD -After $witnessE -Stage 'part E' `
        -What 'the preserved historical payload is still intact'
    Test-P82ChartRegion -Workbook $wb -Dash $dash -Stage 'part E'
    # STEP 6 AGAIN, AFTER THE FROZEN OBSERVATION.
    $calcStatusE = ''
    try { $calcStatusE = [string]$excel.Run('PCCM_CalculationStatus') } catch { $calcStatusE = (Format-Err $_) }
    $null = Add-P82Check 'part E: the model is no longer CURRENT' `
        ($calcStatusE -cne $calcCurrent) $calcStatusE

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-P82Line ''
    Write-P82Line ('THE P8-2 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-P82Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-P82Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-P82Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-P82Line ''
        Write-P82Line 'EXCEL SHUTDOWN'
        Write-P82Line '--------------'
        if ($naturalExit) {
            Write-P82Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                           [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'P8-2'
            Write-P82Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-P82Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:P82Residual)) {
            Write-P82Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-P82Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# VERDICT
# ===========================================================================
$null = Add-P82Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-P82Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-P82Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-P82Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-P82Check 'every COM release left 0 outstanding references' `
    (@($script:P82Residual).Count -eq 0) ((@($script:P82Residual)) -join '; ')
$null = Add-P82Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:P82Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:P82Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-P82Line ''
Write-P82Line 'VERDICT'
Write-P82Line '-------'
Write-P82Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
               [string]$failedPrereqs.Count + ' failed')
Write-P82Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
               [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-P82Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-P82Line ''
if ($ok) {
    Write-P82Line 'P8-2: PASS'
} else {
    Write-P82Line 'P8-2: FAIL'
    Write-P82Line ''
    Write-P82Line 'STOP AND REVIEW. Do not start P8-3 until this is understood.'
}
Write-P82Line ''
Write-P82Line ('report                 : ' + $script:P82Path)
Set-Content -LiteralPath $script:P82Path -Value ($script:P82Lines -join [Environment]::NewLine)
Write-Host ''
Write-Host ('The report is at ' + $script:P82Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
