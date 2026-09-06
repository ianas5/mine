<#
.SYNOPSIS
    PCCM Phase 8 - the MINIMAL P8-1 runner: the Results output surface, live.

.DESCRIPTION
    P8-1 BUILT A SHEET. This runs it.

    The Results surface presents the annual cash flow and the reconciliation,
    and it asks the Phase-7 handoff accessors what state the answer is in
    through four thin volatile adapters. Everything about that is source-proved
    and none of it is runtime-proved, and one question has to be answered before
    any of the rest is worth reading.

    THE FIRST ASSERTION IS THE RISK. Two of the four accessors reach
    `PCCM_SimulationStatus`, which re-derives the simulation status AND PERSISTS
    the two derived rows. Excel does not let a function called from a cell change
    the workbook. The ordinary behaviour is that such a write is ignored and the
    function returns its answer; if instead it raises, the adapter turns it into
    an error value and every state cell reads `#VALUE!`. So the very first thing
    this runner does, before a single simulation has run, is evaluate the four
    state cells and require that none of them is an Excel error.

    IF THAT FAILS, THE RUN STOPS THERE. Nothing is patched around it, nothing is
    suppressed, and no later part is attempted: the evidence that the design
    needs a non-writing semantic accessor is the point, and a runner that worked
    around it would destroy exactly that evidence.

    SIX PARTS, ONE SESSION, ONE DISPOSABLE WORKBOOK:

      0  EMPTY       before anything has run: four state cells, NOT PRODUCED,
                     no fabricated zero, no reconciliation that passes on blanks
      A  CURRENT     the accepted fixture through the accepted workflow, then
                     the whole surface against the persisted payload
      B  SELECTOR    P80 -> P50 with no rerun: CURRENT distributions, OTHER Px
                     profile, the old profile still visible and still labelled
                     P80, and a reconciliation that says whose it is
      C  RERUN       the annual step alone: the sheet follows the owner rather
                     than caching the selector it was built with
      D  REQUEST     iterations 1000 -> 1001, recalculation only: HISTORICAL,
                     which is the volatile adapters' whole reason to exist
      E  INVALID     a cost-line bound below its own minimum: HISTORICAL again,
                     with no simulation and no annual invoked

    NUMBERS ARE READ, NOT LOOKED AT. Every check reads `Value2` and `Formula`;
    the formatted `Text` is captured for the report and is never the authority
    for anything. A reconciliation verdict is checked against the unrounded
    stored values and the project's own identity rule.

    WHAT THIS RUNNER IS NOT. It re-tests no Phase-7 scenario, closes nothing,
    builds no dashboard and draws no chart. It owns the P8-1 output surface and
    that is all.

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
# FOUR OF THE TEN ARE ALSO CALLED HERE, exactly as W8 called them: the projected
# simulation request through `Get-NamedValue` and `Set-NamedValue`, and the
# cost-line register through `Get-TableBody` and `Set-TableCell`.
#
# They are copied BYTE FOR BYTE from `phase7_timing_scenarios.ps1` so there is
# one behaviour rather than two, and a source control pins them to it.
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
$script:P81Lines = New-Object System.Collections.ArrayList
$script:P81Path = ''
$script:P81Checks = New-Object System.Collections.ArrayList
$script:P81Residual = New-Object System.Collections.ArrayList

function Write-P81Line {
    param([string]$Text = '')
    $null = $script:P81Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:P81Path)) {
        try {
            Set-Content -LiteralPath $script:P81Path `
                -Value ($script:P81Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-P81Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:P81Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-P81Line $line
    return $Ok
}

function Invoke-P81Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:P81Residual.Add(
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

function Get-P81SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a P81 result could not be attributed to a source revision')
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

function Get-P81RowsInGroup {
    param($Inspection, [string]$Group)
    $out = @()
    $identity = $Inspection.sim_data.run_identity
    foreach ($key in $identity.rows.PSObject.Properties.Name) {
        if (([string]$identity.groups.$key) -ceq $Group) { $out += [string]$key }
    }
    return @($out)
}

function Get-P81RunInvariants {
    param($Workbook, $Inspection)
    $excluded = @(Get-P81RowsInGroup -Inspection $Inspection -Group 'derived')
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($excluded -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-P81InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-P81Check ($Label + ': no run identity, nonce, pending marker, digest or ' +
                         'simulation attempt row moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

function ConvertTo-P81ColumnNumber {
    param([string]$Letters)
    $number = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - 64)
    }
    return $number
}

function ConvertFrom-P81ColumnNumber {
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

function Get-P81AnnualStamp {
    param($Workbook, $Inspection, $P7, [string]$Bank)
    $stamp = $P7.annual_records.stamp
    $column = [string]$stamp.bank_value_columns.$Bank
    if ([string]::IsNullOrEmpty($column)) {
        throw ('the annual stamp projection carries no column for bank ' + [char]39 + $Bank + [char]39)
    }
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $stamp.rows.PSObject.Properties.Name) {
        $address = $column + [string]([int]$stamp.rows.$key)
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection -Address $address))
    }
    return $out
}

function Get-P81AnnualRecord {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$Offset)
    $records = $P7.annual_records
    $row = [int]$records.first_record_row + $Offset
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $records.index_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.index_columns.$Bank.$key
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address ($column + [string]$row)))
    }
    $count = [int]$records.quantile_count
    foreach ($measure in $records.quantile_first_column.$Bank.PSObject.Properties.Name) {
        $first = ConvertTo-P81ColumnNumber -Letters ([string]$records.quantile_first_column.$Bank.$measure)
        $ladder = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $count; $index++) {
            $column = ConvertFrom-P81ColumnNumber -Number ($first + $index)
            $null = $ladder.Add((Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
                -Address ($column + [string]$row)))
        }
        $out.Add(('ladder_' + $measure), @($ladder))
    }
    foreach ($measure in $records.selected_px_profile_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.selected_px_profile_columns.$Bank.$measure
        $out.Add(('profile_' + $measure), (Get-SimRawCell -Workbook $Workbook `
            -Inspection $Inspection -Address ($column + [string]$row)))
    }
    return $out
}

function Get-P81AnnualSurface {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $surface = New-Object System.Collections.Specialized.OrderedDictionary
    $stamp = Get-P81AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 -Bank $Bank
    foreach ($key in $stamp.Keys) { $surface.Add(('stamp.' + [string]$key), $stamp[$key]) }
    $ladderCount = [int]$P7.annual_records.quantile_count
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-P81AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        $year = 'year' + [string]($offset + 1)
        foreach ($key in $record.Keys) {
            $name = [string]$key
            if ($name.StartsWith('ladder_')) {
                $ladder = @($record[$name])
                for ($index = 0; $index -lt $ladderCount; $index++) {
                    $rung = $null
                    if ($index -lt $ladder.Count) { $rung = $ladder[$index] }
                    $surface.Add(($year + '.' + $name + '.' + [string]($index + 1)), $rung)
                }
            } else {
                $surface.Add(($year + '.' + $name), $record[$name])
            }
        }
    }
    return $surface
}

function Get-P81BankCapture {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $capture = New-Object System.Collections.Specialized.OrderedDictionary
    $block = Get-SimBankBlock -Workbook $Workbook -Inspection $Inspection -Bank $Bank
    foreach ($key in $block.Keys) { $capture.Add(('sim.' + [string]$key), $block[$key]) }
    $annual = Get-P81AnnualSurface -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank $Bank -YearCount $YearCount
    foreach ($key in $annual.Keys) { $capture.Add(('annual.' + [string]$key), $annual[$key]) }
    return $capture
}

function Compare-P81Surface {
    param($Before, $After)
    $moved = New-Object System.Collections.ArrayList
    foreach ($key in $Before.Keys) {
        $name = [string]$key
        if (-not (Test-SimSameValue -A $Before[$name] -B $After[$name])) {
            $null = $moved.Add($name + ': ' + (Format-SimValue $Before[$name]) + ' -> ' +
                               (Format-SimValue $After[$name]))
        }
        if ($moved.Count -gt 12) { break }
    }
    $detail = ($moved -join '; ')
    if ($moved.Count -gt 12) { $detail = $detail + ' ... (truncated)' }
    return [pscustomobject]@{ Moved = $moved.Count; Detail = $detail }
}

function Get-P81IterationBlock {
    param($Workbook, $Inspection, [string]$Bank, [int]$Count, $Ledger)
    $block = $Inspection.sim_data.iteration_records
    $first = [string]$block.banks.$Bank.iteration_index
    $last = [string]$block.banks.$Bank.total_pv
    $firstRow = [int]$block.first_iteration_row
    $lastRow = $firstRow + $Count - 1
    $sheets = $null; $sheet = $null; $range = $null
    $acquired = 0
    try {
        $sheets = $Workbook.Worksheets
        $acquired = $acquired + 1
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $acquired = $acquired + 1
        $range = $sheet.Range($first + [string]$firstRow + ':' + $last + [string]$lastRow)
        $acquired = $acquired + 1
        return [pscustomobject]@{ Values = $range.Value2; Acquired = $acquired }
    } finally {
        if ($null -ne $range)  { Invoke-P81Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-P81Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-P81Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}

function Compare-P81IterationGrid {
    param($Before, $After)
    if (($null -eq $Before) -or ($null -eq $After)) { return 'an iteration block was not read' }
    if (($Before.GetLength(0) -ne $After.GetLength(0)) -or
        ($Before.GetLength(1) -ne $After.GetLength(1))) {
        return 'the iteration block changed shape'
    }
    for ($row = 1; $row -le [int]$Before.GetLength(0); $row++) {
        for ($column = 1; $column -le [int]$Before.GetLength(1); $column++) {
            if (-not (Test-SimSameValue -A $Before[$row, $column] -B $After[$row, $column])) {
                return ('iteration row ' + [string]$row + ' column ' + [string]$column + ': ' +
                        (Format-SimValue $Before[$row, $column]) + ' -> ' +
                        (Format-SimValue $After[$row, $column]))
            }
        }
    }
    return ''
}

function Get-P81Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

function Invoke-P81Endpoint {
    param($Excel, [string]$Endpoint)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run($Endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

function Set-P81NamedText {
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

function Get-P81IdentityAllowance {
    param($Provenance, [double]$ConditioningScale)
    $floor = [double]$Provenance.identity_absolute_floor
    $coefficient = [double]$Provenance.identity_relative_coefficient
    $scaleFloor = [double]$Provenance.conditioning_scale_floor
    $scale = [Math]::Max($scaleFloor, [Math]::Abs($ConditioningScale))
    return [Math]::Max($floor, $coefficient * $scale)
}

function Get-P81Register {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-P81RegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}

function Get-P81RegisterRowIndex {
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
# READING A RESULTS CELL, AND KNOWING WHEN IT DID NOT EVALUATE
# ===========================================================================
# THREE FACTS PER CELL, AND THEY ARE NOT INTERCHANGEABLE.
#
#   Value2   the stored value. THE ONLY NUMERIC AUTHORITY here. An Excel error
#            arrives as an Int32 in the CVErr band, which is why the band is
#            named below rather than inferred from how the text looks.
#   Formula  what the cell was built from. A runtime run must be able to say the
#            state cell really calls the adapter, not merely that its answer
#            looked right this once.
#   Text     what a reader sees. Captured for the report and never compared
#            against a number: `#,##0` turns 4.5E-13 into `0`, and a check that
#            read that string would call a mismatch agreement.
$script:P81ErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}

function Get-P81Cell {
    param($Workbook, $Inspection, [string]$Address)
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sheet)
        $range = $sheet.Range($Address)
        $value = $range.Value2
        $text = ''
        try { $text = [string]$range.Text } catch { $text = '<unreadable>' }
        $formula = ''
        try { $formula = [string]$range.Formula } catch { $formula = '<unreadable>' }
        $errorName = ''
        if (($value -is [int]) -or ($value -is [long])) {
            $code = [int]$value
            if ($script:P81ErrorCodes.ContainsKey($code)) { $errorName = [string]$script:P81ErrorCodes[$code] }
        }
        # A CELL WHOSE TEXT IS AN ERROR NAME IS AN ERROR whatever its Value2
        # arrived as. Both routes are kept because neither is guaranteed on its
        # own, and this runner exists to settle a question about exactly this.
        if ([string]::IsNullOrEmpty($errorName)) {
            foreach ($name in $script:P81ErrorCodes.Values) {
                if (([string]$text) -ceq [string]$name) { $errorName = [string]$name }
            }
        }
        return [pscustomobject]@{
            Address = $Address; Value = $value; Text = $text
            Formula = $formula; ErrorName = $errorName
            IsError = (-not [string]::IsNullOrEmpty($errorName))
        }
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(Results)';    $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(Results)'; $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';         $sheets = $null }
    }
}

function Format-P81Cell {
    param($Cell)
    if ($null -eq $Cell) { return '<not read>' }
    if ($Cell.IsError) { return ([string]$Cell.ErrorName + ' (error)') }
    return ((Format-SimValue $Cell.Value) + ' shown as ' + [char]39 + [string]$Cell.Text + [char]39)
}

# THE FOUR STATE CELLS, TOGETHER, because they are only ever asked together and
# the first question about them is always whether all four evaluated.
function Get-P81StateCells {
    param($Workbook, $Inspection)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $Inspection.state.PSObject.Properties.Name) {
        $address = [string]$Inspection.columns.nominal + [string]([int]$Inspection.state.$key.row)
        $out.Add([string]$key, (Get-P81Cell -Workbook $Workbook -Inspection $Inspection -Address $address))
    }
    return $out
}

# THE ASSERTION THAT COMES BEFORE EVERY OTHER ASSERTION.
function Add-P81StateEvaluatedCheck {
    param($Cells, [string]$Stage)
    $broken = @()
    foreach ($key in $Cells.Keys) {
        if ($Cells[$key].IsError) {
            $broken += ([string]$key + ' = ' + [string]$Cells[$key].ErrorName)
        }
    }
    return (Add-P81Check ($Stage + ': every Results state cell evaluated without an Excel error') `
        ($broken.Count -eq 0) ($broken -join '; '))
}

# ORDINARY RECALCULATION, AND WHAT THE APPLICATION WAS SET TO WHEN IT HAPPENED.
# The endpoints put calculation into manual and restore it; a state that failed
# to update because the workbook was left in manual is a different finding from
# one that failed because the adapter did not fire, and the report must be able
# to tell them apart.
function Invoke-P81Recalculate {
    param($Excel, [string]$Stage)
    $mode = '<unreadable>'
    try { $mode = [string]$Excel.Calculation } catch { $mode = '<unreadable>' }
    $failure = ''
    try { $Excel.Calculate() } catch { $failure = (Format-Err $_) }
    Write-P81Line ('    recalculated at ' + $Stage + ' (Application.Calculation = ' + $mode + ')')
    return (Add-P81Check ($Stage + ': the workbook recalculated') `
        ([string]::IsNullOrWhiteSpace($failure)) ($failure + ' calculation mode ' + $mode) `
        'PREREQUISITE')
}

# ===========================================================================
# THE ANNUAL TABLE, AS THE SHEET SHOWS IT
# ===========================================================================
function Get-P81AnnualDisplayRow {
    param($Workbook, $Inspection, [int]$Offset)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    $row = [int]$Inspection.annual.first_row + $Offset
    foreach ($column in @($Inspection.annual.columns)) {
        $out.Add([string]$column.key, (Get-P81Cell -Workbook $Workbook -Inspection $Inspection `
            -Address ([string]$column.column + [string]$row)))
    }
    return $out
}

function Get-P81ReconciliationCells {
    param($Workbook, $Inspection, [string]$Measure)
    $column = [string]$Inspection.columns.nominal
    if ($Measure -eq 'pv') { $column = [string]$Inspection.columns.pv }
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $Inspection.reconciliation.rows.PSObject.Properties.Name) {
        $out.Add([string]$key, (Get-P81Cell -Workbook $Workbook -Inspection $Inspection `
            -Address ($column + [string]([int]$Inspection.reconciliation.rows.$key))))
    }
    return $out
}

# A DISPLAYED NUMBER AGAINST THE NUMBER IT CLAIMS TO BE. Exact, because a
# presentation cell that merely LOOKS UP a stored value has no arithmetic to
# lose: it is the same Double or it is the wrong cell.
function Test-P81SameNumber {
    param($Cell, $Expected)
    if ($null -eq $Cell) { return $false }
    if ($Cell.IsError) { return $false }
    if ($Expected -isnot [double]) { return $false }
    if ($Cell.Value -isnot [double]) { return $false }
    return ([double]$Cell.Value -eq [double]$Expected)
}

function Test-P81Blank {
    param($Cell)
    if ($null -eq $Cell) { return $false }
    if ($Cell.IsError) { return $false }
    return (Test-SimBlank -Value $Cell.Value)
}

# ===========================================================================
# THE THREE SURFACE CHECKS, PARAMETERISED BY WHAT THE STAGE EXPECTS
# ===========================================================================
# One implementation each, because every part asks the same questions of the
# same cells and only the expected answers differ. A second copy per part would
# be six chances to check six slightly different things.

# ===========================================================================
# THE OBSERVATION ORDER, AND WHY IT IS AN ORDER AND NOT A HABIT
# ===========================================================================
# THE DEFECT THIS SHAPE EXISTS TO PREVENT. The four Results cells are checked
# against the same accessors invoked directly through `Application.Run` - and
# that direct invocation runs in a VBA context, where the status derivation MAY
# legitimately persist the two derived rows. Doing it first would let the probe
# update `simulation_status` and then let the worksheet wrapper look correct on
# rows the probe had just written: the runner would have manufactured the very
# agreement it was sent to test, and the UDF/recalculation question would come
# back "fine" whatever the truth was.
#
# So every phase observes in one order, and one function owns it:
#
#   1  derived rows, BEFORE any recalculation
#   2  ordinary recalculation
#   3  derived rows, immediately AFTER it and before anything out-of-cell
#   4  the four Results cells, FROZEN - formula, Value2, Text, error class
#   5  every assertion that needs only those frozen values
#   6  and ONLY NOW the four accessors, directly
#   7  parity, against the values frozen at step 4 - never re-read
#   8  derived rows after the direct call, reported apart from step 3
#
# NOTHING RECALCULATES BETWEEN 6 AND 7. The frozen record is the observation;
# the direct call is a second, later question asked of the same workbook.

function Get-P81DerivedRows {
    param($Workbook, $Inspection)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in @(Get-P81RowsInGroup -Inspection $Inspection -Group 'derived')) {
        $out.Add([string]$key, (Get-SimField -Workbook $Workbook -Inspection $Inspection `
            -FieldKey $key))
    }
    return $out
}

# EVERY ASSERTION THAT NEEDS ONLY THE FROZEN CELLS. Not one of them may reach a
# procedure: if a check here called an accessor it would be step 6 wearing step
# 5's name.
function Invoke-P81FrozenStateChecks {
    param($Cells, $Inspection, [string]$Stage, [string]$Distribution,
          [string]$Profile, $ExpectedPx, $ExpectedYears)
    foreach ($key in $Cells.Keys) {
        Write-P81Line ('    ' + $Stage.PadRight(10) + ([string]$key).PadRight(20) +
                       (Format-P81Cell $Cells[$key]))
    }
    $null = Add-P81Check ($Stage + ': the annual distributions display ' + $Distribution) `
        (Test-SimExactText -Actual $Cells['distribution_state'].Value -Expected $Distribution) `
        (Format-P81Cell $Cells['distribution_state'])
    $null = Add-P81Check ($Stage + ': the annual profile displays ' + $Profile) `
        (Test-SimExactText -Actual $Cells['profile_state'].Value -Expected $Profile) `
        (Format-P81Cell $Cells['profile_state'])
    if ($null -eq $ExpectedPx) {
        $null = Add-P81Check ($Stage + ': no profile confidence level is shown') `
            (Test-P81Blank -Cell $Cells['profile_px']) (Format-P81Cell $Cells['profile_px'])
    } else {
        $null = Add-P81Check ($Stage + ': the profile confidence level shown is ' + [string]$ExpectedPx) `
            (Test-SimExactText -Actual $Cells['profile_px'].Value -Expected ([string]$ExpectedPx)) `
            (Format-P81Cell $Cells['profile_px'])
    }
    $null = Add-P81Check ($Stage + ': the year count shown is ' + [string]$ExpectedYears) `
        (Test-P81SameNumber -Cell $Cells['year_count'] -Expected ([double]$ExpectedYears)) `
        (Format-P81Cell $Cells['year_count'])
    # AND THE CELL IS STILL A CALL. A right answer from a cell somebody had
    # since typed over would look identical to a right answer from the adapter.
    $wrong = @()
    foreach ($key in $Cells.Keys) {
        $expected = '=' + [string]$Inspection.state.$key.procedure + '()'
        if (([string]$Cells[$key].Formula) -cne $expected) {
            $wrong += ([string]$key + ': ' + [string]$Cells[$key].Formula)
        }
    }
    $null = Add-P81Check ($Stage + ': every state cell is still a call to its adapter') `
        ($wrong.Count -eq 0) ($wrong -join '; ')
}

# STEP 6 AND STEP 7, TOGETHER AND NOWHERE ELSE. The cells are not re-read: the
# comparison is against the record frozen before this function was entered.
function Invoke-P81AccessorParity {
    param($Excel, $Frozen, $Inspection, $P7, [string]$Stage)
    $direct = Get-P81Handoff -Excel $Excel -P7 $P7
    $accessors = @($P7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $disagreed = @()
    $index = 0
    foreach ($key in $Frozen.Keys) {
        $accessor = [string]$accessors[$index]
        $index = $index + 1
        if ([string]$Inspection.state.$key.accessor -cne $accessor) {
            $disagreed += ([string]$key + ': the projection pairs it with ' +
                           [string]$Inspection.state.$key.accessor + ', not ' + $accessor)
            continue
        }
        $shown = $Frozen[$key].Value
        $said = $direct[$accessor]
        # A count arrives as a Long from VBA and as a Double from a cell; the
        # comparison is on the VALUE, and a type difference across that boundary
        # is not a disagreement about the state.
        $same = $false
        if (($shown -is [string]) -or ($said -is [string])) {
            $same = (([string]$shown) -ceq ([string]$said))
        } elseif ($null -eq $shown) {
            $same = ($null -eq $said)
        } else {
            $same = ([double]$shown -eq [double]$said)
        }
        if (-not $same) {
            $disagreed += ([string]$key + ': frozen cell ' + (Format-P81Cell $Frozen[$key]) +
                           ' vs ' + $accessor + ' ' + (Format-SimValue $said))
        }
    }
    $null = Add-P81Check ($Stage + ': every frozen state cell shows what its accessor says') `
        ($disagreed.Count -eq 0) ($disagreed -join '; ')
    return $direct
}

# THE WHOLE OBSERVATION, IN ITS ONE ORDER. Every phase goes through here; a
# phase that assembled the steps for itself would be free to assemble them
# wrongly, which is exactly what happened before this function existed.
function Invoke-P81Observation {
    param($Excel, $Workbook, $Inspection, $SimInspection, $P7, [string]$Stage,
          [string]$Distribution, [string]$Profile, $ExpectedPx, $ExpectedYears,
          [switch]$GateOnly)
    # (1) and (2)
    $beforeRecalc = Get-P81DerivedRows -Workbook $Workbook -Inspection $SimInspection
    $null = Invoke-P81Recalculate -Excel $Excel -Stage $Stage
    # (3) BEFORE ANYTHING OUT-OF-CELL. This is the only capture that can be
    # attributed to the recalculation, and it is taken while that is still true.
    $afterRecalc = Get-P81DerivedRows -Workbook $Workbook -Inspection $SimInspection
    # (4)
    $frozen = Get-P81StateCells -Workbook $Workbook -Inspection $Inspection
    $evaluated = Add-P81StateEvaluatedCheck -Cells $frozen -Stage $Stage
    # (5)
    if ($evaluated -and (-not $GateOnly)) {
        Invoke-P81FrozenStateChecks -Cells $frozen -Inspection $Inspection -Stage $Stage `
            -Distribution $Distribution -Profile $Profile -ExpectedPx $ExpectedPx `
            -ExpectedYears $ExpectedYears
    }
    # (6) and (7)
    $direct = $null
    $afterDirect = $afterRecalc
    if ($evaluated) {
        $direct = Invoke-P81AccessorParity -Excel $Excel -Frozen $frozen `
            -Inspection $Inspection -P7 $P7 -Stage $Stage
        # (8) REPORTED APART. A row that moved here moved because this runner
        # asked, not because the workbook recalculated.
        $afterDirect = Get-P81DerivedRows -Workbook $Workbook -Inspection $SimInspection
    }
    Write-P81Line ('    ' + $Stage + ': THE DERIVED STATUS ROWS, IN OBSERVATION ORDER')
    $movedByRecalc = 0
    $movedByDirect = 0
    foreach ($key in $beforeRecalc.Keys) {
        $recalcVerdict = 'unchanged'
        if (-not (Test-SimSameValue -A $beforeRecalc[[string]$key] -B $afterRecalc[[string]$key])) {
            $recalcVerdict = 'updated'
            $movedByRecalc = $movedByRecalc + 1
        }
        $directVerdict = 'unchanged'
        if (-not (Test-SimSameValue -A $afterRecalc[[string]$key] -B $afterDirect[[string]$key])) {
            $directVerdict = 'updated'
            $movedByDirect = $movedByDirect + 1
        }
        Write-P81Line ('      ' + ([string]$key).PadRight(24) +
                       'recalc: ' + $recalcVerdict.PadRight(11) +
                       'direct: ' + $directVerdict.PadRight(11) +
                       (Format-SimValue $beforeRecalc[[string]$key]) + ' -> ' +
                       (Format-SimValue $afterRecalc[[string]$key]) + ' -> ' +
                       (Format-SimValue $afterDirect[[string]$key]))
    }
    Write-P81Line ('      OBSERVED at ' + $Stage + ': ' + [string]$movedByRecalc +
                   ' derived row(s) moved during the recalculation, ' + [string]$movedByDirect +
                   ' more when the accessors were then called directly.')
    return [pscustomobject]@{
        Cells = $frozen; Evaluated = $evaluated; Direct = $direct
        BeforeRecalc = $beforeRecalc; AfterRecalc = $afterRecalc; AfterDirect = $afterDirect
    }
}


# THE TABLE AGAINST THE PAYLOAD IT CLAIMS TO SHOW, cell by cell, plus the first
# row past the answer. A sheet that showed the right four years and left a fifth
# behind would pass every check that only looked at the four.
function Invoke-P81AnnualChecks {
    param($Workbook, $Inspection, $SimInspection, $P7, [string]$Bank, [int]$YearCount,
          [int]$StartYear, [string]$Stage)
    $problems = New-Object System.Collections.ArrayList
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $shown = Get-P81AnnualDisplayRow -Workbook $Workbook -Inspection $Inspection -Offset $offset
        $stored = Get-P81AnnualRecord -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        foreach ($column in @($Inspection.annual.columns)) {
            $key = [string]$column.key
            $storedKey = [string]$column.field
            if (([string]$column.source) -ceq 'profile') { $storedKey = 'profile_' + $storedKey }
            if (-not (Test-P81SameNumber -Cell $shown[$key] -Expected $stored[$storedKey])) {
                $null = $problems.Add('year ' + [string]($offset + 1) + ' ' + $key + ': ' +
                                      (Format-P81Cell $shown[$key]) + ' vs stored ' +
                                      (Format-SimValue $stored[$storedKey]))
            }
        }
        # THE CALENDAR AXIS IS THE FIXTURE'S, checked against the model rather
        # than against the same cell that produced it.
        if (-not (Test-P81SameNumber -Cell $shown['calendar_year'] `
                -Expected ([double]($StartYear + $offset)))) {
            $null = $problems.Add('year ' + [string]($offset + 1) + ' calendar year: ' +
                                  (Format-P81Cell $shown['calendar_year']))
        }
        if ($problems.Count -gt 12) { break }
    }
    $null = Add-P81Check ($Stage + ': the ' + [string]$YearCount +
                          ' displayed years are the persisted annual payload') `
        ($problems.Count -eq 0) (($problems -join '; '))

    # PAST THE ANSWER. Two rows, because one could be an accident of layout.
    $residue = New-Object System.Collections.ArrayList
    for ($offset = $YearCount; $offset -lt [Math]::Min($YearCount + 2, [int]$Inspection.annual.row_window); $offset++) {
        $shown = Get-P81AnnualDisplayRow -Workbook $Workbook -Inspection $Inspection -Offset $offset
        foreach ($column in @($Inspection.annual.columns)) {
            $key = [string]$column.key
            if (-not (Test-P81Blank -Cell $shown[$key])) {
                $null = $residue.Add('row ' + [string]($offset + 1) + ' ' + $key + ' = ' +
                                     (Format-P81Cell $shown[$key]))
            }
        }
    }
    $null = Add-P81Check ($Stage + ': the rows past the stamped year count show nothing') `
        ($residue.Count -eq 0) (($residue -join '; '))
    # AND THEY ARE BLANK BECAUSE A FORMULA CHOSE TO BE, not because nobody wrote
    # one: an empty cell and a formula that returns "" are different facts, and
    # only the second survives a longer answer being published later.
    $lastRow = Get-P81AnnualDisplayRow -Workbook $Workbook -Inspection $Inspection `
        -Offset ([int]$Inspection.annual.row_window - 1)
    $missing = @()
    foreach ($column in @($Inspection.annual.columns)) {
        if (-not ([string]$lastRow[[string]$column.key].Formula).StartsWith('=')) {
            $missing += [string]$column.key
        }
    }
    $null = Add-P81Check ($Stage + ': the last row of the structural window is still formula-driven') `
        ($missing.Count -eq 0) ($missing -join ', ')
}

# THE RECONCILIATION, RE-DERIVED AND COMPARED. The sheet's own answer is checked
# against the authoritative total, the persisted profile and the project's
# identity rule - never against the formatted string beside it.
function Invoke-P81ReconciliationChecks {
    param($Workbook, $Inspection, $SimInspection, $P7, [string]$Bank, [int]$YearCount,
          [string]$Label, [int]$LadderIndex, $ProfileState, [string]$Stage)
    $rowKey = 'quantile_' + [string]($LadderIndex + 1)
    $contingencyBlock = $SimInspection.sim_data.contingency_ladder
    $measures = @($P7.summary_semantics.contingency_measures | ForEach-Object { [string]$_ })
    foreach ($measure in $measures) {
        $cells = Get-P81ReconciliationCells -Workbook $Workbook -Inspection $Inspection -Measure $measure
        $broken = @()
        foreach ($key in $cells.Keys) { if ($cells[$key].IsError) { $broken += [string]$key } }
        $null = Add-P81Check ($Stage + ': the ' + $measure +
                              ' reconciliation cells evaluated without an Excel error') `
            ($broken.Count -eq 0) ($broken -join ', ')

        # (A) THE TOTAL IS THE TOTAL. The summary block's rung at the selected
        # label - and demonstrably NOT the contingency block's rung, which is
        # the same money at the same Px, smaller by exactly the deterministic
        # base. W5 read the second as the first and failed four reconciliations
        # by that amount.
        $total = Get-SimSummaryValue -Workbook $Workbook -Inspection $SimInspection `
            -Bank $Bank -Measure $measure -RowKey $rowKey
        $contingency = Get-SimRawCell -Workbook $Workbook -Inspection $SimInspection `
            -Address ([string]$contingencyBlock.bank_value_columns.$Bank.$measure +
                      [string]([int]$contingencyBlock.rows.$rowKey))
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure + ' TOTAL is the ' +
                              $Label + ' total percentile') `
            (Test-P81SameNumber -Cell $cells['total'] -Expected $total) `
            ((Format-P81Cell $cells['total']) + ' vs published total ' + (Format-SimValue $total))
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure +
                              ' TOTAL is not the contingency at the same rung') `
            ((-not (Test-P81SameNumber -Cell $cells['total'] -Expected $contingency)) -and
             ($contingency -is [double]) -and ($total -is [double]) -and
             ([double]$contingency -ne [double]$total)) `
            ('total ' + (Format-SimValue $total) + ', contingency ' + (Format-SimValue $contingency))

        # (B) THE SUM IS THE PERSISTED PROFILE'S.
        $sum = 0.0
        $scale = 0.0
        $numeric = $true
        for ($offset = 0; $offset -lt $YearCount; $offset++) {
            $stored = Get-P81AnnualRecord -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
                -Bank $Bank -Offset $offset
            $value = $stored[('profile_' + $measure)]
            if ($value -isnot [double]) { $numeric = $false; continue }
            $sum = $sum + [double]$value
            $scale = $scale + [Math]::Abs([double]$value)
        }
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure +
                              ' profile sum is the sum of the persisted profile') `
            ($numeric -and (Test-P81SameNumber -Cell $cells['profile_sum'] -Expected $sum)) `
            ((Format-P81Cell $cells['profile_sum']) + ' vs summed ' + [string]$sum)

        # (C) THE DELTA AND THE ALLOWANCE, on the project's own rule.
        $expectedDelta = [double]::NaN
        if ($total -is [double]) { $expectedDelta = [double]$total - $sum }
        $conditioning = $scale
        if ($total -is [double]) { $conditioning = $conditioning + [Math]::Abs([double]$total) }
        $allowance = Get-P81IdentityAllowance -Provenance $Inspection.identity `
            -ConditioningScale $conditioning
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure + ' difference is total minus sum') `
            (Test-P81SameNumber -Cell $cells['difference'] -Expected $expectedDelta) `
            ((Format-P81Cell $cells['difference']) + ' vs ' + [string]$expectedDelta)
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure +
                              ' allowance is the accepted identity rule') `
            (Test-P81SameNumber -Cell $cells['allowance'] -Expected $allowance) `
            ((Format-P81Cell $cells['allowance']) + ' vs ' + [string]$allowance +
             ' (conditioning scale ' + [string]$conditioning + ')')

        # (D) THE VERDICT, TAKEN THE WAY THE SHEET SHOULD HAVE TAKEN IT - on the
        # stored values, not on anything rounded for a reader.
        $shouldReconcile = $false
        if (($cells['difference'].Value -is [double]) -and ($cells['allowance'].Value -is [double])) {
            $shouldReconcile = ([Math]::Abs([double]$cells['difference'].Value) -le
                                [double]$cells['allowance'].Value)
        }
        $verdict = [string]$Inspection.reconciliation.verdicts.mismatch
        if ($shouldReconcile) { $verdict = [string]$Inspection.reconciliation.verdicts.reconciled }
        $status = ''
        if (-not $cells['status'].IsError) { $status = [string]$cells['status'].Value }
        $null = Add-P81Check ($Stage + ': the displayed ' + $measure +
                              ' status agrees with the unrounded comparison') `
            ($status.StartsWith($verdict)) ((Format-P81Cell $cells['status']) +
                                            ', unrounded verdict ' + $verdict)
        # (E) AND IT SAYS WHOSE PROFILE IT IS. A reconciliation of a profile the
        # selector is no longer asking for must not read as the current answer.
        $currentProfile = [string]$P7.handoff.profile_states[1]
        $qualifier = [string]$Inspection.reconciliation.verdicts.qualifier_suffix
        if (([string]$ProfileState) -ceq $currentProfile) {
            $null = Add-P81Check ($Stage + ': the ' + $measure +
                                  ' status carries no historical qualifier') `
                (-not $status.Contains($qualifier)) (Format-P81Cell $cells['status'])
        } else {
            $null = Add-P81Check ($Stage + ': the ' + $measure + ' status says it is the ' +
                                  [string]$ProfileState + ' profile, not the current answer') `
                ($status.Contains($qualifier) -and $status.Contains([string]$ProfileState)) `
                (Format-P81Cell $cells['status'])
        }
    }
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 8 P8-1 (the Results output surface, live)' -ForegroundColor Cyan
Write-Host '=====================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$p8InspectPath  = Join-Path $BuildDir 'phase8_results_inspection.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath)) {
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

# THE ACCEPTED BEHAVIOURAL FIXTURE, unchanged. P8-1 presents an answer; it does
# not get to choose a friendlier one to present.
$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no behavioural scenario for P8-1 to reuse.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-P81SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a P8-1 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase8-p1-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath, $p8InspectPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:P81Path = Join-Path $tempRoot 'phase8_p1_results_surface.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$model = $case.model
$yearCount = [int]$model.timeline.duration
$startYear = [int]$model.timeline.start_year
$driverCount = @($model.cost_lines).Count + @($model.risks).Count
$iterations = [int]$case.iterations
$suppliedSeed = [double]$case.supplied_seed
$firstLabel = [string]$case.selected_confidence_level
$secondLabel = [string]$case.second_confidence_level
$quantileLabels = @($gateBCases.vocabulary.quantile_labels | ForEach-Object { [string]$_ })
$firstIndex = [array]::IndexOf($quantileLabels, $firstLabel)
$secondIndex = [array]::IndexOf($quantileLabels, $secondLabel)
$notProduced = [string]$p7.handoff.distribution_states[0]
$annualCurrent = [string]$p7.handoff.distribution_states[1]
$profileCurrent = [string]$p7.handoff.profile_states[1]
$otherPx = [string]$p7.handoff.inconsistent_stamp_state
$historical = [string]$p7.handoff.distribution_states[$p7.handoff.distribution_states.Count - 1]
# THE MODEL'S OWN STATE VOCABULARY. `calc_contract.yaml` owns the two
# orthogonal calculation axes and P8-1's projection of them arrived with W8;
# reading the simulation's CURRENT for a model status would be the axis-merging
# the contract forbids, and typing the word would be a second authority.
$calcCurrent = [string]$p7.model_states.derived_status[1]
$minimumIterations = [int]$gateBCases.bounds.business_minimum_iterations
$staleIterations = $minimumIterations + 1
$costRegister = Get-P81Register -Manifest $manifest -Key 'cost_lines'
$victim = @($model.cost_lines)[0]
$victimId = [string]$victim.permanent_id
$invalidMaximum = [double]$victim.min_value - 1.0
$maximumOrdinal = 0
try { $maximumOrdinal = Get-P81RegisterColumnIndex -Register $costRegister -ColumnKey 'unit_cost_max' }
catch { $maximumOrdinal = 0 }

Write-P81Line 'PCCM - PHASE 8 P8-1: THE RESULTS OUTPUT SURFACE, LIVE'
Write-P81Line '===================================================='
Write-P81Line ''
Write-P81Line 'This is the MINIMAL P8-1 runner. It owns the Results output surface and'
Write-P81Line 'nothing else: no Phase-7 scenario is re-tested, no dashboard is built and'
Write-P81Line 'no chart is drawn.'
Write-P81Line ''
Write-P81Line 'THE FIRST ASSERTION IS THE RISK. Two of the four handoff accessors reach the'
Write-P81Line 'simulation-status derivation, which PERSISTS two rows, and Excel does not let'
Write-P81Line 'a function called from a cell change the workbook. If the four state cells do'
Write-P81Line 'not evaluate, this run STOPS: nothing is patched around it and nothing is'
Write-P81Line 'suppressed, because that evidence is the point.'
Write-P81Line ''
Write-P81Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-P81Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-P81Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-P81Line ('git HEAD               : ' + [string]$revision.Head)
Write-P81Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-P81Line ('model version          : ' + [string]$manifest.model_version)
Write-P81Line ('build directory        : ' + $BuildDir)
Write-P81Line ('working copy           : ' + $tempRoot)
$artefacts = @(
    [pscustomobject]@{ Label = 'Stage-A workbook'
                       Path = (Join-Path $tempRoot ([string]$manifest.stage_a_filename)) },
    [pscustomobject]@{ Label = 'Stage-B workbook'; Path = $stageBPath }
)
foreach ($artefact in $artefacts) {
    $shown = '(not present)'
    if (Test-Path -LiteralPath ([string]$artefact.Path)) {
        try { $shown = [string](Get-FileHash -LiteralPath ([string]$artefact.Path) -Algorithm SHA256).Hash }
        catch { $shown = '(unreadable)' }
    }
    Write-P81Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-P81Line ''
Write-P81Line 'THE SIX PARTS'
Write-P81Line '-------------'
Write-P81Line ('  0  empty      the four state cells before anything has run')
Write-P81Line ('  A  current    ' + [string]$driverCount + ' drivers, ' + [string]$yearCount +
               ' years, FIXED seed ' + [string]$suppliedSeed + ', ' + [string]$iterations +
               ' iterations, ' + $firstLabel)
Write-P81Line ('  B  selector   ' + $firstLabel + ' -> ' + $secondLabel + ', recalculation only')
Write-P81Line ('  C  rerun      the annual endpoint alone, at ' + $secondLabel)
Write-P81Line ('  D  request    iterations ' + [string]$minimumIterations + ' -> ' +
               [string]$staleIterations + ', recalculation only')
Write-P81Line ('  E  invalid    cost line ' + $victimId + ' maximum below its own minimum')
Write-P81Line ''
Write-P81Line ('  identity rule        : ' + [string]$p8.identity.identity_rule)
Write-P81Line ('  annual window        : rows ' + [string]$p8.annual.first_row + '..' +
               [string]([int]$p8.annual.first_row + [int]$p8.annual.row_window - 1))
Write-P81Line ''
Write-P81Line 'EVERY ADDRESS IS PROJECTED. Not one row, column or procedure name is written'
Write-P81Line 'in this runner: they are read from phase8_results_inspection.json, which is'
Write-P81Line 'projected from workbook.yaml and the accepted contracts.'
Write-P81Line ''

$null = Add-P81Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-P81Line ''
    Write-P81Line 'STOP. Excel was never started for the P8-1 session, and nothing was accepted.'
    Write-P81Line ('report                 : ' + $script:P81Path)
    Write-P81Line 'P8-1: FAIL'
    exit 1
}

$null = Add-P81Check 'the P8-1 fixture is the accepted behavioural one' `
    (($driverCount -eq 5) -and ($yearCount -eq 4) -and (([string]$case.seed_mode) -ceq 'FIXED') -and
     ($iterations -eq $minimumIterations) -and ($firstIndex -ge 0) -and ($secondIndex -ge 0) -and
     ($firstLabel -cne $secondLabel)) `
    ([string]$driverCount + ' drivers, ' + [string]$yearCount + ' years, ' +
     [string]$case.seed_mode + ' seed ' + [string]$suppliedSeed + ', ' + [string]$iterations +
     ' iterations, ' + $firstLabel + ' -> ' + $secondLabel) 'PREREQUISITE'
$null = Add-P81Check 'the projection names an adapter for every handoff accessor' `
    ((@($p8.state.PSObject.Properties.Name).Count -eq 4) -and
     (@($p7.command_surface.handoff_accessors).Count -eq 4)) `
    ((@($p8.state.PSObject.Properties.Name) -join ', ')) 'PREREQUISITE'
$null = Add-P81Check 'the projection keeps the total and the contingency on different rows' `
    (([int]$p8.selected.total_row) -ne ([int]$p8.selected.contingency_row)) `
    ('total row ' + [string]$p8.selected.total_row + ', contingency row ' +
     [string]$p8.selected.contingency_row) 'PREREQUISITE'
$null = Add-P81Check 'the invalidating edit puts a cost line maximum below its own minimum' `
    (($invalidMaximum -lt [double]$victim.min_value) -and ($maximumOrdinal -ge 1)) `
    ($victimId + ': min ' + [string]$victim.min_value + ' -> max ' + [string]$invalidMaximum) `
    'PREREQUISITE'

# ===========================================================================
# THE P8-1 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$stoppedOnStateEvaluation = $false
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 8 P8-1 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-P81Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                   [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-P81Line 'ever terminated, and the workbook is never saved.'
    Write-P81Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-P81Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }

    # ===================================================================
    # PART 0 - EMPTY, AND THE ONE QUESTION THAT COMES FIRST
    # ===================================================================
    Write-P81Line ''
    Write-P81Line 'PART 0 - NOTHING HAS RUN'
    Write-P81Line '------------------------'
    # THE SAME ORDER AS EVERY OTHER PHASE, and the gate is inside it: the cells
    # are frozen before anything is asked out of a cell, so a failure here is a
    # failure of the worksheet and not of the order it was observed in.
    $observation0 = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part 0' -Distribution $notProduced `
        -Profile $notProduced -ExpectedPx $null -ExpectedYears 0
    if (-not $observation0.Evaluated) {
        $stoppedOnStateEvaluation = $true
        Write-P81Line ''
        Write-P81Line 'STOP. THE FOUR RESULTS STATE CELLS DID NOT EVALUATE.'
        Write-P81Line ''
        Write-P81Line 'This is the runtime evidence the run was prepared to collect. A worksheet'
        Write-P81Line 'function may not change the workbook, and two of the four handoff accessors'
        Write-P81Line 'reach a status derivation that persists two rows. Nothing here will be'
        Write-P81Line 'patched around it and no later part will be attempted: the decision about'
        Write-P81Line 'whether a non-writing semantic accessor is required is not this runner' + [char]39 + 's.'
        Write-P81Line ''
        foreach ($key in $observation0.Cells.Keys) {
            Write-P81Line ('    ' + ([string]$key).PadRight(20) +
                           [string]$p8.state.$key.procedure + '()  ->  ' +
                           (Format-P81Cell $observation0.Cells[$key]))
        }
        # AND ONE DIAGNOSTIC, AFTER THE OBSERVATION IS ALREADY FROZEN AND FAILED.
        # No parity is claimed and no check is added: whether the same accessors
        # answer normally OUTSIDE a cell is the single most useful fact for
        # deciding what to do about this, and taking it now cannot change what
        # was already recorded above.
        Write-P81Line ''
        Write-P81Line '    DIAGNOSTIC (not a check): the same accessors, called outside a cell'
        try {
            $diagnostic = Get-P81Handoff -Excel $excel -P7 $p7
            foreach ($accessor in @($p7.command_surface.handoff_accessors)) {
                Write-P81Line ('      ' + ([string]$accessor).PadRight(34) +
                               (Format-SimValue $diagnostic[[string]$accessor]))
            }
        } catch {
            Write-P81Line ('      the accessors also failed outside a cell: ' + (Format-Err $_))
        }
        throw 'the Results state cells did not evaluate; P8-1 stopped at part 0'
    }

    # NO FABRICATED ZERO. An empty cash flow is not a cash flow of zeros: a
    # zero would sum, reconcile and one day chart exactly like a real one.
    $fabricated = New-Object System.Collections.ArrayList
    foreach ($offset in @(0, 1, [int]$p8.annual.row_window - 1)) {
        $shown = Get-P81AnnualDisplayRow -Workbook $wb -Inspection $p8 -Offset $offset
        foreach ($column in @($p8.annual.columns)) {
            if (-not (Test-P81Blank -Cell $shown[[string]$column.key])) {
                $null = $fabricated.Add('row ' + [string]($offset + 1) + ' ' +
                                        [string]$column.key + ' = ' +
                                        (Format-P81Cell $shown[[string]$column.key]))
            }
        }
    }
    $null = Add-P81Check 'part 0: the annual table shows nothing, not zeros' `
        ($fabricated.Count -eq 0) (($fabricated -join '; '))

    # AND NO RECONCILIATION PASSES ON BLANKS. "Reconciled" from two empty cells
    # is the most convincing wrong answer this sheet could give.
    $falsePass = New-Object System.Collections.ArrayList
    foreach ($measure in @('nominal', 'pv')) {
        $cells = Get-P81ReconciliationCells -Workbook $wb -Inspection $p8 -Measure $measure
        foreach ($key in $cells.Keys) {
            if ($cells[$key].IsError) {
                $null = $falsePass.Add($measure + ' ' + [string]$key + ' = ' +
                                       [string]$cells[$key].ErrorName)
                continue
            }
            if (-not (Test-SimBlank -Value $cells[$key].Value)) {
                $null = $falsePass.Add($measure + ' ' + [string]$key + ' = ' +
                                       (Format-P81Cell $cells[$key]))
            }
        }
    }
    $null = Add-P81Check 'part 0: the reconciliation reports nothing at all, and passes nothing' `
        ($falsePass.Count -eq 0) (($falsePass -join '; '))

    # THE SECTIONS EXIST WHERE THE PROJECTION SAYS, and the table is formulas.
    $layout = New-Object System.Collections.ArrayList
    foreach ($key in $p8.state.PSObject.Properties.Name) {
        $labelCell = Get-P81Cell -Workbook $wb -Inspection $p8 `
            -Address ([string]$p8.columns.label + [string]([int]$p8.state.$key.row))
        if (-not (Test-SimExactText -Actual $labelCell.Value -Expected ([string]$p8.state.$key.label))) {
            $null = $layout.Add([string]$key + ' label ' + (Format-P81Cell $labelCell))
        }
    }
    foreach ($column in @($p8.annual.columns)) {
        $headerCell = Get-P81Cell -Workbook $wb -Inspection $p8 `
            -Address ([string]$column.column + [string]([int]$p8.annual.header_row))
        if (-not (Test-SimExactText -Actual $headerCell.Value -Expected ([string]$column.header))) {
            $null = $layout.Add('header ' + [string]$column.key + ' ' + (Format-P81Cell $headerCell))
        }
    }
    $null = Add-P81Check 'part 0: the P8-1 sections are materialised where the projection says' `
        ($layout.Count -eq 0) (($layout -join '; '))

    # ===================================================================
    # PART A - THE SUCCESSFUL CURRENT RESULT
    # ===================================================================
    Write-P81Line ''
    Write-P81Line 'PART A - A SUCCESSFUL ANNUAL RESULT'
    Write-P81Line '-----------------------------------'
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
        Set-P81NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
            -Value $firstLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-P81Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the P8-1 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'P8-1 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-P81Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-P81Check 'PCCM_RunSimulation succeeded' `
        (Test-Phase6Announced -Result $simResult -Kind 'OK') $simResult 'PREREQUISITE'
    if (-not (Test-Phase6Announced -Result $simResult -Kind 'OK')) {
        throw 'the P8-1 simulation did not succeed'
    }
    $annual = Invoke-P81Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P81Check ([string]$p7.command_surface.annual_endpoint + ' succeeded') `
        ($annual -like 'OK|*') $annual 'PREREQUISITE'
    if (-not ($annual -like 'OK|*')) { throw ('the P8-1 annual run did not succeed: ' + $annual) }

    $stateA = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bank = Get-Phase6ActiveBank -State $stateA
    $invariantsA = Get-P81RunInvariants -Workbook $wb -Inspection $simInspection
    $captureA = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $iterationsA = Get-P81IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsA.Acquired

    $observationA = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part A' -Distribution $annualCurrent `
        -Profile $profileCurrent -ExpectedPx $firstLabel -ExpectedYears $yearCount
    $null = $observationA
    Invoke-P81AnnualChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -StartYear $startYear -Stage 'part A'
    Invoke-P81ReconciliationChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection `
        -P7 $p7 -Bank $bank -YearCount $yearCount -Label $firstLabel -LadderIndex $firstIndex `
        -ProfileState $profileCurrent -Stage 'part A'

    # ===================================================================
    # PART B - THE SELECTOR MOVES AND NOTHING ELSE DOES
    # ===================================================================
    Write-P81Line ''
    Write-P81Line ('PART B - ' + $firstLabel + ' -> ' + $secondLabel + ', RECALCULATION ONLY')
    Write-P81Line '--------------------------------------------'
    Set-P81NamedText -Workbook $wb `
        -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
        -Value $secondLabel
    $observationB = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part B' -Distribution $annualCurrent `
        -Profile $otherPx -ExpectedPx $firstLabel -ExpectedYears $yearCount
    $null = $observationB
    # THE PERSISTED ANSWER IS STILL THE ONE ON THE SHEET, and it is still the
    # P80 one. A presentation layer that relabelled it would be claiming a blend
    # nobody computed.
    Invoke-P81AnnualChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -StartYear $startYear -Stage 'part B'
    Invoke-P81ReconciliationChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection `
        -P7 $p7 -Bank $bank -YearCount $yearCount -Label $firstLabel -LadderIndex $firstIndex `
        -ProfileState $otherPx -Stage 'part B'
    $captureB = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $movedB = Compare-P81Surface -Before $captureA -After $captureB
    $null = Add-P81Check 'part B: the persisted run and annual payload are value-identical' `
        ($movedB.Moved -eq 0) $movedB.Detail
    $null = Add-P81InvariantChecks 'part B: a presentation recalculation' $invariantsA `
        (Get-P81RunInvariants -Workbook $wb -Inspection $simInspection)
    $iterationsB = Get-P81IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsB.Acquired
    $gridB = Compare-P81IterationGrid -Before $iterationsA.Values -After $iterationsB.Values
    $null = Add-P81Check 'part B: no published iteration value moved' `
        ([string]::IsNullOrWhiteSpace($gridB)) $gridB

    # ===================================================================
    # PART C - THE ANNUAL STEP ALONE
    # ===================================================================
    Write-P81Line ''
    Write-P81Line ('PART C - THE ANNUAL ENDPOINT ALONE, AT ' + $secondLabel)
    Write-P81Line '----------------------------------------'
    $rerun = Invoke-P81Endpoint -Excel $excel `
        -Endpoint ([string]$p7.command_surface.annual_endpoint)
    $null = Add-P81Check ('part C: ' + [string]$p7.command_surface.annual_endpoint +
                          ' succeeded on the moved selector') ($rerun -like 'OK|*') $rerun
    $observationC = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part C' -Distribution $annualCurrent `
        -Profile $profileCurrent -ExpectedPx $secondLabel -ExpectedYears $yearCount
    $null = $observationC
    Invoke-P81AnnualChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -StartYear $startYear -Stage 'part C'
    Invoke-P81ReconciliationChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection `
        -P7 $p7 -Bank $bank -YearCount $yearCount -Label $secondLabel -LadderIndex $secondIndex `
        -ProfileState $profileCurrent -Stage 'part C'
    # THE PROFILE REALLY CHANGED. If the sheet had cached the P80 answer it
    # would still reconcile - against the wrong total - so the payload itself is
    # required to have moved while the run identity did not.
    $captureC = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $movedC = Compare-P81Surface -Before $captureB -After $captureC
    $null = Add-P81Check 'part C: the annual rerun republished the profile' `
        ($movedC.Moved -gt 0) ('changed keys: ' + $movedC.Detail)
    $null = Add-P81InvariantChecks 'part C: the annual rerun' $invariantsA `
        (Get-P81RunInvariants -Workbook $wb -Inspection $simInspection)
    $iterationsC = Get-P81IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsC.Acquired
    $gridC = Compare-P81IterationGrid -Before $iterationsA.Values -After $iterationsC.Values
    $null = Add-P81Check 'part C: no published iteration value moved' `
        ([string]::IsNullOrWhiteSpace($gridC)) $gridC

    # ===================================================================
    # PART D - THE REQUEST MOVES, AND ONLY A RECALCULATION FOLLOWS
    # ===================================================================
    Write-P81Line ''
    Write-P81Line 'PART D - ITERATIONS CHANGE, RECALCULATION ONLY'
    Write-P81Line '----------------------------------------------'
    Write-P81Line 'THE VOLATILE ADAPTERS EXIST FOR THIS. No endpoint is invoked; if the sheet'
    Write-P81Line 'still says CURRENT here, the correction did not work.'
    $captureD0 = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $iterationsControl = [string]$simInspection.controls.monte_carlo_iterations.defined_name
    $requestBefore = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    Set-NamedValue -Workbook $wb -DefinedName $iterationsControl -Value ([double]$staleIterations)
    $requestAfter = Get-NamedValue -Workbook $wb -DefinedName $iterationsControl
    $null = Add-P81Check 'part D: the projected iteration control really moved' `
        ((([string]$requestAfter) -ceq ([string]$staleIterations)) -and
         (([string]$requestBefore) -cne ([string]$requestAfter))) `
        ($iterationsControl + ': ' + [string]$requestBefore + ' -> ' + [string]$requestAfter) `
        'PREREQUISITE'
    $observationD = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part D' -Distribution $historical `
        -Profile $historical -ExpectedPx $secondLabel -ExpectedYears $yearCount
    $null = $observationD
    # THE CALCULATION STATUS IS ASKED AFTER THE OBSERVATION IS FROZEN. It is an
    # out-of-cell call like any other, and part D's whole question is what the
    # worksheet said before this runner asked anything.
    $null = Add-P81Check 'part D: the deterministic calculation is still CURRENT' `
        ((([string]$excel.Run('PCCM_CalculationStatus')) -ceq $calcCurrent)) `
        ([string]$excel.Run('PCCM_CalculationStatus'))
    Invoke-P81AnnualChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -StartYear $startYear -Stage 'part D'
    Invoke-P81ReconciliationChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection `
        -P7 $p7 -Bank $bank -YearCount $yearCount -Label $secondLabel -LadderIndex $secondIndex `
        -ProfileState $historical -Stage 'part D'
    $captureD1 = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $movedD = Compare-P81Surface -Before $captureD0 -After $captureD1
    $null = Add-P81Check 'part D: the persisted annual payload survived the recalculation' `
        ($movedD.Moved -eq 0) $movedD.Detail

    # THE FOUR-WAY ANSWER PART D EXISTS FOR, already recorded above in
    # observation order: the derived rows before the recalculation, after it,
    # the frozen cells taken at that moment, and the rows again after the
    # accessors were finally called directly. The four outcomes stay apart -
    # the recalculation updated the rows; it left them alone and the cells were
    # right anyway; the cells errored; or only the later direct call moved them.
    Write-P81Line ''
    Write-P81Line ('    part D: the worksheet observation was frozen before any accessor was ' +
                   'invoked out of a cell, so the row movement above is the recalculation' +
                   [char]39 + 's alone.')
    Write-P81Line ''

    # ===================================================================
    # PART E - THE MODEL STOPS BEING VALID
    # ===================================================================
    Write-P81Line ''
    Write-P81Line 'PART E - AN INVALID MODEL'
    Write-P81Line '-------------------------'
    $victimRow = Get-P81RegisterRowIndex -Workbook $wb -Register $costRegister -PermanentId $victimId
    $null = Add-P81Check 'part E: the cost line the corpus names is in the register' `
        ($victimRow -ge 1) ($victimId + ' at body row ' + [string]$victimRow) 'PREREQUISITE'
    if ($victimRow -lt 1) { throw ('the P8-1 register edit could not find ' + $victimId) }
    Set-TableCell -Workbook $wb -SheetName ([string]$costRegister.sheet) `
        -TableName ([string]$costRegister.table_name) -RowIndex $victimRow `
        -ColumnIndex $maximumOrdinal -Value $invalidMaximum
    $calcAnnouncement = Invoke-P81Endpoint -Excel $excel -Endpoint 'PCCM_Calculate'
    $null = Add-P81Check 'part E: PCCM_Calculate refuses the edited model' `
        ($calcAnnouncement -like 'FAIL|*') $calcAnnouncement
    $observationE = Invoke-P81Observation -Excel $excel -Workbook $wb -Inspection $p8 `
        -SimInspection $simInspection -P7 $p7 -Stage 'part E' -Distribution $historical `
        -Profile $historical -ExpectedPx $secondLabel -ExpectedYears $yearCount
    $null = $observationE
    # AFTER THE FROZEN OBSERVATION, for the reason part D states.
    $null = Add-P81Check 'part E: the model is no longer CURRENT' `
        ((([string]$excel.Run('PCCM_CalculationStatus')) -cne $calcCurrent)) `
        ([string]$excel.Run('PCCM_CalculationStatus'))
    Invoke-P81AnnualChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -StartYear $startYear -Stage 'part E'
    Invoke-P81ReconciliationChecks -Workbook $wb -Inspection $p8 -SimInspection $simInspection `
        -P7 $p7 -Bank $bank -YearCount $yearCount -Label $secondLabel -LadderIndex $secondIndex `
        -ProfileState $historical -Stage 'part E'
    $captureE = Get-P81BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $movedE = Compare-P81Surface -Before $captureD1 -After $captureE
    $null = Add-P81Check 'part E: the persisted annual payload is still intact' `
        ($movedE.Moved -eq 0) $movedE.Detail
    $null = Add-P81InvariantChecks 'part E: the invalid model' $invariantsA `
        (Get-P81RunInvariants -Workbook $wb -Inspection $simInspection)

    Write-P81Line ''
    Write-P81Line (Format-Phase6State -State (Get-Phase6State -Workbook $wb `
        -Inspection $simInspection) -Label '  the run identity at the end of P8-1')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-P81Line ''
    Write-P81Line ('THE P8-1 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-P81Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-P81Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-P81Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-P81Line ''
        Write-P81Line 'EXCEL SHUTDOWN'
        Write-P81Line '--------------'
        if ($naturalExit) {
            Write-P81Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                           [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'P8-1'
            Write-P81Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-P81Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:P81Residual)) {
            Write-P81Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-P81Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-P81Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-P81Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-P81Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-P81Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-P81Check 'every COM release left 0 outstanding references' `
    (@($script:P81Residual).Count -eq 0) ((@($script:P81Residual)) -join '; ')
$null = Add-P81Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:P81Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:P81Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-P81Line ''
Write-P81Line 'VERDICT'
Write-P81Line '-------'
Write-P81Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
               [string]$failedPrereqs.Count + ' failed')
Write-P81Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
               [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-P81Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-P81Line ''
if ($ok) {
    Write-P81Line 'P8-1: PASS'
} elseif ($stoppedOnStateEvaluation) {
    Write-P81Line 'P8-1: STOPPED AT THE STATE-EVALUATION GATE'
    Write-P81Line ''
    Write-P81Line 'Return this report. The question of whether a non-writing semantic accessor'
    Write-P81Line 'is required is not settled by patching a harness around the answer.'
} else {
    Write-P81Line 'P8-1: FAIL'
    Write-P81Line ''
    Write-P81Line 'STOP AND REVIEW. Do not treat the Results surface as accepted until this is'
    Write-P81Line 'understood.'
}
Write-P81Line ''
Write-P81Line ('report                 : ' + $script:P81Path)
Write-Host ''
Write-Host ('The report is at ' + $script:P81Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
