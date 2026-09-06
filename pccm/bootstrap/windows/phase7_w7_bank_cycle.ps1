<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W7 runner: A to B and back to A, and the reused
    bank publishes a SHORTER answer than it held.

.DESCRIPTION
    W7 PROVES TWO PERSISTENCE PROPERTIES AT ONCE, because one cannot be shown
    without the other: the publication banks really do cycle A -> B -> A, and
    the bank that comes back round replaces a long annual answer with a short
    one without leaving the tail of the old one readable.

    THREE SUCCESSFUL RUNS, ONE SESSION, ALL FIXED-SEED:

      A1  twenty project years, published to bank A
      B   the same model with a DIFFERENT FIXED SEED, published to bank B
      A2  four project years, published back to bank A

    B CHANGES THE SEED, NOT THE CLOCK. A run is told apart by its identity, and
    a different FIXED seed is a different effective seed, a different request
    fingerprint and a different result digest - reproducibly, on any machine, in
    any order. Nothing here depends on a timestamp.

    THE SHRINK IS THE POINT, and the contract says so in as many words:
    "A four-year run published after a twenty-year one must not leave years 5-20
    readable as current, and a count is what says where the answer stops - not
    the last non-blank row." So this runner reads authority the contract's way -
    the publication marker plus the stamped year_count - and then checks the
    physical fact separately, because the stamp also declares that surplus rows
    are cleared. Rows 5 to 20 are checked ONE BY ONE across the whole span A1
    used to occupy, in the index columns, both ladders and both profile columns.
    Not row 5, and not a scan of the sheet: the contracted annual region, read
    at the coordinates the projection owns.

    THE INACTIVE BANK IS EVIDENCE, NOT SCENERY. Bank A is captured whole after
    A1 and must be value-identical after B commits - nothing is cleared merely
    because another bank became active. Bank B is captured whole after its own
    annual publication and must be value-identical after A2 commits.

    EACH ANNUAL RESULT IS REAL, NOT JUST WELL-SHAPED. All three are reconciled
    the way W5 and W7 reconcile: the total read from the block the projection
    names as the total percentile block, cross-checked against the contract's
    own Type-7 over that run's published iteration column, the profile summed to
    it, and the contingency checked against the contract's formula on the same
    measure's deterministic base - under the project's accepted identity rule,
    with no scaling and no new tolerance. There is no independent annual Windows
    oracle and none is claimed.

    WHAT W7 IS NOT. It moves no reporting selector - W7 owns that - provokes no
    stale or invalid refusal - W8 owns that - runs no sensitivity, and writes
    nothing to bank-selection state: every bank change is the consequence of a
    successful run through the ordinary endpoints.

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
# NOTHING IN THIS FILE CALLS ANY OF THEM. The accepted Phase-5 fixture
# choreography does - and its own file does not define them, because in a
# Gate-B run the Phase-4 driver dot-sources the scenarios file and the helpers
# are already in scope. Reaching the scenarios file directly leaves that
# dependency unmet, and W1 died on exactly that, one helper at a time.
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

# ===========================================================================
# THE REPORT AND THE CHECK LEDGER
# ===========================================================================
$script:W7Lines = New-Object System.Collections.ArrayList
$script:W7Path = ''
$script:W7Checks = New-Object System.Collections.ArrayList

function Write-W7Line {
    param([string]$Text = '')
    $null = $script:W7Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W7Path)) {
        try {
            Set-Content -LiteralPath $script:W7Path `
                -Value ($script:W7Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W7 had to perform to reach
# its subject - from a RESULT, which is W7's own claim about the calculation. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W7Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W7Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W7Line $line
    return $Ok
}

function Format-W7Value {
    param($Value)
    if ($null -eq $Value) { return '<null>' }
    if ($Value -is [System.DBNull]) { return '<empty>' }
    $text = ''
    try { $text = [string]$Value } catch { return '<unprintable>' }
    if ([string]::IsNullOrEmpty($text)) { return '<blank>' }
    return $text
}

# ===========================================================================
# EVERY RELEASE, WITH ITS COUNT
# ===========================================================================
# THE DISCIPLINE W1 CLOSED ON. ReleaseComObject returns the count REMAINING on
# that RCW: zero means the runtime let go, and anything above zero is a retained
# reference, named. `Invoke-NamedRelease` writes the ledger line but keeps that
# number to itself, and `Release-Transient` says nothing at all on success -
# which is how W1's first failure went unexplained. This wraps the SAME accepted
# primitive and records the number. com_lifecycle.ps1 is not modified.
$script:W7Residual = New-Object System.Collections.ArrayList

function Invoke-W7Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W7Residual.Add(
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
# THE CANDIDATE IDENTITY
# ===========================================================================
function Get-W7SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W7 result could not be attributed to a source revision')
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

# ===========================================================================
# THE COMPARISON
# ===========================================================================
# ONE RULE, APPLIED TO EVERY PROJECTED COLUMN, and the rule is chosen by what
# the ORACLE holds rather than by a list of column names typed into this file.
# That is what makes "all the other already-contracted fields" true by
# construction: a column the contract adds to the table is compared the moment
# the oracle carries it, and no edit here is needed or possible to forget.
#
#   the oracle has nothing    -> the workbook must publish a blank. A fabricated
#                                zero where the model has no risk is the single
#                                most valuable thing this can catch.
#   the oracle has a string   -> the published value must be that string
#   the oracle has a number   -> the published value must BE a number, within
#                                the accepted allowance. The typed reader hands
#                                back what Excel actually published, so a number
#                                written as text arrives as a String and fails
#                                here rather than being coerced into agreement.
function Compare-W7Cell {
    param($Got, $Expect, [double]$Allowance)
    if ($null -eq $Expect) {
        if (($null -eq $Got) -or ($Got -is [System.DBNull]) -or
            (($Got -is [string]) -and ([string]$Got).Length -eq 0)) { return '' }
        return ('published ' + (Format-W7Value $Got) + ' where the oracle has none')
    }
    if ($Expect -is [string]) {
        if (($Got -is [string]) -and (([string]$Got) -ceq ([string]$Expect))) { return '' }
        return ((Format-W7Value $Got) + ' vs ' + [string]$Expect)
    }
    if ($Got -isnot [double]) {
        return ((Format-W7Value $Got) + ' is not a number')
    }
    $difference = [Math]::Abs([double]$Got - [double]$Expect)
    if ($difference -le $Allowance) { return '' }
    return ([string]$Got + ' vs ' + [string]$Expect + ' (difference ' + [string]$difference + ')')
}

# THE PUBLISHED TABLE, ROW FOR ROW, EVERY PROJECTED COLUMN. A failing table
# reports the first twelve disagreements and the count: the count says how bad
# it is, the examples make it diagnosable, and 6,300 lines would do neither.
function Compare-W7Table {
    param($Live, $Expected, $Inspection, [string]$TableKey, [double]$Allowance,
          [string]$Label, $FieldMap)
    $want = @($Expected)
    $columns = @($Inspection.calc.tables.$TableKey.columns)
    $problems = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $want.Count; $index++) {
        foreach ($column in $columns) {
            $field = [string]$column
            if ($null -ne $FieldMap) {
                if ($FieldMap.ContainsKey($field)) { $field = [string]$FieldMap[$field] }
            }
            $at = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey $TableKey -ColumnKey $column
            if ($at -lt 0) {
                $null = $problems.Add([string]$column + ': not a projected column'); continue
            }
            $expect = $null
            if ($want[$index].PSObject.Properties.Name -contains $field) {
                $expect = $want[$index].$field
            } else {
                $null = $problems.Add([string]$column + ': the oracle carries no ' + $field)
                continue
            }
            $problem = Compare-W7Cell -Got $Live[$index][$at] -Expect $expect -Allowance $Allowance
            if (-not [string]::IsNullOrWhiteSpace($problem)) {
                $null = $problems.Add('row ' + [string]($index + 1) + ' ' + [string]$column +
                                      ': ' + $problem)
            }
        }
        if ($problems.Count -gt 12) { break }
    }
    $detail = ''
    if ($problems.Count -gt 0) {
        $detail = ($problems -join '; ')
        if ($problems.Count -gt 12) { $detail = $detail + ' ... (truncated)' }
    }
    return (Add-W7Check ($Label + ': ' + $TableKey +
                         ' matches the accepted Phase-5 oracle in every projected column') `
        ($problems.Count -eq 0) $detail)
}


# ===========================================================================
# THE PHASE-7 ANNUAL SURFACE, READ WITHOUT RUNNING IT
# ===========================================================================
# EVERY ADDRESS COMES FROM THE PROJECTION. Not one column letter or row number
# is written here: they are read from phase7_acceptance_inspection.json, which
# is projected from sim_contract.yaml by the build that projects them into
# modSimContract. A contract move therefore moves this runner too.
function Get-W7AnnualStamp {
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

# The first record row of the annual grid, in both index columns. A bank that
# published nothing has nothing here, and this is a different fact from an empty
# stamp: a run could in principle write one and not the other.
function Get-W7AnnualFirstRecord {
    param($Workbook, $Inspection, $P7, [string]$Bank)
    $records = $P7.annual_records
    $row = [string]([int]$records.first_record_row)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $records.index_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.index_columns.$Bank.$key
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address ($column + $row)))
    }
    foreach ($measure in $records.selected_px_profile_columns.$Bank.PSObject.Properties.Name) {
        $column = [string]$records.selected_px_profile_columns.$Bank.$measure
        $out.Add(('profile_' + $measure), (Get-SimRawCell -Workbook $Workbook `
            -Inspection $Inspection -Address ($column + $row)))
    }
    return $out
}

function Get-W7Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

# THE ANNUAL ATTEMPT, THROUGH THE PUBLISHED ENDPOINT AND ITS PUBLISHED ANSWER.
# The runner never judges the outcome itself: it asks production what happened.
function Invoke-W7Annual {
    param($Excel, $P7)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run([string]$P7.command_surface.annual_endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

# ===========================================================================
# THE INVARIANTS A REFUSED ANNUAL ATTEMPT MUST HOLD
# ===========================================================================
# The annual step is not a simulation. These are the published facts that must
# not move across a refused annual invocation: the whole run identity in both
# banks, the AUTO nonce and its durable pending marker, and the result digest.
#
# TWO ROWS ARE EXCLUDED, AND NOT QUIETLY. `simulation_status` and
# `status_evaluated_at` are DERIVED: PCCM_SimulationStatus recomputes and
# rewrites both every time it is asked, and the annual endpoint's own
# precondition asks it. Requiring the evaluation TIMESTAMP not to move would be
# requiring the annual step not to check whether it may run. So they are held
# out of the frozen set and checked separately on the terms that matter: the
# STATUS WORD itself must be unchanged, and only the timestamp beside it may
# move.
$script:W7DerivedRows = @('simulation_status', 'status_evaluated_at')

function Get-W7RunInvariants {
    param($Workbook, $Inspection)
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($script:W7DerivedRows -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-W7InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-W7Check ($Label + ': no run identity, nonce, pending marker or digest moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# A defined name that carries TEXT. Set-NamedValue writes a Double - PowerShell
# binds a COM property call site per argument type, so one polymorphic
# assignment silently fails the second time it is reached with another type -
# and the selected confidence level is a label, not a number.
function Set-W7NamedText {
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

# THE ITERATION COLUMN, READ AS ONE RANGE rather than as N round trips. The
# three COM objects it opens are released into the run's own ledger and counted
# where they are acquired, so the acquire/release balance stays true.
function Get-W7IterationBlock {
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
        if ($null -ne $range)  { Invoke-W7Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-W7Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W7Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}


# ===========================================================================
# COLUMN ARITHMETIC, AND THE ANNUAL LADDER IT REACHES
# ===========================================================================
# The projection gives the FIRST ladder column per bank per measure and the
# ladder's LENGTH; the eleven columns follow it. That is the contract's shape,
# so the offset is computed rather than eleven letters being typed in.
# THE DEFECT THAT STOPPED THE FIRST W7 RUN, and the reason it survived every
# static control: this loop used to read
#
#     foreach ($character in [string]$Letters.ToUpperInvariant().ToCharArray())
#
# and a cast binds tighter than the enumeration, so `[string]` was applied to
# the whole char[] rather than to each character. PowerShell renders an array as
# a string by JOINING its elements with $OFS - a single space by default - so
# ['A','D'] became the ONE-element sequence "A D", and the body then asked for
# [char]"A D": "String must be exactly one character long."
#
# It worked for every single-letter column, because ['A'] renders as "A" and
# [char]"A" is fine. The first multi-letter column it ever met was the projected
# quantile first column, and that is exactly where the run died. There is no
# cast here now: .ToCharArray() already yields the chars, one at a time, which
# is the form the accepted harness has always used.
function ConvertTo-W7ColumnNumber {
    param([string]$Letters)
    $number = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - 64)
    }
    return $number
}

# The inverse, in the same bijective base 26: there is no zero digit, so the
# remainder is taken on (n - 1) and the quotient steps down by the digit that
# was just emitted. Z -> AA is the boundary that proves it.
function ConvertFrom-W7ColumnNumber {
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

# ONE ANNUAL RECORD, WHOLE: its index pair, both ladders, both profile values.
# Read as three ranges rather than as fifty round trips, and every address comes
# from the projection.
function Get-W7AnnualRecord {
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
        $first = ConvertTo-W7ColumnNumber -Letters ([string]$records.quantile_first_column.$Bank.$measure)
        $ladder = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $count; $index++) {
            $column = ConvertFrom-W7ColumnNumber -Number ($first + $index)
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

# THE WHOLE CONTRACTED ANNUAL REGION, IN ONE READ, so compactness can be proved
# rather than sampled: every row the region can hold, in the index column.
function Get-W7AnnualRegionIndex {
    param($Workbook, $Inspection, $P7, [string]$Bank, $Ledger)
    $records = $P7.annual_records
    $column = [string]$records.index_columns.$Bank.project_index
    $firstRow = [int]$records.first_record_row
    $lastRow = $firstRow + [int]$records.max_record_rows - 1
    $sheets = $null; $sheet = $null; $range = $null
    $acquired = 0
    try {
        $sheets = $Workbook.Worksheets
        $acquired = $acquired + 1
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $acquired = $acquired + 1
        $range = $sheet.Range($column + [string]$firstRow + ':' + $column + [string]$lastRow)
        $acquired = $acquired + 1
        return [pscustomobject]@{ Values = $range.Value2; Acquired = $acquired }
    } finally {
        if ($null -ne $range)  { Invoke-W7Release $Ledger $range  'Range(annual region)'; $range  = $null }
        if ($null -ne $sheet)  { Invoke-W7Release $Ledger $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W7Release $Ledger $sheets 'Worksheets';           $sheets = $null }
    }
}

# ===========================================================================
# THE CONTRACT'S OWN TYPE-7 PERCENTILE
# ===========================================================================
# sim_contract.yaml statistics.percentile, transcribed and nothing else:
#
#     h  = (n - 1) * p
#     lo = floor(h)
#     hi = min(lo + 1, n - 1)
#     f  = h - lo
#     value = (1 - f) * x[lo] + f * x[hi]
#
# Convex interpolation, on a COPY that is sorted - the contract's
# `sorting: on_copies_only`. This derives the TOTAL percentile from the
# published iteration column. It is deliberately NOT the selected-Px annual
# BLEND: that is production's, and reimplementing it here would compare an
# implementation against a copy of itself.
function Get-W7Type7Value {
    param([double[]]$Values, [double]$Probability)
    $sorted = @($Values | Sort-Object)
    $n = $sorted.Count
    if ($n -lt 1) { throw 'a type-7 percentile needs at least one value' }
    if ($n -eq 1) { return [double]$sorted[0] }
    $h = ($n - 1) * $Probability
    $lo = [int][Math]::Floor($h)
    $hi = [Math]::Min($lo + 1, $n - 1)
    $f = $h - $lo
    return ((1 - $f) * [double]$sorted[$lo]) + ($f * [double]$sorted[$hi])
}

# THE IDENTITY RULE, READ FROM THE CORPUS. |delta| <= max(floor, coefficient *
# max(scale floor, conditioning scale)). The conditioning scale names the
# MAGNITUDE OF THE ARITHMETIC PERFORMED - here the annual terms summed and the
# aggregate they are compared against - never the magnitude of the net result,
# which is the ERRATUM C1 correction this project already accepted.
function Get-W7IdentityAllowance {
    param($Provenance, [double]$ConditioningScale)
    $floor = [double]$Provenance.identity_absolute_floor
    $coefficient = [double]$Provenance.identity_relative_coefficient
    $scaleFloor = [double]$Provenance.conditioning_scale_floor
    $scale = [Math]::Max($scaleFloor, [Math]::Abs($ConditioningScale))
    return [Math]::Max($floor, $coefficient * $scale)
}


# ===========================================================================
# THE WHOLE ANNUAL ANSWER, AS PLAIN DATA
# ===========================================================================
# Captured before a move and compared after it. What comes back holds strings,
# doubles and nulls only - never a COM object - so a captured surface can be
# held across an endpoint call without keeping anything alive.
function Get-W7AnnualSurface {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $surface = New-Object System.Collections.Specialized.OrderedDictionary
    $stamp = Get-W7AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 -Bank $Bank
    foreach ($key in $stamp.Keys) { $surface.Add(('stamp.' + [string]$key), $stamp[$key]) }
    $ladderCount = [int]$P7.annual_records.quantile_count
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W7AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
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

# Two captured surfaces, compared key by key on the accepted "same value" rule -
# type included, so a number that became text is a change. `Only` narrows the
# comparison to one family of keys; `Except` excludes one.
function Compare-W7Surface {
    param($Before, $After, [string]$Only = '', [string]$Except = '')
    $moved = New-Object System.Collections.ArrayList
    foreach ($key in $Before.Keys) {
        $name = [string]$key
        if ((-not [string]::IsNullOrEmpty($Only)) -and (-not $name.Contains($Only))) { continue }
        if ((-not [string]::IsNullOrEmpty($Except)) -and $name.Contains($Except)) { continue }
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

# The published iteration block, before and after, compared value for value.
function Compare-W7IterationGrid {
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

# ===========================================================================
# ONE RECONCILIATION, AT WHATEVER RUNG IS SELECTED
# ===========================================================================
# W5's three checks, parameterised by the rung so the SAME code proves the P80
# baseline and the P50 rerun. A second copy for the second selector would be a
# second behaviour, and the point of W7 is that only the rung changes.
function Invoke-W7Reconciliation {
    param($Workbook, $Inspection, $P7, $Provenance, $Grid, [string]$Bank,
          [string]$Label, [double]$Probability, [int]$LadderIndex,
          $ProfileSums, $ProfileScale, [string]$Stage)
    $semantics = $P7.summary_semantics
    $rowKey = 'quantile_' + [string]($LadderIndex + 1)
    $contingencyBlock = $Inspection.sim_data.contingency_ladder
    foreach ($measure in @($semantics.contingency_measures | ForEach-Object { [string]$_ })) {
        $column = 2
        if ($measure -eq 'pv') { $column = 3 }
        $totals = New-Object System.Collections.ArrayList
        for ($row = 1; $row -le [int]$Grid.GetLength(0); $row++) {
            $value = $Grid[$row, $column]
            if ($value -is [double]) { $null = $totals.Add([double]$value) }
        }
        $derived = Get-W7Type7Value -Values ([double[]]@($totals)) -Probability $Probability

        # (A) THE TOTAL, from the block the projection names as the total
        # percentile block, cross-checked before anything is reconciled to it.
        $total = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey $rowKey
        $totalScale = [Math]::Abs($derived)
        if ($total -is [double]) { $totalScale = $totalScale + [Math]::Abs([double]$total) }
        $totalAllowance = Get-W7IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $totalScale
        $totalDelta = [double]::PositiveInfinity
        if ($total -is [double]) { $totalDelta = [Math]::Abs([double]$total - $derived) }
        $null = Add-W7Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' TOTAL equals the contract' + [char]39 +
                             's Type-7 value over the iteration column') `
            (($total -is [double]) -and ($totalDelta -le $totalAllowance)) `
            ('published ' + (Format-SimValue $total) + ', derived ' + [string]$derived +
             ', delta ' + [string]$totalDelta + ', allowance ' + [string]$totalAllowance)

        # (B) sum_y Profile_Px(y) = reported Px TOTAL.
        $sum = [double]$ProfileSums[$measure]
        $scale = [double]$ProfileScale[$measure]
        if ($total -is [double]) { $scale = $scale + [Math]::Abs([double]$total) }
        $identityAllowance = Get-W7IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $scale
        $delta = [double]::PositiveInfinity
        if ($total -is [double]) { $delta = [Math]::Abs($sum - [double]$total) }
        $null = Add-W7Check ($Stage + ': the ' + $measure +
                             ' selected-Px profile sums to the reported ' + $Label + ' TOTAL') `
            (($total -is [double]) -and ($delta -le $identityAllowance)) `
            ('sum ' + [string]$sum + ', total ' + (Format-SimValue $total) +
             ', delta ' + [string]$delta + ', allowance ' + [string]$identityAllowance +
             ' (conditioning scale ' + [string]$scale + ')')

        # (C) THE CONTINGENCY, on the contract's own formula and the SAME
        # measure's deterministic base.
        $baseline = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey ([string]$semantics.baseline_metric_key)
        $contingency = Get-SimRawCell -Workbook $Workbook -Inspection $Inspection `
            -Address ([string]$contingencyBlock.bank_value_columns.$Bank.$measure +
                      [string]([int]$contingencyBlock.rows.$rowKey))
        $expectedContingency = [double]::NaN
        $contingencyDelta = [double]::PositiveInfinity
        $contingencyScale = 0.0
        if (($total -is [double]) -and ($baseline -is [double])) {
            $expectedContingency = [double]$total - [double]$baseline
            $contingencyScale = [Math]::Abs([double]$total) + [Math]::Abs([double]$baseline)
            if ($contingency -is [double]) {
                $contingencyDelta = [Math]::Abs([double]$contingency - $expectedContingency)
            }
        }
        $contingencyAllowance = Get-W7IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $contingencyScale
        $null = Add-W7Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' contingency is ' + [string]$semantics.contingency_formula) `
            (($contingency -is [double]) -and ($baseline -is [double]) -and
             ($total -is [double]) -and ($contingencyDelta -le $contingencyAllowance)) `
            ('published ' + (Format-SimValue $contingency) + ', total ' +
             (Format-SimValue $total) + ' - base ' + (Format-SimValue $baseline) + ' = ' +
             [string]$expectedContingency + ', delta ' + [string]$contingencyDelta +
             ', allowance ' + [string]$contingencyAllowance)

        Write-W7Line ('    ' + $Stage.PadRight(10) + $measure.PadRight(8) +
                      ' profile sum ' + [string]$sum + '  total ' + (Format-SimValue $total) +
                      '  contingency ' + (Format-SimValue $contingency))
    }
}

# The per-year profile sums and their conditioning scale, read back from the
# persisted records. The blend itself is never recomputed here.
function Get-W7ProfileSums {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount, $Measures)
    $sums = @{}
    $scale = @{}
    $problems = New-Object System.Collections.ArrayList
    foreach ($measure in @($Measures)) { $sums[$measure] = 0.0; $scale[$measure] = 0.0 }
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W7AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        foreach ($measure in @($Measures)) {
            $value = $record[('profile_' + $measure)]
            if ($value -isnot [double]) {
                $null = $problems.Add('year ' + [string]($offset + 1) + ' ' + $measure +
                                      ' profile published ' + (Format-SimValue $value))
                continue
            }
            $sums[$measure] = $sums[$measure] + [double]$value
            $scale[$measure] = $scale[$measure] + [Math]::Abs([double]$value)
        }
    }
    return [pscustomobject]@{ Sums = $sums; Scale = $scale; Problems = @($problems) }
}


# ===========================================================================
# ONE BANK, CAPTURED WHOLE
# ===========================================================================
# The simulation block AND the annual answer for one bank, as plain data, so a
# bank can be compared against itself across another bank's publication. Only
# strings, doubles and nulls cross this boundary.
function Get-W7BankCapture {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $capture = New-Object System.Collections.Specialized.OrderedDictionary
    $block = Get-SimBankBlock -Workbook $Workbook -Inspection $Inspection -Bank $Bank
    foreach ($key in $block.Keys) { $capture.Add(('sim.' + [string]$key), $block[$key]) }
    $annual = Get-W7AnnualSurface -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank $Bank -YearCount $YearCount
    foreach ($key in $annual.Keys) { $capture.Add(('annual.' + [string]$key), $annual[$key]) }
    return $capture
}

# ===========================================================================
# THE FORMER AUTHORITATIVE SPAN, ROW BY ROW
# ===========================================================================
# After a shorter answer replaces a longer one, every row the LONG answer used
# to occupy beyond the new year count is read in full - both index columns,
# every rung of both ladders, and both profile columns - and required to be
# blank. The stamp already says where the answer stops; this says the tail is
# not merely out of scope but gone, which is what `surplus_rows_cleared`
# promises.
function Get-W7SurplusResidue {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$FromRow, [int]$ToRow)
    $residue = New-Object System.Collections.ArrayList
    for ($offset = $FromRow - 1; $offset -le $ToRow - 1; $offset++) {
        $record = Get-W7AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
            -Bank $Bank -Offset $offset
        foreach ($key in $record.Keys) {
            $name = [string]$key
            $value = $record[$name]
            if ($name.StartsWith('ladder_')) {
                $ladder = @($value)
                for ($index = 0; $index -lt $ladder.Count; $index++) {
                    if (-not (Test-SimBlank -Value $ladder[$index])) {
                        $null = $residue.Add('row ' + [string]($offset + 1) + ' ' + $name + ' rung ' +
                                             [string]($index + 1) + ' = ' +
                                             (Format-SimValue $ladder[$index]))
                    }
                }
                continue
            }
            if (-not (Test-SimBlank -Value $value)) {
                $null = $residue.Add('row ' + [string]($offset + 1) + ' ' + $name + ' = ' +
                                     (Format-SimValue $value))
            }
        }
        if ($residue.Count -gt 12) { break }
    }
    return @($residue)
}

# THE BANK A SUCCESSFUL RUN IS ABOUT TO PUBLISH TO, from the projected cycle.
# The runner never assumes a blank selector means A, or that A is followed by B.
function Get-W7CandidateBank {
    param($P7, [string]$ActiveBank)
    foreach ($entry in @($P7.publication_semantics.candidate_target)) {
        $active = ''
        if ($null -ne $entry.active_bank) { $active = [string]$entry.active_bank }
        if ($active -ceq $ActiveBank) { return [string]$entry.candidate_bank }
    }
    return ''
}

# ONE COMPLETE RUN OF THE ACCEPTED WORKFLOW: apply the request, calculate if the
# model changed, simulate, publish the annual answer. Every step through its own
# production endpoint; nothing is written to machine state.
function Invoke-W7Run {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $P7,
          $Model, [double]$Seed, [int]$Iterations, [string]$Label, [switch]$Recalculate)
    $applied = ''
    if ($null -ne $Model) {
        $applied = [string](Set-Phase5Fixture -Excel $Excel -Workbook $Workbook `
            -Manifest $Manifest -Inspection $Inspection -Model $Model)
        $null = Add-W7Check ($Label + ': the model was applied through the accepted fixture') `
            ($applied -like 'OK|*') $applied 'PREREQUISITE'
        if (-not ($applied -like 'OK|*')) { return $false }
    }
    Set-NamedValue -Workbook $Workbook `
        -DefinedName ([string]$SimInspection.controls.monte_carlo_iterations.defined_name) `
        -Value ([double]$Iterations)
    Set-NamedValue -Workbook $Workbook `
        -DefinedName ([string]$SimInspection.controls.random_seed.defined_name) -Value $Seed
    if ($Recalculate) {
        $calcFailure = ''
        try {
            $null = [string](Invoke-Phase5ProductionOperation -Excel $Excel `
                -Operation 'PCCM_Calculate' -Stage ($Label + ' calculate'))
        } catch { $calcFailure = (Format-Err $_) }
        $null = Add-W7Check ($Label + ': PCCM_Calculate succeeded') `
            ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
        if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { return $false }
        $null = Add-W7Check ($Label + ': the calculation reports CURRENT') `
            ((([string]$Excel.Run('PCCM_CalculationStatus')) -ceq 'CURRENT')) '' 'PREREQUISITE'
    }
    $simResult = Invoke-Phase6Simulation -Excel $Excel
    $null = Add-W7Check ($Label + ': PCCM_RunSimulation succeeded') `
        (Test-Phase6Announced -Result $simResult -Kind 'OK') $simResult
    if (-not (Test-Phase6Announced -Result $simResult -Kind 'OK')) { return $false }
    $null = Add-W7Check ($Label + ': the simulation reports CURRENT') `
        ((([string]$Excel.Run('PCCM_SimulationStatus')) -ceq 'CURRENT')) ''
    $annual = Invoke-W7Annual -Excel $Excel -P7 $P7
    $null = Add-W7Check ($Label + ': PCCM_RunAnnualStochastic succeeded') `
        ($annual -like 'OK|*') $annual
    return ($annual -like 'OK|*')
}

# THE AUTHORITY RULE, APPLIED AS THE CONTRACT STATES IT: the marker plus the
# stamped year count, never the last non-blank row.
function Test-W7Authoritative {
    param($Stamp, $P7, [int]$ExpectedYears)
    $marker = [string]$P7.annual_records.stamp.published_marker
    $markerOk = (Test-SimExactText -Actual $Stamp['published'] -Expected $marker)
    $countOk = (Test-SimExactDouble -Actual $Stamp['year_count'] -Expected ([double]$ExpectedYears))
    return [pscustomobject]@{
        Ok = ($markerOk -and $countOk)
        Detail = ('marker ' + (Format-SimValue $Stamp['published']) + ', year_count ' +
                  (Format-SimValue $Stamp['year_count']) + ', expected ' + [string]$ExpectedYears)
    }
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W7 (bank cycling and the duration shrink)' -ForegroundColor Cyan
Write-Host '=======================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath  = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
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

$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W7' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no W7 scenario.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-W7SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W7 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w7-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W7Path = Join-Path $tempRoot 'phase7_w7_bank_cycle.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$longModel = $case.model
$shortModel = $case.shrink_model
$longYears = [int]$longModel.timeline.duration
$shortYears = [int]$shortModel.timeline.duration
$driverCount = @($longModel.cost_lines).Count + @($longModel.risks).Count
$iterations = [int]$case.iterations
$firstSeed = [double]$case.supplied_seed
$secondSeed = [double]$case.second_supplied_seed
$selectedLabel = [string]$case.selected_confidence_level
$quantileLabels = @($gateBCases.vocabulary.quantile_labels | ForEach-Object { [string]$_ })
$selectedIndex = [array]::IndexOf($quantileLabels, $selectedLabel)
$selectedProbability = 0.0
if ($selectedLabel -match '^P(\d+)$') { $selectedProbability = [double]$Matches[1] / 100.0 }
$measures = @($p7.summary_semantics.contingency_measures | ForEach-Object { [string]$_ })
$publication = $p7.publication_semantics
$currentDistributionState = [string]$p7.handoff.distribution_states[1]
$currentProfileState = [string]$p7.handoff.profile_states[1]
$allowance = [double]$cases.provenance.comparison_absolute_floor

Write-W7Line 'PCCM - PHASE 7 W7: BANK CYCLING AND THE DURATION SHRINK'
Write-W7Line '======================================================'
Write-W7Line ''
Write-W7Line 'This is the MINIMAL W7 runner. Three successful FIXED-seed runs in one'
Write-W7Line 'session: A1 long, B on a different seed, A2 short and back in the first'
Write-W7Line 'bank. No selector move (W6), no stale or invalid refusal (W8), no'
Write-W7Line 'sensitivity, and nothing written to bank-selection state.'
Write-W7Line ''
Write-W7Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W7Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W7Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W7Line ('git HEAD               : ' + [string]$revision.Head)
Write-W7Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W7Line ('model version          : ' + [string]$manifest.model_version)
Write-W7Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W7Line ('build directory        : ' + $BuildDir)
Write-W7Line ('working copy           : ' + $tempRoot)
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
    Write-W7Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W7Line ''
Write-W7Line 'THE THREE RUNS'
Write-W7Line '--------------'
Write-W7Line ('  A1  ' + [string]$driverCount + ' drivers, ' + [string]$longYears +
              ' project years, FIXED seed ' + [string]$firstSeed)
Write-W7Line ('  B   the same model, FIXED seed ' + [string]$secondSeed +
              ' - a different request, deterministically, never a timestamp')
Write-W7Line ('  A2  ' + [string]$shortYears + ' project years, FIXED seed ' + [string]$firstSeed)
Write-W7Line ('  selected Px          : ' + $selectedLabel)
Write-W7Line ('  iterations           : ' + [string]$iterations)
Write-W7Line ('  identity rule        : ' + [string]$cases.provenance.identity_rule)
Write-W7Line ''
Write-W7Line 'THE PUBLICATION RULE THIS SCENARIO TESTS, projected rather than encoded:'
Write-W7Line ('  ' + [string]$publication.authority_rule)
Write-W7Line ('  marker written last  : ' + [string]$publication.published_written_last)
Write-W7Line ('  cleared before write : ' + [string]$publication.cleared_before_write)
Write-W7Line ('  surplus rows cleared : ' + [string]$publication.surplus_rows_cleared)
Write-W7Line ''
Write-W7Line 'THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE, and none is claimed. Each of'
Write-W7Line 'the three annual results is reconciled the way W5 and W6 reconcile, so a'
Write-W7Line 'bank capture is a real successful result and not only a storage shape.'
Write-W7Line ''

$null = Add-W7Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W7Line ''
    Write-W7Line 'STOP. Excel was never started for the W7 session, and nothing was accepted.'
    Write-W7Line ('report                 : ' + $script:W7Path)
    Write-W7Line 'W7: FAIL'
    exit 1
}

$null = Add-W7Check 'the W7 fixture shrinks and the two seeds differ' `
    (($longYears -gt $shortYears) -and ($longYears -eq 20) -and ($shortYears -eq 4) -and
     ($firstSeed -ne $secondSeed) -and (([string]$case.seed_mode) -ceq 'FIXED') -and
     ($iterations -eq 1000)) `
    ([string]$longYears + ' -> ' + [string]$shortYears + ' years, seeds ' +
     [string]$firstSeed + ' and ' + [string]$secondSeed) 'PREREQUISITE'
$null = Add-W7Check 'both models carry the same drivers, so only the duration shrinks' `
    ((@($shortModel.cost_lines).Count -eq @($longModel.cost_lines).Count) -and
     (@($shortModel.risks).Count -eq @($longModel.risks).Count)) `
    ([string]$driverCount + ' drivers in both') 'PREREQUISITE'
# THE BANK CYCLE IS THE PROJECTION'S, not an assumption that A is followed by B.
$null = Add-W7Check 'the projected bank cycle is two banks and three transitions' `
    ((@($publication.candidate_target).Count -eq 3) -and
     (@($publication.bank_labels).Count -eq 2) -and
     ((Get-W7CandidateBank -P7 $p7 -ActiveBank '') -ceq 'A') -and
     ((Get-W7CandidateBank -P7 $p7 -ActiveBank 'A') -ceq 'B') -and
     ((Get-W7CandidateBank -P7 $p7 -ActiveBank 'B') -ceq 'A')) `
    ('none->' + (Get-W7CandidateBank -P7 $p7 -ActiveBank '') + ', A->' +
     (Get-W7CandidateBank -P7 $p7 -ActiveBank 'A') + ', B->' +
     (Get-W7CandidateBank -P7 $p7 -ActiveBank 'B')) 'PREREQUISITE'
$null = Add-W7Check 'the contract clears surplus annual rows, so the shrink is checkable' `
    ([bool]$publication.surplus_rows_cleared) `
    ([string]$publication.authority_rule) 'PREREQUISITE'

# ===========================================================================
# THE W7 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 7 W7 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W7Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W7Line 'ever terminated, and the workbook is never saved.'
    Write-W7Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-W7Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection
    Set-W7NamedText -Workbook $wb `
        -DefinedName ([string]$inspection.inputs.($p7.selector_semantics.selector_input_key).defined_name) `
        -Value $selectedLabel

    $stateStart = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $expectedFirstBank = Get-W7CandidateBank -P7 $p7 -ActiveBank (Get-Phase6ActiveBank -State $stateStart)
    $nonceStart = $stateStart['shared']['next_auto_nonce']
    $pendingStart = $stateStart['pending_auto_nonce']

    # ===================================================================
    # A1 - THE LONG RESULT
    # ===================================================================
    Write-W7Line ''
    Write-W7Line ('A1 - ' + [string]$longYears + ' PROJECT YEARS')
    Write-W7Line '--------------------------'
    if (-not (Invoke-W7Run -Excel $excel -Workbook $wb -Manifest $manifest -Inspection $inspection `
            -SimInspection $simInspection -P7 $p7 -Model $longModel -Seed $firstSeed `
            -Iterations $iterations -Label 'A1' -Recalculate)) {
        throw 'A1 did not complete'
    }
    $stateA1 = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bankA = Get-Phase6ActiveBank -State $stateA1
    $null = Add-W7Check ('A1 published to the bank the projected cycle names (' +
                         $expectedFirstBank + ')') `
        ($bankA -ceq $expectedFirstBank) ($bankA + ' vs ' + $expectedFirstBank)
    $runIdA1 = $stateA1[('bank_' + $bankA)]['run_id']
    $captureA1 = Get-W7BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA -YearCount $longYears
    $authorityA1 = Test-W7Authoritative -Stamp (Get-W7AnnualStamp -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Bank $bankA) -P7 $p7 -ExpectedYears $longYears
    $null = Add-W7Check ('A1 is authoritative for ' + [string]$longYears + ' years') `
        ([bool]$authorityA1.Ok) ([string]$authorityA1.Detail)
    $handoffA1 = Get-W7Handoff -Excel $excel -P7 $p7
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $null = Add-W7Check 'A1 reports both annual products CURRENT and the long year count' `
        (((([string]$handoffA1[$accessors[0]]) -ceq $currentDistributionState)) -and
         ((([string]$handoffA1[$accessors[1]]) -ceq $currentProfileState)) -and
         (([int]$handoffA1[$accessors[3]]) -eq $longYears)) `
        ((Format-SimValue $handoffA1[$accessors[0]]) + ' / ' +
         (Format-SimValue $handoffA1[$accessors[1]]) + ' / ' +
         (Format-SimValue $handoffA1[$accessors[3]]))
    # THE YEAR AXIS THE LONG ANSWER PUBLISHED, checked against the fixture.
    $axisProblems = New-Object System.Collections.ArrayList
    for ($offset = 0; $offset -lt $longYears; $offset++) {
        $expectedIndex = [double]($offset + 1)
        $expectedYear = [double]([int]$longModel.timeline.start_year + $offset)
        if (-not (Test-SimExactDouble -Actual $captureA1[('annual.year' + [string]($offset + 1) +
                                                          '.project_index')] -Expected $expectedIndex)) {
            $null = $axisProblems.Add('row ' + [string]($offset + 1) + ' project_index')
        }
        if (-not (Test-SimExactDouble -Actual $captureA1[('annual.year' + [string]($offset + 1) +
                                                          '.calendar_year')] -Expected $expectedYear)) {
            $null = $axisProblems.Add('row ' + [string]($offset + 1) + ' calendar_year')
        }
    }
    $null = Add-W7Check ('A1 published ' + [string]$longYears +
                         ' correctly indexed project years') `
        ($axisProblems.Count -eq 0) (($axisProblems -join '; '))
    $iterationsA1 = Get-W7IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bankA -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsA1.Acquired
    $profileA1 = Get-W7ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA -YearCount $longYears -Measures $measures
    $null = Add-W7Check 'every A1 profile value is a number' `
        (@($profileA1.Problems).Count -eq 0) ((@($profileA1.Problems)) -join '; ')
    Invoke-W7Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsA1.Values -Bank $bankA `
        -Label $selectedLabel -Probability $selectedProbability -LadderIndex $selectedIndex `
        -ProfileSums $profileA1.Sums -ProfileScale $profileA1.Scale -Stage 'A1'
    # THE ROW AFTER THE LONG ANSWER IS ALREADY BLANK, so the shrink check below
    # is measuring a change rather than a state that was always true.
    $tailA1 = Get-W7SurplusResidue -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA -FromRow ($longYears + 1) -ToRow ($longYears + 1)
    $null = Add-W7Check ('the row after A1' + [char]39 + 's answer is blank') `
        (@($tailA1).Count -eq 0) ((@($tailA1)) -join '; ')

    # ===================================================================
    # B - THE NEXT RUN, IN THE OTHER BANK
    # ===================================================================
    Write-W7Line ''
    Write-W7Line 'B - THE NEXT RUN, ON A DIFFERENT FIXED SEED'
    Write-W7Line '-------------------------------------------'
    $expectedSecondBank = Get-W7CandidateBank -P7 $p7 -ActiveBank $bankA
    # THE MODEL IS UNCHANGED, so no recalculation is required: only the request
    # moves, which is what makes B a different run rather than a different model.
    if (-not (Invoke-W7Run -Excel $excel -Workbook $wb -Manifest $manifest -Inspection $inspection `
            -SimInspection $simInspection -P7 $p7 -Model $null -Seed $secondSeed `
            -Iterations $iterations -Label 'B')) {
        throw 'B did not complete'
    }
    $stateB = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bankB = Get-Phase6ActiveBank -State $stateB
    $null = Add-W7Check ('B published to the bank the projected cycle names (' +
                         $expectedSecondBank + ')') `
        (($bankB -ceq $expectedSecondBank) -and ($bankB -cne $bankA)) `
        ($bankB + ' vs ' + $expectedSecondBank)
    $runIdB = $stateB[('bank_' + $bankB)]['run_id']
    $null = Add-W7Check 'the B run id advanced from A1' `
        ((Test-SimExactDouble -Actual $runIdB -Expected ([double]$runIdA1 + 1))) `
        ((Format-SimValue $runIdA1) + ' -> ' + (Format-SimValue $runIdB))
    $null = Add-W7Check 'the B run is a different request from A1' `
        ((-not (Test-SimSameValue -A $stateA1[('bank_' + $bankA)]['effective_seed'] `
                    -B $stateB[('bank_' + $bankB)]['effective_seed'])) -and
         (-not (Test-SimSameValue -A $stateA1[('bank_' + $bankA)]['request_fingerprint'] `
                    -B $stateB[('bank_' + $bankB)]['request_fingerprint']))) `
        ('seed ' + (Format-SimValue $stateB[('bank_' + $bankB)]['effective_seed']) +
         ', fingerprint ' + (Format-SimValue $stateB[('bank_' + $bankB)]['request_fingerprint']))
    # BANK A IS NOT CLEARED MERELY BECAUSE B BECAME ACTIVE.
    $captureA1AfterB = Get-W7BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA -YearCount $longYears
    $aAfterB = Compare-W7Surface -Before $captureA1 -After $captureA1AfterB
    $null = Add-W7Check ('bank ' + $bankA + ' is value-identical after B committed') `
        ($aAfterB.Moved -eq 0) $aAfterB.Detail
    $captureB = Get-W7BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankB -YearCount $longYears
    $authorityB = Test-W7Authoritative -Stamp (Get-W7AnnualStamp -Workbook $wb `
        -Inspection $simInspection -P7 $p7 -Bank $bankB) -P7 $p7 -ExpectedYears $longYears
    $null = Add-W7Check ('B is authoritative for ' + [string]$longYears + ' years') `
        ([bool]$authorityB.Ok) ([string]$authorityB.Detail)
    $iterationsB = Get-W7IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bankB -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsB.Acquired
    $profileB = Get-W7ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankB -YearCount $longYears -Measures $measures
    $null = Add-W7Check 'every B profile value is a number' `
        (@($profileB.Problems).Count -eq 0) ((@($profileB.Problems)) -join '; ')
    Invoke-W7Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsB.Values -Bank $bankB `
        -Label $selectedLabel -Probability $selectedProbability -LadderIndex $selectedIndex `
        -ProfileSums $profileB.Sums -ProfileScale $profileB.Scale -Stage 'B'

    # ===================================================================
    # A2 - BACK TO THE FIRST BANK, SHORTER
    # ===================================================================
    Write-W7Line ''
    Write-W7Line ('A2 - BACK TO BANK ' + $bankA + ' WITH ' + [string]$shortYears + ' YEARS')
    Write-W7Line '-------------------------------------'
    $expectedThirdBank = Get-W7CandidateBank -P7 $p7 -ActiveBank $bankB
    if (-not (Invoke-W7Run -Excel $excel -Workbook $wb -Manifest $manifest -Inspection $inspection `
            -SimInspection $simInspection -P7 $p7 -Model $shortModel -Seed $firstSeed `
            -Iterations $iterations -Label 'A2' -Recalculate)) {
        throw 'A2 did not complete'
    }
    $stateA2 = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bankA2 = Get-Phase6ActiveBank -State $stateA2
    $null = Add-W7Check ('A2 cycled back to bank ' + $expectedThirdBank) `
        (($bankA2 -ceq $expectedThirdBank) -and ($bankA2 -ceq $bankA)) `
        ($bankA2 + ' vs ' + $expectedThirdBank)
    $runIdA2 = $stateA2[('bank_' + $bankA2)]['run_id']
    $null = Add-W7Check 'the A2 run id advanced from B' `
        ((Test-SimExactDouble -Actual $runIdA2 -Expected ([double]$runIdB + 1))) `
        ((Format-SimValue $runIdB) + ' -> ' + (Format-SimValue $runIdA2))
    $null = Add-W7Check 'the A2 identity is distinct from both A1 and B' `
        ((-not (Test-SimSameValue -A $runIdA2 -B $runIdA1)) -and
         (-not (Test-SimSameValue -A $runIdA2 -B $runIdB)) -and
         (-not (Test-SimSameValue -A $captureA1['sim.request_fingerprint'] `
                    -B $stateA2[('bank_' + $bankA2)]['request_fingerprint'])) -and
         (-not (Test-SimSameValue -A $captureB['sim.request_fingerprint'] `
                    -B $stateA2[('bank_' + $bankA2)]['request_fingerprint']))) `
        ('run ' + (Format-SimValue $runIdA2) + ', fingerprint ' +
         (Format-SimValue $stateA2[('bank_' + $bankA2)]['request_fingerprint']))
    $null = Add-W7Check 'the A2 applied timeline is the shorter one' `
        ((-not (Test-SimSameValue -A $captureA1['sim.applied_timeline'] `
                    -B $stateA2[('bank_' + $bankA2)]['applied_timeline']))) `
        ((Format-SimValue $captureA1['sim.applied_timeline']) + ' -> ' +
         (Format-SimValue $stateA2[('bank_' + $bankA2)]['applied_timeline']))

    # ---- THE SHRINK ----------------------------------------------------
    $stampA2 = Get-W7AnnualStamp -Workbook $wb -Inspection $simInspection -P7 $p7 -Bank $bankA2
    $authorityA2 = Test-W7Authoritative -Stamp $stampA2 -P7 $p7 -ExpectedYears $shortYears
    $null = Add-W7Check ('A2 is authoritative for exactly ' + [string]$shortYears + ' years') `
        ([bool]$authorityA2.Ok) ([string]$authorityA2.Detail)
    $null = Add-W7Check ('the A2 annual stamp belongs to the A2 run') `
        ((Test-SimSameValue -A $stampA2['run_id'] -B $runIdA2) -and
         (Test-SimSameValue -A $stampA2['result_digest'] `
             -B $stateA2[('bank_' + $bankA2)]['result_digest'])) `
        ('stamp run ' + (Format-SimValue $stampA2['run_id']) + ', simulation run ' +
         (Format-SimValue $runIdA2))
    $captureA2 = Get-W7BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA2 -YearCount $shortYears
    $shrinkAxis = New-Object System.Collections.ArrayList
    for ($offset = 0; $offset -lt $shortYears; $offset++) {
        $expectedIndex = [double]($offset + 1)
        $expectedYear = [double]([int]$shortModel.timeline.start_year + $offset)
        if (-not (Test-SimExactDouble -Actual $captureA2[('annual.year' + [string]($offset + 1) +
                                                          '.project_index')] -Expected $expectedIndex)) {
            $null = $shrinkAxis.Add('row ' + [string]($offset + 1) + ' project_index')
        }
        if (-not (Test-SimExactDouble -Actual $captureA2[('annual.year' + [string]($offset + 1) +
                                                          '.calendar_year')] -Expected $expectedYear)) {
            $null = $shrinkAxis.Add('row ' + [string]($offset + 1) + ' calendar_year')
        }
    }
    $null = Add-W7Check ('rows 1 to ' + [string]$shortYears + ' carry the A2 answer') `
        ($shrinkAxis.Count -eq 0) (($shrinkAxis -join '; '))
    # THE WHOLE FORMER SPAN, ROW BY ROW. Not row 5 alone: every row A1 used to
    # occupy beyond the new count, in both index columns, every rung of both
    # ladders and both profile columns.
    $residue = Get-W7SurplusResidue -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA2 -FromRow ($shortYears + 1) -ToRow $longYears
    $null = Add-W7Check ('no A1 record survives anywhere in former rows ' +
                         [string]($shortYears + 1) + ' to ' + [string]$longYears) `
        (@($residue).Count -eq 0) ((@($residue)) -join '; ')
    $handoffA2 = Get-W7Handoff -Excel $excel -P7 $p7
    $null = Add-W7Check 'A2 reports both annual products CURRENT and the short year count' `
        (((([string]$handoffA2[$accessors[0]]) -ceq $currentDistributionState)) -and
         ((([string]$handoffA2[$accessors[1]]) -ceq $currentProfileState)) -and
         (([int]$handoffA2[$accessors[3]]) -eq $shortYears)) `
        ((Format-SimValue $handoffA2[$accessors[0]]) + ' / ' +
         (Format-SimValue $handoffA2[$accessors[1]]) + ' / ' +
         (Format-SimValue $handoffA2[$accessors[3]]))
    $null = Add-W7Check ('the A2 profile is stamped ' + $selectedLabel) `
        ((Test-SimExactText -Actual $stampA2['selected_px_label'] -Expected $selectedLabel) -and
         (Test-SimExactDouble -Actual $stampA2['selected_px_probability'] `
             -Expected $selectedProbability)) `
        ((Format-SimValue $stampA2['selected_px_label']) + ' / ' +
         (Format-SimValue $stampA2['selected_px_probability']))
    $iterationsA2 = Get-W7IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bankA2 -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsA2.Acquired
    $profileA2 = Get-W7ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankA2 -YearCount $shortYears -Measures $measures
    $null = Add-W7Check 'every A2 profile value is a number' `
        (@($profileA2.Problems).Count -eq 0) ((@($profileA2.Problems)) -join '; ')
    Invoke-W7Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsA2.Values -Bank $bankA2 `
        -Label $selectedLabel -Probability $selectedProbability -LadderIndex $selectedIndex `
        -ProfileSums $profileA2.Sums -ProfileScale $profileA2.Scale -Stage 'A2'

    # ---- BANK B SURVIVED A2 --------------------------------------------
    $captureBAfterA2 = Get-W7BankCapture -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bankB -YearCount $longYears
    $bAfterA2 = Compare-W7Surface -Before $captureB -After $captureBAfterA2
    $null = Add-W7Check ('bank ' + $bankB + ' is value-identical after A2 committed') `
        ($bAfterA2.Moved -eq 0) $bAfterA2.Detail
    $bIterationsAfter = Get-W7IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bankB -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$bIterationsAfter.Acquired
    $bMoved = Compare-W7IterationGrid -Before $iterationsB.Values -After $bIterationsAfter.Values
    $null = Add-W7Check ('no published iteration value in bank ' + $bankB +
                         ' changed when A2 committed') `
        ([string]::IsNullOrWhiteSpace($bMoved)) $bMoved

    # ---- THE FIXED-SEED NONCE DISCIPLINE, ACROSS ALL THREE RUNS --------
    $null = Add-W7Check 'no FIXED run consumed an AUTO nonce' `
        ((Test-SimBlank -Value $stateA1[('bank_' + $bankA)]['consumed_auto_nonce']) -and
         (Test-SimBlank -Value $stateB[('bank_' + $bankB)]['consumed_auto_nonce']) -and
         (Test-SimBlank -Value $stateA2[('bank_' + $bankA2)]['consumed_auto_nonce'])) `
        ('A1 ' + (Format-SimValue $stateA1[('bank_' + $bankA)]['consumed_auto_nonce']) +
         ', B ' + (Format-SimValue $stateB[('bank_' + $bankB)]['consumed_auto_nonce']) +
         ', A2 ' + (Format-SimValue $stateA2[('bank_' + $bankA2)]['consumed_auto_nonce']))
    $null = Add-W7Check 'the AUTO nonce and its pending marker never moved' `
        ((Test-SimSameValue -A $nonceStart -B $stateA2['shared']['next_auto_nonce']) -and
         (Test-SimSameValue -A $pendingStart -B $stateA2['pending_auto_nonce'])) `
        ((Format-SimValue $nonceStart) + ' -> ' +
         (Format-SimValue $stateA2['shared']['next_auto_nonce']) + ', pending ' +
         (Format-SimValue $pendingStart) + ' -> ' +
         (Format-SimValue $stateA2['pending_auto_nonce']))

    Write-W7Line ''
    Write-W7Line (Format-Phase6State -State $stateA2 -Label '  the run identity after A1, B and A2')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W7Line ''
    Write-W7Line ('THE W7 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W7Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W7Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W7Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W7Line ''
        Write-W7Line 'EXCEL SHUTDOWN'
        Write-W7Line '--------------'
        if ($naturalExit) {
            Write-W7Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W7'
            Write-W7Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W7Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W7Residual)) {
            Write-W7Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W7Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W7Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W7Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W7Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W7Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W7Check 'every COM release left 0 outstanding references' `
    (@($script:W7Residual).Count -eq 0) ((@($script:W7Residual)) -join '; ')
$null = Add-W7Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W7Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W7Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W7Line ''
Write-W7Line 'VERDICT'
Write-W7Line '-------'
Write-W7Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W7Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W7Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W7Line ''
if ($ok) {
    Write-W7Line 'W7: PASS'
} else {
    Write-W7Line 'W7: FAIL'
    Write-W7Line ''
    Write-W7Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W7Line ''
Write-W7Line ('report                 : ' + $script:W7Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W7Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
