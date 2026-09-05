<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W4 runner: the no-run annual refusal, then the
    deterministic FIXED-seed baseline.

.DESCRIPTION
    W4 ESTABLISHES THE BASELINE W5 AND W6 WILL ANALYSE, and it proves two things
    in one session because the state each needs exists exactly once:

      PART A  the annual endpoint REFUSES while no successful simulation exists,
              and that refusal moves nothing - no run identity, no AUTO nonce,
              no pending marker, no digest, no annual publication in either bank
      PART B  one FIXED-seed simulation then succeeds and publishes a stable
              CURRENT result whose identity W5 can bind to

    THE REFUSAL IS NOT A SEPARATE WINDOWS SCENARIO. The state it needs - a
    calculated model with no successful simulation - exists for a few seconds
    per session, immediately before the first run, so asking for it separately
    would mean building the whole fixture twice to observe the same thing.

    W4 STOPS AT A SUCCESSFUL SIMULATION. It never runs the annual endpoint
    successfully: the annual records, the stamp, the accessors on a current run
    and the profile reconciliation are W5's, and running them here would consume
    the very first-annual-run state W5 exists to observe.

    WHAT IS INDEPENDENT HERE, AND WHAT IS NOT - stated plainly because the
    difference decides what this scenario is worth.

      INDEPENDENT. The Phase-5 calculation the simulation runs on is compared
      against `pccm_builder.calc_oracle.calculate` through the same corpus W2
      and W4 used: every projected column of calc_drivers, calc_years,
      calc_inflation_factors and calc_annual, and the ten totals, at the
      accepted 1e-6 identity floor. The workbook's own published deterministic
      base - the anchor beside the simulation summary - is compared against that
      oracle's A totals too, which ties the simulation block to a pre-Phase-7
      authority rather than to itself.

      EXACT AND CONTRACT-OWNED. The effective seed, the supplied seed, the seed
      mode, the iteration count, the attempt result, the simulation state and
      the active bank are compared against the projection's own vocabulary and
      the corpus's own request. The published summary minimum and maximum are
      SELECTIONS from the iteration column, so they are compared to the exact
      minimum and maximum of that column - no tolerance is needed or used.

      NOT AVAILABLE, AND NOT FAKED. There is no independent stochastic
      expectation for THIS model. The accepted Phase-6 Gate-B oracle owns four
      analytical parity plan cases at supplied seed 12345, not a whole-workbook
      simulation of the W4 fixture at seed 20260905, so it cannot speak for this
      run's mean or percentile ladder. W4 therefore does not compare them
      against an oracle, and it does not manufacture one by re-running itself
      and calling the agreement independence.

    THE SHAPE IS W2's AND W4's, and the seven functions the three share are
    pinned to W2's so the mandated reuse cannot drift into a third behaviour.

    TEN HELPERS ARE COPIED RATHER THAN DOT-SOURCED, for the reason W1 defect #2
    established: the accepted fixture calls them, and their owner
    (`phase4_functional_test.ps1`) is NOT definition-only - dot-sourcing it would
    run the entire Phase-4 matrix. They are copied byte for byte from
    `phase7_timing_scenarios.ps1`, a control pins them to it, and a second
    control walks the whole call closure so a missing helper is found here
    rather than by a failed Windows session.

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
$script:W4Lines = New-Object System.Collections.ArrayList
$script:W4Path = ''
$script:W4Checks = New-Object System.Collections.ArrayList

function Write-W4Line {
    param([string]$Text = '')
    $null = $script:W4Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W4Path)) {
        try {
            Set-Content -LiteralPath $script:W4Path `
                -Value ($script:W4Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W4 had to perform to reach
# its subject - from a RESULT, which is W4's own claim about the calculation. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W4Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W4Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W4Line $line
    return $Ok
}

function Format-W4Value {
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
$script:W4Residual = New-Object System.Collections.ArrayList

function Invoke-W4Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W4Residual.Add(
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
function Get-W4SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W4 result could not be attributed to a source revision')
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
function Compare-W4Cell {
    param($Got, $Expect, [double]$Allowance)
    if ($null -eq $Expect) {
        if (($null -eq $Got) -or ($Got -is [System.DBNull]) -or
            (($Got -is [string]) -and ([string]$Got).Length -eq 0)) { return '' }
        return ('published ' + (Format-W4Value $Got) + ' where the oracle has none')
    }
    if ($Expect -is [string]) {
        if (($Got -is [string]) -and (([string]$Got) -ceq ([string]$Expect))) { return '' }
        return ((Format-W4Value $Got) + ' vs ' + [string]$Expect)
    }
    if ($Got -isnot [double]) {
        return ((Format-W4Value $Got) + ' is not a number')
    }
    $difference = [Math]::Abs([double]$Got - [double]$Expect)
    if ($difference -le $Allowance) { return '' }
    return ([string]$Got + ' vs ' + [string]$Expect + ' (difference ' + [string]$difference + ')')
}

# THE PUBLISHED TABLE, ROW FOR ROW, EVERY PROJECTED COLUMN. A failing table
# reports the first twelve disagreements and the count: the count says how bad
# it is, the examples make it diagnosable, and 6,300 lines would do neither.
function Compare-W4Table {
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
            $problem = Compare-W4Cell -Got $Live[$index][$at] -Expect $expect -Allowance $Allowance
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
    return (Add-W4Check ($Label + ': ' + $TableKey +
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
function Get-W4AnnualStamp {
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
function Get-W4AnnualFirstRecord {
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

function Get-W4Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

# THE ANNUAL ATTEMPT, THROUGH THE PUBLISHED ENDPOINT AND ITS PUBLISHED ANSWER.
# The runner never judges the outcome itself: it asks production what happened.
function Invoke-W4Annual {
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
$script:W4DerivedRows = @('simulation_status', 'status_evaluated_at')

function Get-W4RunInvariants {
    param($Workbook, $Inspection)
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($script:W4DerivedRows -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-W4InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-W4Check ($Label + ': no run identity, nonce, pending marker or digest moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# A defined name that carries TEXT. Set-NamedValue writes a Double - PowerShell
# binds a COM property call site per argument type, so one polymorphic
# assignment silently fails the second time it is reached with another type -
# and the selected confidence level is a label, not a number.
function Set-W4NamedText {
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
function Get-W4IterationBlock {
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
        if ($null -ne $range)  { Invoke-W4Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-W4Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W4Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W4 (the refusal, then the FIXED-seed baseline)' -ForegroundColor Cyan
Write-Host '============================================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath  = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath   = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$gateBCasePath = Join-Path $BuildDir 'phase6_gate_b_cases.json'
$p7InspectPath = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$casesPath     = Join-Path $BuildDir 'phase7_acceptance_cases.json'
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

$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no W4 scenario.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-W4SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W4 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w4-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W4Path = Join-Path $tempRoot 'phase7_w4_base_simulation.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$model = $case.model
$driverCount = @($model.cost_lines).Count + @($model.risks).Count
$yearCount = [int]$model.timeline.duration
$allowance = [double]$cases.provenance.comparison_absolute_floor
$iterations = [int]$case.iterations
$suppliedSeed = [double]$case.supplied_seed

Write-W4Line 'PCCM - PHASE 7 W4: THE NO-RUN REFUSAL, THEN THE FIXED-SEED BASELINE'
Write-W4Line '=================================================================='
Write-W4Line ''
Write-W4Line 'This is the MINIMAL W4 runner. It establishes the baseline W5 and W6 will'
Write-W4Line 'analyse and STOPS at a successful simulation: the annual endpoint is never'
Write-W4Line 'run successfully here. No Gate-B result is produced or implied, and the'
Write-W4Line 'historical Phase-6 runtime authority remains Run 6 on its own closure commit.'
Write-W4Line ''
Write-W4Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W4Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W4Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W4Line ('git HEAD               : ' + [string]$revision.Head)
Write-W4Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W4Line ('model version          : ' + [string]$manifest.model_version)
Write-W4Line ('calc contract version  : ' + [string]$cases.provenance.calc_contract_version)
Write-W4Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W4Line ('build directory        : ' + $BuildDir)
Write-W4Line ('working copy           : ' + $tempRoot)
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
    Write-W4Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W4Line ''
Write-W4Line 'THE SCENARIO AND ITS AUTHORITY'
Write-W4Line '------------------------------'
Write-W4Line ('  dimension under test : ' + [string]$case.dimension)
Write-W4Line ('  drivers              : ' + [string]$driverCount +
              ' (' + [string]@($model.cost_lines).Count + ' cost lines, ' +
              [string]@($model.risks).Count + ' risks)')
Write-W4Line ('  project years        : ' + [string]$yearCount)
Write-W4Line ('  seed mode            : ' + [string]$case.seed_mode +
              ', supplied seed ' + [string]$case.supplied_seed)
Write-W4Line ('  iterations           : ' + [string]$iterations)
Write-W4Line ('  selected Px          : ' + [string]$case.selected_confidence_level)
Write-W4Line ('  expectation source   : ' + [string]$cases.provenance.expectation_source +
              ' (the deterministic Phase-5 answer the simulation runs on)')
Write-W4Line ('  comparison allowance : ' + [string]$allowance +
              ' (the accepted Phase-5 identity absolute floor; no tolerance is invented here)')
Write-W4Line ''
Write-W4Line 'THERE IS NO INDEPENDENT STOCHASTIC EXPECTATION FOR THIS MODEL, and W4 does'
Write-W4Line 'not pretend otherwise. The accepted Phase-6 Gate-B oracle owns four'
Write-W4Line 'analytical parity plan cases at supplied seed 12345, not a whole-workbook'
Write-W4Line 'simulation of this fixture at seed 20260905. So the mean and the percentile'
Write-W4Line 'ladder are recorded and NOT compared against an oracle here; what is'
Write-W4Line 'compared is what an authority actually owns.'
Write-W4Line ''

$null = Add-W4Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W4Line ''
    Write-W4Line 'STOP. Excel was never started for the W4 session, and nothing was accepted.'
    Write-W4Line ('report                 : ' + $script:W4Path)
    Write-W4Line 'W4: FAIL'
    exit 1
}

$null = Add-W4Check 'the expectations come from the independent Phase-5 oracle' `
    (([string]$cases.provenance.expectation_source) -ceq 'pccm_builder.calc_oracle.calculate') `
    ([string]$cases.provenance.expectation_source) 'PREREQUISITE'
$null = Add-W4Check 'the W4 case is the behavioural fixture' `
    (([string]$case.dimension) -ceq 'behavioural') ([string]$case.dimension) 'PREREQUISITE'
# THE FIXTURE IS SMALL AND DETERMINISTIC ON PURPOSE. A behavioural baseline that
# had grown into a performance case would be a different scenario.
$null = Add-W4Check 'the W4 fixture is the authorised size and request' `
    (($driverCount -eq 5) -and ($yearCount -eq 4) -and ($iterations -eq 1000) -and
     (([string]$case.seed_mode) -ceq 'FIXED') -and ($suppliedSeed -eq 20260905)) `
    ([string]$driverCount + ' drivers over ' + [string]$yearCount + ' years, ' +
     [string]$iterations + ' iterations, ' + [string]$case.seed_mode + ' seed ' +
     [string]$case.supplied_seed) 'PREREQUISITE'
# AND THE REQUEST IS VOCABULARY THE CONTRACT OWNS, not a word typed here.
$null = Add-W4Check 'FIXED is a seed mode the sim contract projects' `
    (@($gateBCases.vocabulary.seed_modes) -contains ([string]$case.seed_mode)) `
    ((@($gateBCases.vocabulary.seed_modes)) -join ', ') 'PREREQUISITE'
$null = Add-W4Check 'the selected confidence level is a projected quantile label' `
    (@($gateBCases.vocabulary.quantile_labels) -contains ([string]$case.selected_confidence_level)) `
    ([string]$case.selected_confidence_level) 'PREREQUISITE'

# ===========================================================================
# THE W4 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 7 W4 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W4Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W4Line 'ever terminated, and the workbook is never saved.'
    Write-W4Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    Write-W4Line 'PREREQUISITES'
    Write-W4Line '-------------'

    # THE COMPILE TRIGGER, DELIBERATELY LIGHT. VBA compiles the whole project
    # before it executes any statement, so the first Application.Run IS a
    # full-project compile - and a published READ accessor runs nothing and
    # consumes no run identity. W1 owns the surface matrix.
    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $compiled = [string]::IsNullOrWhiteSpace($compileFailure)
    $null = Add-W4Check 'the current VBAProject compiles in real Excel' `
        $compiled $compileFailure 'PREREQUISITE'
    if (-not $compiled) {
        Write-W4Line ''
        Write-W4Line 'STOP. Nothing observed after a compile failure is evidence about anything.'
        throw ('the VBAProject does not compile: ' + $compileFailure)
    }

    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # THE BEHAVIOURAL FIXTURE: the corpus model through the accepted Phase-5
    # choreography, then the simulation REQUEST through the projected controls.
    # A named seed makes the request FIXED, so the effective seed is a number
    # this report can state and W5 can require to be unchanged - and no AUTO
    # nonce is consumed by the fixture or by the run.
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
        Set-W4NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.selected_confidence_level.defined_name) `
            -Value ([string]$case.selected_confidence_level)
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-W4Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) {
        throw ('the W4 model was not applied: ' + $applied + $fixtureFailure)
    }

    $calc = ''
    $calcFailure = ''
    try {
        $calc = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'W4 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-W4Check 'PCCM_Calculate succeeded on the behavioural model' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) ($calc + $calcFailure) 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) {
        throw ('PCCM_Calculate did not succeed: ' + $calcFailure)
    }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

    # -------------------------------------------------------------------
    # THE DETERMINISTIC BASE, AGAINST THE INDEPENDENT ORACLE
    # -------------------------------------------------------------------
    # The simulation samples around this calculation, so a wrong base would make
    # every stochastic observation below meaningless. This is the same
    # comparison W2 and W3 make, on a five-driver, four-year model.
    Write-W4Line ''
    Write-W4Line 'THE DETERMINISTIC BASE, AGAINST THE INDEPENDENT PHASE-5 ORACLE'
    Write-W4Line '--------------------------------------------------------------'
    $appliedTimeline = Get-CalcScalar -Workbook $wb -Inspection $inspection `
        -Block 'calc_state' -FieldKey 'last_successful_applied_timeline'
    $null = Add-W4Check 'the applied timeline is the one the oracle computed against' `
        ((([string]$appliedTimeline) -ceq ([string]$case.expected.applied_timeline))) `
        ((Format-W4Value $appliedTimeline) + ' vs ' + [string]$case.expected.applied_timeline)

    $inflationMap = @{ 'inflation_profile' = 'profile'; 'cumulative_inflation_factor' = 'cumulative_factor' }
    $surfaces = @(
        [pscustomobject]@{ Key = 'calc_drivers';            Expected = $case.expected.drivers; Map = $null },
        [pscustomobject]@{ Key = 'calc_years';              Expected = $case.expected.calc_years; Map = $null },
        [pscustomobject]@{ Key = 'calc_inflation_factors';  Expected = $case.expected.inflation_factors; Map = $inflationMap },
        [pscustomobject]@{ Key = 'calc_annual';             Expected = $case.expected.annual; Map = $null }
    )
    foreach ($surface in $surfaces) {
        $key = [string]$surface.Key
        $live = @(Get-CalcTableRows -Workbook $wb -Inspection $inspection -TableKey $key)
        $want = @($surface.Expected)
        $sized = Add-W4Check ($key + ' publishes one row per expected row') `
            ($live.Count -eq $want.Count) `
            ('published ' + [string]$live.Count + ', expected ' + [string]$want.Count)
        if ($sized) {
            $null = Compare-W4Table -Live $live -Expected $want -Inspection $inspection `
                -TableKey $key -Allowance $allowance -Label 'W4' -FieldMap $surface.Map
        }
    }
    $totalProblems = @()
    foreach ($key in $case.expected.totals.PSObject.Properties.Name) {
        $got = Get-CalcScalar -Workbook $wb -Inspection $inspection `
            -Block 'calc_totals' -FieldKey $key
        $problem = Compare-W4Cell -Got $got -Expect $case.expected.totals.$key -Allowance $allowance
        if (-not [string]::IsNullOrWhiteSpace($problem)) {
            $totalProblems += ([string]$key + ': ' + $problem)
        }
    }
    $null = Add-W4Check 'the ten published totals match the accepted Phase-5 oracle' `
        ($totalProblems.Count -eq 0) ($totalProblems -join '; ')

    # -------------------------------------------------------------------
    # PART A - THE ANNUAL ENDPOINT REFUSES WITH NO SUCCESSFUL SIMULATION
    # -------------------------------------------------------------------
    Write-W4Line ''
    Write-W4Line 'PART A - THE NO-RUN REFUSAL'
    Write-W4Line '---------------------------'
    $statusBefore = [string]$excel.Run('PCCM_SimulationStatus')
    $null = Add-W4Check 'no simulation has been published yet' `
        ([string]::IsNullOrEmpty($statusBefore)) `
        ('PCCM_SimulationStatus returned ' + (Format-SimValue $statusBefore))

    $stateBefore = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $before = Get-W4RunInvariants -Workbook $wb -Inspection $simInspection

    $announcement = Invoke-W4Annual -Excel $excel -P7 $p7
    $null = Add-W4Check 'the annual endpoint REFUSES with no successful simulation' `
        ($announcement -like 'FAIL|*') $announcement

    $after = Get-W4RunInvariants -Workbook $wb -Inspection $simInspection
    $null = Add-W4InvariantChecks 'the refused annual attempt' $before $after
    $stateAfter = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $null = Add-W4Check 'the refused attempt left the derived simulation status word unchanged' `
        (Test-SimSameValue -A $stateBefore['shared']['simulation_status'] `
                           -B $stateAfter['shared']['simulation_status']) `
        ((Format-SimValue $stateBefore['shared']['simulation_status']) + ' -> ' +
         (Format-SimValue $stateAfter['shared']['simulation_status']))

    # NOTHING WAS PUBLISHED IN EITHER BANK: not a stamp, not a record row.
    foreach ($bank in @($simInspection.publication.bank_labels)) {
        $stamp = Get-W4AnnualStamp -Workbook $wb -Inspection $simInspection -P7 $p7 -Bank $bank
        $written = @()
        foreach ($key in $stamp.Keys) {
            if (-not (Test-SimBlank -Value $stamp[$key])) {
                $written += ([string]$key + ' = ' + (Format-SimValue $stamp[$key]))
            }
        }
        $null = Add-W4Check ('the refused attempt wrote no annual stamp in bank ' + $bank) `
            ($written.Count -eq 0) ($written -join '; ')

        $record = Get-W4AnnualFirstRecord -Workbook $wb -Inspection $simInspection -P7 $p7 -Bank $bank
        $filled = @()
        foreach ($key in $record.Keys) {
            if (-not (Test-SimBlank -Value $record[$key])) {
                $filled += ([string]$key + ' = ' + (Format-SimValue $record[$key]))
            }
        }
        $null = Add-W4Check ('the refused attempt published no annual record in bank ' + $bank) `
            ($filled.Count -eq 0) ($filled -join '; ')
    }

    # AND THE PUBLISHED HANDOFF STILL ANSWERS THE UNRUN STATE.
    $handoff = Get-W4Handoff -Excel $excel -P7 $p7
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $notProduced = [string]$p7.handoff.distribution_states[0]
    $null = Add-W4Check 'the annual distribution state is still NOT PRODUCED' `
        ((([string]$handoff[$accessors[0]]) -ceq $notProduced)) `
        (Format-SimValue $handoff[$accessors[0]])
    $null = Add-W4Check 'the annual profile state is still NOT PRODUCED' `
        ((([string]$handoff[$accessors[1]]) -ceq ([string]$p7.handoff.profile_states[0]))) `
        (Format-SimValue $handoff[$accessors[1]])
    $null = Add-W4Check 'no annual profile Px is reported' `
        (Test-SimBlank -Value $handoff[$accessors[2]]) (Format-SimValue $handoff[$accessors[2]])
    $null = Add-W4Check 'no fake year count is reported' `
        (0 -eq [int]$handoff[$accessors[3]]) (Format-SimValue $handoff[$accessors[3]])

    # THE REFUSAL DID NOT DAMAGE THE DETERMINISTIC CALCULATION EITHER.
    $calcStatus = [string]$excel.Run('PCCM_CalculationStatus')
    $null = Add-W4Check 'the deterministic calculation is still CURRENT after the refusal' `
        ($calcStatus -ceq 'CURRENT') $calcStatus

    Write-W4Line ''
    Write-W4Line (Format-Phase6State -State $stateAfter -Label '  the run identity after the refusal')

    # -------------------------------------------------------------------
    # PART B - THE FIXED-SEED BASELINE
    # -------------------------------------------------------------------
    Write-W4Line ''
    Write-W4Line 'PART B - THE FIXED-SEED SIMULATION BASELINE'
    Write-W4Line '------------------------------------------'
    $nonceBefore = $stateAfter['shared']['next_auto_nonce']
    $pendingBefore = $stateAfter['pending_auto_nonce']

    $result = Invoke-Phase6Simulation -Excel $excel
    $ran = Add-W4Check 'PCCM_RunSimulation succeeded' `
        (Test-Phase6Announced -Result $result -Kind 'OK') $result
    if (-not $ran) { throw ('PCCM_RunSimulation did not succeed: ' + $result) }

    $status = [string]$excel.Run('PCCM_SimulationStatus')
    $null = Add-W4Check 'the simulation reports CURRENT' ($status -ceq 'CURRENT') $status
    $null = Add-W4Check 'CURRENT is a state the sim contract projects' `
        (@($gateBCases.vocabulary.sim_states) -contains $status) `
        ((@($gateBCases.vocabulary.sim_states)) -join ', ')

    $state = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bank = Get-Phase6ActiveBank -State $state
    $null = Add-W4Check 'a publication bank is active' `
        ((-not [string]::IsNullOrEmpty($bank)) -and
         (@($simInspection.publication.bank_labels) -contains $bank)) $bank
    if ([string]::IsNullOrEmpty($bank)) { throw 'no publication bank was activated' }
    $block = $state[('bank_' + $bank)]

    # THE REQUEST, EXACTLY AS ASKED FOR.
    $null = Add-W4Check 'the effective seed is exactly the FIXED seed requested' `
        (Test-SimExactDouble -Actual $block['effective_seed'] -Expected $suppliedSeed) `
        ('published ' + (Format-SimValue $block['effective_seed']) + ', requested ' +
         [string]$case.supplied_seed)
    $null = Add-W4Check 'the supplied seed is republished unchanged' `
        (Test-SimExactDouble -Actual $block['supplied_seed'] -Expected $suppliedSeed) `
        (Format-SimValue $block['supplied_seed'])
    $null = Add-W4Check 'the run records the FIXED seed mode' `
        (Test-SimExactText -Actual $block['seed_mode'] -Expected ([string]$case.seed_mode)) `
        (Format-SimValue $block['seed_mode'])
    $null = Add-W4Check 'the accepted iteration count is the requested one' `
        (Test-SimExactDouble -Actual $block['iterations_run'] -Expected ([double]$iterations)) `
        (Format-SimValue $block['iterations_run'])

    # THE IDENTITY W5 WILL BIND TO.
    foreach ($field in @('run_id', 'request_fingerprint', 'result_digest',
                         'last_successful_stamp', 'applied_timeline')) {
        $null = Add-W4Check ('the run published a ' + $field) `
            (-not (Test-SimBlank -Value $block[$field])) (Format-SimValue $block[$field])
    }
    $null = Add-W4Check 'the shared last run id is the run just published' `
        (Test-SimSameValue -A $state['shared']['last_run_id'] -B $block['run_id']) `
        ((Format-SimValue $state['shared']['last_run_id']) + ' vs ' +
         (Format-SimValue $block['run_id']))
    $null = Add-W4Check 'the attempt is recorded as a success' `
        (Test-SimExactText -Actual $state['shared']['last_attempt_result'] -Expected 'SUCCESS') `
        (Format-SimValue $state['shared']['last_attempt_result'])
    $null = Add-W4Check 'SUCCESS is an attempt result the sim contract projects' `
        (@($gateBCases.vocabulary.attempt_results) -contains 'SUCCESS') `
        ((@($gateBCases.vocabulary.attempt_results)) -join ', ')

    # THE FIXED-SEED NONCE CONTRACT. The AUTO nonce lifecycle governs AUTO runs;
    # a FIXED request allocates nothing, so the nonce, its durable pending
    # marker and the consumed field must all say so. An AUTO nonce change is NOT
    # expected here and would be a defect, not a normal outcome.
    $null = Add-W4Check 'the FIXED run consumed no AUTO nonce' `
        (Test-SimBlank -Value $block['consumed_auto_nonce']) `
        (Format-SimValue $block['consumed_auto_nonce'])
    $null = Add-W4Check 'the FIXED run did not advance the AUTO nonce' `
        (Test-SimSameValue -A $nonceBefore -B $state['shared']['next_auto_nonce']) `
        ((Format-SimValue $nonceBefore) + ' -> ' + (Format-SimValue $state['shared']['next_auto_nonce']))
    $null = Add-W4Check 'the FIXED run left the pending AUTO nonce marker alone' `
        (Test-SimSameValue -A $pendingBefore -B $state['pending_auto_nonce']) `
        ((Format-SimValue $pendingBefore) + ' -> ' + (Format-SimValue $state['pending_auto_nonce']))

    # -------------------------------------------------------------------
    # THE PUBLISHED ITERATION COLUMN, AND WHAT CAN BE CHECKED EXACTLY
    # -------------------------------------------------------------------
    $iterationBlock = Get-W4IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationBlock.Acquired
    $grid = $iterationBlock.Values
    $rows = 0
    if ($null -ne $grid) { $rows = [int]$grid.GetLength(0) }
    $null = Add-W4Check ('the iteration column carries ' + [string]$iterations + ' rows') `
        ($rows -eq $iterations) ('read ' + [string]$rows)

    if ($rows -eq $iterations) {
        $indexProblem = ''
        $nonNumeric = 0
        $minimum = [double]::PositiveInfinity
        $maximum = [double]::NegativeInfinity
        for ($row = 1; $row -le $rows; $row++) {
            $index = $grid[$row, 1]
            if (-not (Test-SimExactDouble -Actual $index -Expected ([double]$row))) {
                if ([string]::IsNullOrWhiteSpace($indexProblem)) {
                    $indexProblem = ('row ' + [string]$row + ' carries iteration index ' +
                                     (Format-SimValue $index))
                }
            }
            $nominal = $grid[$row, 2]
            $pv = $grid[$row, 3]
            if (($nominal -isnot [double]) -or ($pv -isnot [double])) {
                $nonNumeric = $nonNumeric + 1
                continue
            }
            if ([double]$nominal -lt $minimum) { $minimum = [double]$nominal }
            if ([double]$nominal -gt $maximum) { $maximum = [double]$nominal }
        }
        $null = Add-W4Check 'the published iteration indices are 1..N with no hole' `
            ([string]::IsNullOrWhiteSpace($indexProblem)) $indexProblem
        $null = Add-W4Check 'every published iteration total is a number in both measures' `
            ($nonNumeric -eq 0) ([string]$nonNumeric + ' rows published a non-number')

        # THE TWO SUMMARY ROWS THAT NEED NO TOLERANCE. A minimum and a maximum
        # are SELECTIONS from the column above, not accumulations, so they are
        # compared exactly. The mean and the ladder are accumulations with no
        # independent expectation for this model, so they are recorded below and
        # not compared.
        if ($nonNumeric -eq 0) {
            $publishedMin = Get-SimSummaryValue -Workbook $wb -Inspection $simInspection `
                -Bank $bank -Measure 'nominal' -RowKey 'minimum'
            $publishedMax = Get-SimSummaryValue -Workbook $wb -Inspection $simInspection `
                -Bank $bank -Measure 'nominal' -RowKey 'maximum'
            $null = Add-W4Check 'the published nominal minimum is the minimum of the published column' `
                (Test-SimExactDouble -Actual $publishedMin -Expected $minimum) `
                ('published ' + (Format-SimValue $publishedMin) + ', column ' + [string]$minimum)
            $null = Add-W4Check 'the published nominal maximum is the maximum of the published column' `
                (Test-SimExactDouble -Actual $publishedMax -Expected $maximum) `
                ('published ' + (Format-SimValue $publishedMax) + ', column ' + [string]$maximum)
        }
    }

    # THE DETERMINISTIC ANCHOR BESIDE THE SIMULATION, AGAINST THE INDEPENDENT
    # ORACLE. This is the one number in the simulation block that a pre-Phase-7
    # authority owns, and it ties the block to something other than itself.
    foreach ($measure in @(
        [pscustomobject]@{ Name = 'nominal'; Key = 'a_nom' },
        [pscustomobject]@{ Name = 'pv';      Key = 'a_pv' })) {
        $measureName = [string]$measure.Name
        $totalKey = [string]$measure.Key
        $published = Get-SimSummaryValue -Workbook $wb -Inspection $simInspection `
            -Bank $bank -Measure $measureName -RowKey 'deterministic_base_a'
        $problem = Compare-W4Cell -Got $published -Expect $case.expected.totals.$totalKey `
            -Allowance $allowance
        $null = Add-W4Check ('the published deterministic base (' + $measureName +
                             ') matches the independent Phase-5 oracle') `
            ([string]::IsNullOrWhiteSpace($problem)) $problem
    }

    # RECORDED, NOT COMPARED: the accumulations no authority owns for this model.
    Write-W4Line ''
    Write-W4Line '  the summary ladder, RECORDED and not compared (no oracle owns it here):'
    foreach ($rowKey in @('mean', 'sample_standard_deviation', 'minimum', 'maximum')) {
        $nominal = Get-SimSummaryValue -Workbook $wb -Inspection $simInspection `
            -Bank $bank -Measure 'nominal' -RowKey $rowKey
        $pv = Get-SimSummaryValue -Workbook $wb -Inspection $simInspection `
            -Bank $bank -Measure 'pv' -RowKey $rowKey
        Write-W4Line ('    ' + $rowKey.PadRight(28) + ' nominal ' + (Format-SimValue $nominal) +
                      '  pv ' + (Format-SimValue $pv))
    }

    Write-W4Line ''
    Write-W4Line (Format-Phase6State -State $state -Label '  THE BASELINE IDENTITY W5 WILL BIND TO')

    # W4 STOPS HERE. The annual endpoint is NOT run successfully: the annual
    # records, the stamp, the accessors on a current run and the profile
    # reconciliation are W5's subject, and producing them here would consume the
    # first-annual-run state W5 exists to observe.
    Write-W4Line ''
    Write-W4Line 'W4 STOPS AT A SUCCESSFUL SIMULATION. The annual endpoint was NOT run'
    Write-W4Line 'successfully in this session; that is W5.'

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W4Line ''
    Write-W4Line ('THE W4 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W4Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W4Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W4Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W4Line ''
        Write-W4Line 'EXCEL SHUTDOWN'
        Write-W4Line '--------------'
        if ($naturalExit) {
            Write-W4Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W4'
            Write-W4Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W4Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W4Residual)) {
            Write-W4Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W4Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W4Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W4Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W4Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W4Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W4Check 'every COM release left 0 outstanding references' `
    (@($script:W4Residual).Count -eq 0) ((@($script:W4Residual)) -join '; ')
$null = Add-W4Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W4Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W4Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W4Line ''
Write-W4Line 'VERDICT'
Write-W4Line '-------'
Write-W4Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W4Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W4Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W4Line ''
if ($ok) {
    Write-W4Line 'W4: PASS'
} else {
    Write-W4Line 'W4: FAIL'
    Write-W4Line ''
    Write-W4Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W4Line ''
Write-W4Line ('report                 : ' + $script:W4Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W4Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
