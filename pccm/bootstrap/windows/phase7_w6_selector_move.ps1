<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W6 runner: the reporting selector moves, and only
    the profile follows it.

.DESCRIPTION
    W6 PROVES THE SELECTOR SEMANTIC, and that semantic is a DIFFERENCE between
    two products that share a sheet. sim_contract.yaml states it as four
    booleans, projected into the acceptance inspection so this runner reports
    the rule it is testing rather than encoding it:

        distribution_currentness_is_selector_specific : false
        profile_currentness_is_selector_specific      : true
        profile_relabelled_on_selector_change         : false
        selector_change_requires_new_simulation       : false

    The per-year percentile LADDERS are a property of the run alone - every rung
    of every year is computed across all iterations and no selector enters them
    - so moving the reporting selector cannot make them wrong and must not make
    them stale. The selected-Px PROFILE is the blend at ONE resolved Px: it
    stays historically valid for the Px stamped into it and stops being current
    the moment the selector resolves to a different one, and it is NEVER
    relabelled.

    SO W6 IS TWO MOVES, IN ONE SESSION, ON ONE SIMULATION.

      PART A  the selector moves P80 -> P50 through the ordinary user input and
              NOTHING ELSE RUNS. The simulation must be untouched down to the
              last iteration value; the distributions must still be CURRENT and
              value-identical; the old profile must still be there, still
              stamped P80 at 0.8, still reported as belonging to another Px -
              not cleared, not rewritten, not relabelled.
      PART B  the annual endpoint alone reruns. The simulation must STILL be
              untouched; the distributions must STILL be value-identical, rung
              for rung, year for year, because a reporting selector may not
              recompute a distribution; and the profile must now be stamped P50
              and reconcile to the P50 total.

    IT IS NOT A NEW STOCHASTIC RUN, and the report proves that rather than
    asserting it: same active bank, same run_id, same request fingerprint, same
    result digest, no AUTO nonce consumed, no pending marker moved, and every
    one of the published iteration values identical before and after both moves.

    THE RECONCILIATION IS W6's, AT A DIFFERENT RUNG. The total is read from the
    block the projection names as the total percentile block and cross-checked
    against the contract's own Type-7 over the published iteration column; the
    profile sum must equal that total; and the contingency rung must equal the
    contract's formula against the SAME measure's deterministic base. The
    allowance is the project's accepted identity rule, read from the corpus.
    There is no independent annual Windows oracle and none is claimed.

    ONE THING IS DELIBERATELY NOT A FAILURE. The P50 profile is expected to
    differ from the P80 one, but a fixture COULD produce equal values, and
    failing on equality would be asserting a property of this fixture rather
    than of the code. So the change is REPORTED, and correctness is proved by
    the reconciliation and by the stamp - which must move.

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
$script:W6Lines = New-Object System.Collections.ArrayList
$script:W6Path = ''
$script:W6Checks = New-Object System.Collections.ArrayList

function Write-W6Line {
    param([string]$Text = '')
    $null = $script:W6Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W6Path)) {
        try {
            Set-Content -LiteralPath $script:W6Path `
                -Value ($script:W6Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W6 had to perform to reach
# its subject - from a RESULT, which is W6's own claim about the calculation. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W6Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W6Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W6Line $line
    return $Ok
}

function Format-W6Value {
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
$script:W6Residual = New-Object System.Collections.ArrayList

function Invoke-W6Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W6Residual.Add(
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
function Get-W6SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W6 result could not be attributed to a source revision')
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
function Compare-W6Cell {
    param($Got, $Expect, [double]$Allowance)
    if ($null -eq $Expect) {
        if (($null -eq $Got) -or ($Got -is [System.DBNull]) -or
            (($Got -is [string]) -and ([string]$Got).Length -eq 0)) { return '' }
        return ('published ' + (Format-W6Value $Got) + ' where the oracle has none')
    }
    if ($Expect -is [string]) {
        if (($Got -is [string]) -and (([string]$Got) -ceq ([string]$Expect))) { return '' }
        return ((Format-W6Value $Got) + ' vs ' + [string]$Expect)
    }
    if ($Got -isnot [double]) {
        return ((Format-W6Value $Got) + ' is not a number')
    }
    $difference = [Math]::Abs([double]$Got - [double]$Expect)
    if ($difference -le $Allowance) { return '' }
    return ([string]$Got + ' vs ' + [string]$Expect + ' (difference ' + [string]$difference + ')')
}

# THE PUBLISHED TABLE, ROW FOR ROW, EVERY PROJECTED COLUMN. A failing table
# reports the first twelve disagreements and the count: the count says how bad
# it is, the examples make it diagnosable, and 6,300 lines would do neither.
function Compare-W6Table {
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
            $problem = Compare-W6Cell -Got $Live[$index][$at] -Expect $expect -Allowance $Allowance
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
    return (Add-W6Check ($Label + ': ' + $TableKey +
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
function Get-W6AnnualStamp {
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
function Get-W6AnnualFirstRecord {
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

function Get-W6Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

# THE ANNUAL ATTEMPT, THROUGH THE PUBLISHED ENDPOINT AND ITS PUBLISHED ANSWER.
# The runner never judges the outcome itself: it asks production what happened.
function Invoke-W6Annual {
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
$script:W6DerivedRows = @('simulation_status', 'status_evaluated_at')

function Get-W6RunInvariants {
    param($Workbook, $Inspection)
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($script:W6DerivedRows -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-W6InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-W6Check ($Label + ': no run identity, nonce, pending marker or digest moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# A defined name that carries TEXT. Set-NamedValue writes a Double - PowerShell
# binds a COM property call site per argument type, so one polymorphic
# assignment silently fails the second time it is reached with another type -
# and the selected confidence level is a label, not a number.
function Set-W6NamedText {
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
function Get-W6IterationBlock {
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
        if ($null -ne $range)  { Invoke-W6Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-W6Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W6Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}


# ===========================================================================
# COLUMN ARITHMETIC, AND THE ANNUAL LADDER IT REACHES
# ===========================================================================
# The projection gives the FIRST ladder column per bank per measure and the
# ladder's LENGTH; the eleven columns follow it. That is the contract's shape,
# so the offset is computed rather than eleven letters being typed in.
# THE DEFECT THAT STOPPED THE FIRST W6 RUN, and the reason it survived every
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
function ConvertTo-W6ColumnNumber {
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
function ConvertFrom-W6ColumnNumber {
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
function Get-W6AnnualRecord {
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
        $first = ConvertTo-W6ColumnNumber -Letters ([string]$records.quantile_first_column.$Bank.$measure)
        $ladder = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $count; $index++) {
            $column = ConvertFrom-W6ColumnNumber -Number ($first + $index)
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
function Get-W6AnnualRegionIndex {
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
        if ($null -ne $range)  { Invoke-W6Release $Ledger $range  'Range(annual region)'; $range  = $null }
        if ($null -ne $sheet)  { Invoke-W6Release $Ledger $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W6Release $Ledger $sheets 'Worksheets';           $sheets = $null }
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
function Get-W6Type7Value {
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
function Get-W6IdentityAllowance {
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
function Get-W6AnnualSurface {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount)
    $surface = New-Object System.Collections.Specialized.OrderedDictionary
    $stamp = Get-W6AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 -Bank $Bank
    foreach ($key in $stamp.Keys) { $surface.Add(('stamp.' + [string]$key), $stamp[$key]) }
    $ladderCount = [int]$P7.annual_records.quantile_count
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W6AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
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
function Compare-W6Surface {
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
function Compare-W6IterationGrid {
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
# second behaviour, and the point of W6 is that only the rung changes.
function Invoke-W6Reconciliation {
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
        $derived = Get-W6Type7Value -Values ([double[]]@($totals)) -Probability $Probability

        # (A) THE TOTAL, from the block the projection names as the total
        # percentile block, cross-checked before anything is reconciled to it.
        $total = Get-SimSummaryValue -Workbook $Workbook -Inspection $Inspection `
            -Bank $Bank -Measure $measure -RowKey $rowKey
        $totalScale = [Math]::Abs($derived)
        if ($total -is [double]) { $totalScale = $totalScale + [Math]::Abs([double]$total) }
        $totalAllowance = Get-W6IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $totalScale
        $totalDelta = [double]::PositiveInfinity
        if ($total -is [double]) { $totalDelta = [Math]::Abs([double]$total - $derived) }
        $null = Add-W6Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' TOTAL equals the contract' + [char]39 +
                             's Type-7 value over the iteration column') `
            (($total -is [double]) -and ($totalDelta -le $totalAllowance)) `
            ('published ' + (Format-SimValue $total) + ', derived ' + [string]$derived +
             ', delta ' + [string]$totalDelta + ', allowance ' + [string]$totalAllowance)

        # (B) sum_y Profile_Px(y) = reported Px TOTAL.
        $sum = [double]$ProfileSums[$measure]
        $scale = [double]$ProfileScale[$measure]
        if ($total -is [double]) { $scale = $scale + [Math]::Abs([double]$total) }
        $identityAllowance = Get-W6IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $scale
        $delta = [double]::PositiveInfinity
        if ($total -is [double]) { $delta = [Math]::Abs($sum - [double]$total) }
        $null = Add-W6Check ($Stage + ': the ' + $measure +
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
        $contingencyAllowance = Get-W6IdentityAllowance -Provenance $Provenance `
            -ConditioningScale $contingencyScale
        $null = Add-W6Check ($Stage + ': the published ' + $measure + ' ' + $Label +
                             ' contingency is ' + [string]$semantics.contingency_formula) `
            (($contingency -is [double]) -and ($baseline -is [double]) -and
             ($total -is [double]) -and ($contingencyDelta -le $contingencyAllowance)) `
            ('published ' + (Format-SimValue $contingency) + ', total ' +
             (Format-SimValue $total) + ' - base ' + (Format-SimValue $baseline) + ' = ' +
             [string]$expectedContingency + ', delta ' + [string]$contingencyDelta +
             ', allowance ' + [string]$contingencyAllowance)

        Write-W6Line ('    ' + $Stage.PadRight(10) + $measure.PadRight(8) +
                      ' profile sum ' + [string]$sum + '  total ' + (Format-SimValue $total) +
                      '  contingency ' + (Format-SimValue $contingency))
    }
}

# The per-year profile sums and their conditioning scale, read back from the
# persisted records. The blend itself is never recomputed here.
function Get-W6ProfileSums {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount, $Measures)
    $sums = @{}
    $scale = @{}
    $problems = New-Object System.Collections.ArrayList
    foreach ($measure in @($Measures)) { $sums[$measure] = 0.0; $scale[$measure] = 0.0 }
    for ($offset = 0; $offset -lt $YearCount; $offset++) {
        $record = Get-W6AnnualRecord -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
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
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W6 (the reporting selector moves)' -ForegroundColor Cyan
Write-Host '===============================================' -ForegroundColor Cyan
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

$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no W4 behavioural fixture.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-W6SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W6 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w6-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W6Path = Join-Path $tempRoot 'phase7_w6_selector_move.txt'
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
$firstLabel = [string]$case.selected_confidence_level
$secondLabel = [string]$case.second_confidence_level
$quantileLabels = @($gateBCases.vocabulary.quantile_labels | ForEach-Object { [string]$_ })
$firstIndex = [array]::IndexOf($quantileLabels, $firstLabel)
$secondIndex = [array]::IndexOf($quantileLabels, $secondLabel)
$firstProbability = 0.0
if ($firstLabel -match '^P(\d+)$') { $firstProbability = [double]$Matches[1] / 100.0 }
$secondProbability = 0.0
if ($secondLabel -match '^P(\d+)$') { $secondProbability = [double]$Matches[1] / 100.0 }
$selectorSemantics = $p7.selector_semantics
$measures = @($p7.summary_semantics.contingency_measures | ForEach-Object { [string]$_ })
$otherPxState = [string]$p7.handoff.profile_states[2]
$currentDistributionState = [string]$p7.handoff.distribution_states[1]
$currentProfileState = [string]$p7.handoff.profile_states[1]

Write-W6Line 'PCCM - PHASE 7 W6: THE REPORTING SELECTOR MOVES'
Write-W6Line '==============================================='
Write-W6Line ''
Write-W6Line 'This is the MINIMAL W6 runner. One simulation, one session, two selector'
Write-W6Line 'positions. No duration shrink (W7), no bank cycling (W7), no stale or'
Write-W6Line 'invalid refusal (W8), no sensitivity and no Phase-8 presentation.'
Write-W6Line ''
Write-W6Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W6Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W6Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W6Line ('git HEAD               : ' + [string]$revision.Head)
Write-W6Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W6Line ('model version          : ' + [string]$manifest.model_version)
Write-W6Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W6Line ('build directory        : ' + $BuildDir)
Write-W6Line ('working copy           : ' + $tempRoot)
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
    Write-W6Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W6Line ''
Write-W6Line 'THE FIXTURE AND THE RULE UNDER TEST'
Write-W6Line '-----------------------------------'
Write-W6Line ('  drivers              : ' + [string]$driverCount)
Write-W6Line ('  project years        : ' + [string]$yearCount)
Write-W6Line ('  seed mode            : ' + [string]$case.seed_mode +
              ', supplied seed ' + [string]$case.supplied_seed)
Write-W6Line ('  iterations           : ' + [string]$iterations)
Write-W6Line ('  selector             : ' + $firstLabel + ' -> ' + $secondLabel)
Write-W6Line ('  identity rule        : ' + [string]$cases.provenance.identity_rule)
Write-W6Line ''
Write-W6Line 'THE CONTRACT RULE THIS SCENARIO TESTS, projected rather than encoded:'
Write-W6Line ('  distribution currentness is selector-specific : ' +
              [string]$selectorSemantics.distribution_currentness_is_selector_specific)
Write-W6Line ('  profile currentness is selector-specific      : ' +
              [string]$selectorSemantics.profile_currentness_is_selector_specific)
Write-W6Line ('  profile relabelled on selector change         : ' +
              [string]$selectorSemantics.profile_relabelled_on_selector_change)
Write-W6Line ('  selector change requires a new simulation     : ' +
              [string]$selectorSemantics.selector_change_requires_new_simulation)
Write-W6Line ''
Write-W6Line 'THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE, and none is claimed. The'
Write-W6Line 'annual computation is proved independently on Linux; what only a Windows'
Write-W6Line 'run can show is that moving the selector leaves the run and the'
Write-W6Line 'distributions alone and moves the profile, and that the new profile'
Write-W6Line 'reconciles to the new total.'
Write-W6Line ''

$null = Add-W6Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W6Line ''
    Write-W6Line 'STOP. Excel was never started for the W6 session, and nothing was accepted.'
    Write-W6Line ('report                 : ' + $script:W6Path)
    Write-W6Line 'W6: FAIL'
    exit 1
}

$null = Add-W6Check 'the W6 fixture is the accepted W5 behavioural baseline' `
    (($driverCount -eq 5) -and ($yearCount -eq 4) -and ($iterations -eq 1000) -and
     (([string]$case.seed_mode) -ceq 'FIXED') -and ($suppliedSeed -eq 20260905)) `
    ([string]$driverCount + ' drivers over ' + [string]$yearCount + ' years, ' +
     [string]$iterations + ' iterations, ' + [string]$case.seed_mode + ' seed ' +
     [string]$case.supplied_seed) 'PREREQUISITE'
$null = Add-W6Check 'the two selector positions are distinct projected ladder rungs' `
    (($firstIndex -ge 0) -and ($secondIndex -ge 0) -and ($firstIndex -ne $secondIndex) -and
     ($firstProbability -gt 0) -and ($secondProbability -gt 0) -and
     ($firstProbability -ne $secondProbability)) `
    ($firstLabel + ' at rung ' + [string]($firstIndex + 1) + ' (' + [string]$firstProbability +
     ') -> ' + $secondLabel + ' at rung ' + [string]($secondIndex + 1) + ' (' +
     [string]$secondProbability + ')') 'PREREQUISITE'
# THE RULE THIS SCENARIO EXISTS FOR, asserted as a precondition. If the contract
# ever said the two products shared a currentness rule, W6 would be proving
# something else and must say so rather than quietly testing the wrong thing.
$null = Add-W6Check 'the contract still says the two annual products differ on selector currentness' `
    ((-not [bool]$selectorSemantics.distribution_currentness_is_selector_specific) -and
     ([bool]$selectorSemantics.profile_currentness_is_selector_specific) -and
     (-not [bool]$selectorSemantics.profile_relabelled_on_selector_change) -and
     (-not [bool]$selectorSemantics.selector_change_requires_new_simulation)) `
    ('distribution ' + [string]$selectorSemantics.distribution_currentness_is_selector_specific +
     ', profile ' + [string]$selectorSemantics.profile_currentness_is_selector_specific +
     ', relabelled ' + [string]$selectorSemantics.profile_relabelled_on_selector_change +
     ', needs rerun ' + [string]$selectorSemantics.selector_change_requires_new_simulation) `
    'PREREQUISITE'
$null = Add-W6Check 'the projected profile vocabulary carries an OTHER-Px state' `
    ((-not [string]::IsNullOrWhiteSpace($otherPxState)) -and
     ($otherPxState -cne $currentProfileState) -and
     (@($p7.handoff.profile_states).Count -eq 4)) `
    ((@($p7.handoff.profile_states)) -join ', ') 'PREREQUISITE'

# ===========================================================================
# THE W6 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 7 W6 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W6Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W6Line 'ever terminated, and the workbook is never saved.'
    Write-W6Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    Write-W6Line 'PREREQUISITES - THE P80 BASELINE'
    Write-W6Line '--------------------------------'

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-W6Check 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }

    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

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
        Set-W6NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.($selectorSemantics.selector_input_key).defined_name) `
            -Value $firstLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-W6Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the W6 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'W6 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-W6Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }
    $null = Add-W6Check 'the calculation reports CURRENT' `
        (([string]$excel.Run('PCCM_CalculationStatus')) -ceq 'CURRENT') '' 'PREREQUISITE'

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-W6Check 'PCCM_RunSimulation succeeded' `
        (Test-Phase6Announced -Result $simResult -Kind 'OK') $simResult 'PREREQUISITE'
    if (-not (Test-Phase6Announced -Result $simResult -Kind 'OK')) { throw $simResult }
    $null = Add-W6Check 'the simulation reports CURRENT' `
        ((([string]$excel.Run('PCCM_SimulationStatus')) -ceq 'CURRENT')) '' 'PREREQUISITE'

    $state = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bank = Get-Phase6ActiveBank -State $state
    $null = Add-W6Check 'a publication bank is active' `
        ((-not [string]::IsNullOrEmpty($bank)) -and
         (@($simInspection.publication.bank_labels) -contains $bank)) $bank 'PREREQUISITE'
    if ([string]::IsNullOrEmpty($bank)) { throw 'no publication bank was activated' }

    $annualFirst = Invoke-W6Annual -Excel $excel -P7 $p7
    $null = Add-W6Check ('PCCM_RunAnnualStochastic succeeded at ' + $firstLabel) `
        ($annualFirst -like 'OK|*') $annualFirst 'PREREQUISITE'
    if (-not ($annualFirst -like 'OK|*')) { throw $annualFirst }

    # -------------------------------------------------------------------
    # THE BASELINE, CAPTURED WHOLE
    # -------------------------------------------------------------------
    $simBefore = Get-W6RunInvariants -Workbook $wb -Inspection $simInspection
    $stateBefore = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $simBlock = $stateBefore[('bank_' + $bank)]
    $iterationsBefore = Get-W6IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsBefore.Acquired
    $annualBefore = Get-W6AnnualSurface -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $handoffBefore = Get-W6Handoff -Excel $excel -P7 $p7
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })

    $null = Add-W6Check ('the baseline profile is stamped ' + $firstLabel) `
        ((Test-SimExactText -Actual $annualBefore[('stamp.selected_px_label')] -Expected $firstLabel) -and
         (Test-SimExactDouble -Actual $annualBefore[('stamp.selected_px_probability')] `
             -Expected $firstProbability)) `
        ((Format-SimValue $annualBefore['stamp.selected_px_label']) + ' / ' +
         (Format-SimValue $annualBefore['stamp.selected_px_probability'])) 'PREREQUISITE'
    $null = Add-W6Check 'the baseline handoff reports both products CURRENT' `
        (((([string]$handoffBefore[$accessors[0]]) -ceq $currentDistributionState)) -and
         ((([string]$handoffBefore[$accessors[1]]) -ceq $currentProfileState)) -and
         (([int]$handoffBefore[$accessors[3]]) -eq $yearCount)) `
        ((Format-SimValue $handoffBefore[$accessors[0]]) + ' / ' +
         (Format-SimValue $handoffBefore[$accessors[1]]) + ' / ' +
         (Format-SimValue $handoffBefore[$accessors[2]]) + ' / ' +
         (Format-SimValue $handoffBefore[$accessors[3]])) 'PREREQUISITE'

    $profileFirst = Get-W6ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -Measures $measures
    $null = Add-W6Check ('every baseline ' + $firstLabel + ' profile value is a number') `
        (@($profileFirst.Problems).Count -eq 0) ((@($profileFirst.Problems)) -join '; ') 'PREREQUISITE'
    Write-W6Line ''
    Invoke-W6Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsBefore.Values -Bank $bank `
        -Label $firstLabel -Probability $firstProbability -LadderIndex $firstIndex `
        -ProfileSums $profileFirst.Sums -ProfileScale $profileFirst.Scale -Stage 'baseline'

    Write-W6Line ''
    Write-W6Line (Format-Phase6State -State $stateBefore -Label '  the run identity BEFORE the selector move')

    # ===================================================================
    # PART A - THE SELECTOR MOVES, AND NOTHING ELSE RUNS
    # ===================================================================
    Write-W6Line ''
    Write-W6Line ('PART A - THE SELECTOR MOVES ' + $firstLabel + ' -> ' + $secondLabel)
    Write-W6Line '------------------------------------------'
    # THE ORDINARY USER-FACING INPUT, through the projection that names it. No
    # machine cell is written, no endpoint is called, and nothing is rerun.
    Set-W6NamedText -Workbook $wb `
        -DefinedName ([string]$inspection.inputs.($selectorSemantics.selector_input_key).defined_name) `
        -Value $secondLabel

    $simAfterMove = Get-W6RunInvariants -Workbook $wb -Inspection $simInspection
    $null = Add-W6InvariantChecks 'the selector move' $simBefore $simAfterMove
    $stateAfterMove = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $null = Add-W6Check 'the simulation is still CURRENT after the selector move' `
        ((([string]$excel.Run('PCCM_SimulationStatus')) -ceq 'CURRENT')) ''
    $null = Add-W6Check 'the active bank did not move' `
        ((Get-Phase6ActiveBank -State $stateAfterMove) -ceq $bank) `
        ((Get-Phase6ActiveBank -State $stateAfterMove) + ' vs ' + $bank)
    # THE FINGERPRINT SEMANTIC, NAMED AND PROVED rather than inferred from a
    # CURRENT: the selector does not enter the simulation request identity, so
    # the fingerprint and the digest of the successful run may not move.
    foreach ($field in @('request_fingerprint', 'result_digest', 'run_id', 'effective_seed',
                         'supplied_seed', 'iterations_run', 'consumed_auto_nonce')) {
        $null = Add-W6Check ('the selector move left the successful run' + [char]39 + 's ' +
                             $field + ' unchanged') `
            (Test-SimSameValue -A $simBlock[$field] -B $stateAfterMove[('bank_' + $bank)][$field]) `
            ((Format-SimValue $simBlock[$field]) + ' -> ' +
             (Format-SimValue $stateAfterMove[('bank_' + $bank)][$field]))
    }
    $iterationsAfterMove = Get-W6IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsAfterMove.Acquired
    $movedIterations = Compare-W6IterationGrid -Before $iterationsBefore.Values `
        -After $iterationsAfterMove.Values
    $null = Add-W6Check 'the selector move changed no published iteration value' `
        ([string]::IsNullOrWhiteSpace($movedIterations)) $movedIterations

    # THE DISTRIBUTIONS ARE A PROPERTY OF THE RUN ALONE.
    $annualAfterMove = Get-W6AnnualSurface -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $wholeMove = Compare-W6Surface -Before $annualBefore -After $annualAfterMove
    $null = Add-W6Check 'the selector move rewrote nothing in the persisted annual answer' `
        ($wholeMove.Moved -eq 0) $wholeMove.Detail
    $ladderMove = Compare-W6Surface -Before $annualBefore -After $annualAfterMove -Only '.ladder_'
    $null = Add-W6Check 'every persisted distribution rung is value-identical after the move' `
        ($ladderMove.Moved -eq 0) $ladderMove.Detail
    $stampMove = Compare-W6Surface -Before $annualBefore -After $annualAfterMove -Only 'stamp.'
    $null = Add-W6Check 'the annual stamp is still bound to the same successful simulation' `
        ($stampMove.Moved -eq 0) $stampMove.Detail

    # AND THE HANDOFF SPLITS: distribution CURRENT, profile OTHER Px.
    $handoffAfterMove = Get-W6Handoff -Excel $excel -P7 $p7
    $null = Add-W6Check ($accessors[0] + ' is still CURRENT after the selector move') `
        ((([string]$handoffAfterMove[$accessors[0]]) -ceq $currentDistributionState)) `
        (Format-SimValue $handoffAfterMove[$accessors[0]])
    $null = Add-W6Check ($accessors[1] + ' reports the projected OTHER-Px state') `
        ((([string]$handoffAfterMove[$accessors[1]]) -ceq $otherPxState)) `
        ((Format-SimValue $handoffAfterMove[$accessors[1]]) + ' vs ' + $otherPxState)
    $null = Add-W6Check ($accessors[2] + ' still identifies the OLD ' + $firstLabel + ' profile') `
        ((([string]$handoffAfterMove[$accessors[2]]) -match ([regex]::Escape($firstLabel)))) `
        (Format-SimValue $handoffAfterMove[$accessors[2]])
    $null = Add-W6Check ($accessors[3] + ' still reports the project year count') `
        (([int]$handoffAfterMove[$accessors[3]]) -eq $yearCount) `
        (Format-SimValue $handoffAfterMove[$accessors[3]])
    # THE OLD PROFILE IS NEITHER CLEARED NOR RELABELLED.
    $null = Add-W6Check ('the old profile is still stamped ' + $firstLabel +
                         ' and was not relabelled') `
        ((Test-SimExactText -Actual $annualAfterMove['stamp.selected_px_label'] `
             -Expected $firstLabel) -and
         (Test-SimExactDouble -Actual $annualAfterMove['stamp.selected_px_probability'] `
             -Expected $firstProbability)) `
        ((Format-SimValue $annualAfterMove['stamp.selected_px_label']) + ' / ' +
         (Format-SimValue $annualAfterMove['stamp.selected_px_probability']))

    # ===================================================================
    # PART B - THE ANNUAL ENDPOINT ALONE RERUNS
    # ===================================================================
    Write-W6Line ''
    Write-W6Line ('PART B - ANNUAL POST-PROCESSING RERUNS AT ' + $secondLabel)
    Write-W6Line '-------------------------------------------------'
    $annualSecond = Invoke-W6Annual -Excel $excel -P7 $p7
    $null = Add-W6Check ('PCCM_RunAnnualStochastic succeeded at ' + $secondLabel) `
        ($annualSecond -like 'OK|*') $annualSecond
    if (-not ($annualSecond -like 'OK|*')) { throw $annualSecond }

    $simAfterRerun = Get-W6RunInvariants -Workbook $wb -Inspection $simInspection
    $null = Add-W6InvariantChecks 'the annual rerun' $simBefore $simAfterRerun
    $stateAfterRerun = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $null = Add-W6Check 'the annual rerun allocated no second run identity' `
        ((Test-SimSameValue -A $simBlock['run_id'] `
             -B $stateAfterRerun[('bank_' + $bank)]['run_id']) -and
         (Test-SimSameValue -A $stateBefore['shared']['last_run_id'] `
             -B $stateAfterRerun['shared']['last_run_id'])) `
        ((Format-SimValue $simBlock['run_id']) + ' -> ' +
         (Format-SimValue $stateAfterRerun[('bank_' + $bank)]['run_id']))
    $null = Add-W6Check 'the annual rerun opened no second publication bank' `
        ((Get-Phase6ActiveBank -State $stateAfterRerun) -ceq $bank) `
        ((Get-Phase6ActiveBank -State $stateAfterRerun) + ' vs ' + $bank)
    foreach ($field in @('request_fingerprint', 'result_digest', 'effective_seed',
                         'supplied_seed', 'iterations_run', 'consumed_auto_nonce')) {
        $null = Add-W6Check ('the annual rerun left the successful run' + [char]39 + 's ' +
                             $field + ' unchanged') `
            (Test-SimSameValue -A $simBlock[$field] -B $stateAfterRerun[('bank_' + $bank)][$field]) `
            ((Format-SimValue $simBlock[$field]) + ' -> ' +
             (Format-SimValue $stateAfterRerun[('bank_' + $bank)][$field]))
    }
    $null = Add-W6Check 'the annual rerun consumed no AUTO nonce and moved no pending marker' `
        ((Test-SimSameValue -A $stateBefore['shared']['next_auto_nonce'] `
             -B $stateAfterRerun['shared']['next_auto_nonce']) -and
         (Test-SimSameValue -A $stateBefore['pending_auto_nonce'] `
             -B $stateAfterRerun['pending_auto_nonce'])) `
        ((Format-SimValue $stateBefore['shared']['next_auto_nonce']) + ' / ' +
         (Format-SimValue $stateBefore['pending_auto_nonce']))
    $iterationsAfterRerun = Get-W6IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsAfterRerun.Acquired
    $movedAgain = Compare-W6IterationGrid -Before $iterationsBefore.Values `
        -After $iterationsAfterRerun.Values
    $null = Add-W6Check 'the annual rerun changed no published iteration value' `
        ([string]::IsNullOrWhiteSpace($movedAgain)) $movedAgain

    # THE DISTRIBUTIONS DID NOT MOVE. A reporting selector may not recompute a
    # distribution, and this is where that would show.
    $annualAfterRerun = Get-W6AnnualSurface -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount
    $ladderRerun = Compare-W6Surface -Before $annualBefore -After $annualAfterRerun -Only '.ladder_'
    $null = Add-W6Check ('every distribution rung is value-identical to the ' + $firstLabel +
                         ' baseline after the rerun') `
        ($ladderRerun.Moved -eq 0) $ladderRerun.Detail
    $identityRerun = Compare-W6Surface -Before $annualBefore -After $annualAfterRerun `
        -Only '.project_index'
    $null = Add-W6Check 'the project-year identities are unchanged' `
        ($identityRerun.Moved -eq 0) $identityRerun.Detail
    $calendarRerun = Compare-W6Surface -Before $annualBefore -After $annualAfterRerun `
        -Only '.calendar_year'
    $null = Add-W6Check 'the calendar years are unchanged' `
        ($calendarRerun.Moved -eq 0) $calendarRerun.Detail
    foreach ($field in @('run_id', 'effective_seed', 'request_fingerprint', 'result_digest',
                         'iterations', 'year_count')) {
        $null = Add-W6Check ('the annual stamp' + [char]39 + 's ' + $field +
                             ' is still the baseline value') `
            (Test-SimSameValue -A $annualBefore[('stamp.' + $field)] `
                -B $annualAfterRerun[('stamp.' + $field)]) `
            ((Format-SimValue $annualBefore[('stamp.' + $field)]) + ' -> ' +
             (Format-SimValue $annualAfterRerun[('stamp.' + $field)]))
    }

    # THE PROFILE MOVED, AND ITS STAMP MOVED WITH IT.
    $null = Add-W6Check ('the profile is now stamped ' + $secondLabel + ' at its projected probability') `
        ((Test-SimExactText -Actual $annualAfterRerun['stamp.selected_px_label'] `
             -Expected $secondLabel) -and
         (Test-SimExactDouble -Actual $annualAfterRerun['stamp.selected_px_probability'] `
             -Expected $secondProbability)) `
        ((Format-SimValue $annualAfterRerun['stamp.selected_px_label']) + ' / ' +
         (Format-SimValue $annualAfterRerun['stamp.selected_px_probability']))
    # REPORTED, NOT REQUIRED. The P50 profile is expected to differ from the P80
    # one, but a fixture COULD produce equal values and failing on equality
    # would assert a property of this fixture rather than of the code. The
    # reconciliation below is what proves the new profile correct.
    $profileMoved = Compare-W6Surface -Before $annualBefore -After $annualAfterRerun -Only '.profile_'
    Write-W6Line ('    the selected-Px profile changed in ' + [string]$profileMoved.Moved +
                  ' of ' + [string]($yearCount * @($measures).Count) +
                  ' published values (reported, not required)')

    $handoffAfterRerun = Get-W6Handoff -Excel $excel -P7 $p7
    $null = Add-W6Check ($accessors[0] + ' is CURRENT after the rerun') `
        ((([string]$handoffAfterRerun[$accessors[0]]) -ceq $currentDistributionState)) `
        (Format-SimValue $handoffAfterRerun[$accessors[0]])
    $null = Add-W6Check ($accessors[1] + ' is CURRENT again after the rerun') `
        ((([string]$handoffAfterRerun[$accessors[1]]) -ceq $currentProfileState)) `
        (Format-SimValue $handoffAfterRerun[$accessors[1]])
    $null = Add-W6Check ($accessors[2] + ' now identifies ' + $secondLabel) `
        ((([string]$handoffAfterRerun[$accessors[2]]) -match ([regex]::Escape($secondLabel)))) `
        (Format-SimValue $handoffAfterRerun[$accessors[2]])
    $null = Add-W6Check ($accessors[3] + ' still reports the project year count') `
        (([int]$handoffAfterRerun[$accessors[3]]) -eq $yearCount) `
        (Format-SimValue $handoffAfterRerun[$accessors[3]])

    # THE NEW RUNG, RECONCILED BY THE SAME CODE THAT RECONCILED THE OLD ONE.
    $profileSecond = Get-W6ProfileSums -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -YearCount $yearCount -Measures $measures
    $null = Add-W6Check ('every new ' + $secondLabel + ' profile value is a number') `
        (@($profileSecond.Problems).Count -eq 0) ((@($profileSecond.Problems)) -join '; ')
    Write-W6Line ''
    Invoke-W6Reconciliation -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Provenance $cases.provenance -Grid $iterationsBefore.Values -Bank $bank `
        -Label $secondLabel -Probability $secondProbability -LadderIndex $secondIndex `
        -ProfileSums $profileSecond.Sums -ProfileScale $profileSecond.Scale -Stage 'rerun'

    # THE REGION IS STILL COMPACT.
    $region = Get-W6AnnualRegionIndex -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -Ledger $rel
    $comAcquired = $comAcquired + [int]$region.Acquired
    $regionValues = $region.Values
    $populated = 0
    $regionRows = 0
    if ($null -ne $regionValues) { $regionRows = [int]$regionValues.GetLength(0) }
    for ($row = 1; $row -le $regionRows; $row++) {
        if (-not (Test-SimBlank -Value $regionValues[$row, 1])) { $populated = $populated + 1 }
    }
    $null = Add-W6Check ('the annual region still holds exactly ' + [string]$yearCount + ' records') `
        ($populated -eq $yearCount) ('populated ' + [string]$populated + ' of ' +
                                     [string]$regionRows + ' available rows')
    $boundaryRow = $yearCount + 1
    $boundaryValue = $null
    if ($boundaryRow -le $regionRows) { $boundaryValue = $regionValues[$boundaryRow, 1] }
    $null = Add-W6Check 'the row immediately after the annual result is still blank' `
        (Test-SimBlank -Value $boundaryValue) `
        ('row ' + [string]$boundaryRow + ' = ' + (Format-SimValue $boundaryValue))

    Write-W6Line ''
    Write-W6Line (Format-Phase6State -State $stateAfterRerun -Label '  the run identity AFTER both moves')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W6Line ''
    Write-W6Line ('THE W6 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W6Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W6Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W6Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W6Line ''
        Write-W6Line 'EXCEL SHUTDOWN'
        Write-W6Line '--------------'
        if ($naturalExit) {
            Write-W6Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W6'
            Write-W6Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W6Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W6Residual)) {
            Write-W6Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W6Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W6Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W6Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W6Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W6Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W6Check 'every COM release left 0 outstanding references' `
    (@($script:W6Residual).Count -eq 0) ((@($script:W6Residual)) -join '; ')
$null = Add-W6Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W6Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W6Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W6Line ''
Write-W6Line 'VERDICT'
Write-W6Line '-------'
Write-W6Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W6Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W6Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W6Line ''
if ($ok) {
    Write-W6Line 'W6: PASS'
} else {
    Write-W6Line 'W6: FAIL'
    Write-W6Line ''
    Write-W6Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W6Line ''
Write-W6Line ('report                 : ' + $script:W6Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W6Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
