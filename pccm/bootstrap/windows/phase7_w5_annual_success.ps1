<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W5 runner: the first successful annual run, its
    persistence, and its reconciliation.

.DESCRIPTION
    W5 IS THE FIRST TIME `PCCM_RunAnnualStochastic` SUCCEEDS IN REAL EXCEL. W5
    proved it refuses with no simulation and that a FIXED-seed baseline
    publishes cleanly; W5 rebuilds that baseline in its own disposable session
    and then asks the annual endpoint to produce an answer, persist it, stamp it
    to the run that produced it, and hand it to Phase 8 correctly.

    WHAT IS INDEPENDENT HERE, STATED HONESTLY. There is NO independent annual
    Windows oracle, and W5 does not pretend there is one. The independent proof
    of the annual COMPUTATION is the Linux P7-5/P7-6 work, where the Python
    implementation and the VBA are proved against each other statically. W5's
    job is different and cannot be done there: live Excel/VBA EXECUTION,
    PERSISTENCE, READ-BACK and RECONCILIATION. What it can lean on is real:

      * the accepted Phase-5 oracle for the deterministic base underneath;
      * the published simulation iteration columns, which W5 established as the
        authoritative run data;
      * the contract's own Type-7 definition, applied to those columns, to
        derive the total Px the profile must reconcile to;
      * the persisted annual result, read back through its contracted surface.

    THE PROFILE IS NEVER RECOMPUTED HERE. The convex Type-7 BLEND that produces
    the selected-Px annual profile is production's, and reimplementing it in
    PowerShell would compare an implementation against a copy of itself. What is
    computed here is the TOTAL percentile - one Type-7 value over the published
    iteration totals, by the contract's own formula - and the identity that must
    then hold is the one sim_contract names: sum_y Profile_Px(y) = reported Px.

    THE ALLOWANCE IS THE PROJECT'S OWN IDENTITY RULE, projected into the corpus
    rather than chosen here: |delta| <= max(floor, coefficient * scale), with the
    conditioning scale naming the magnitude of the arithmetic performed rather
    than the magnitude of its net result - the ERRATUM C1 correction, applied to
    the terms this identity actually sums. The deltas are reported either way,
    and nothing is ever scaled after the fact to make them fit.

    WHAT IT PROVES
    --------------
      1. the candidate identity and a Stage-B workbook from the current build
      2. the W5 baseline, rebuilt and re-proved: calculation CURRENT, simulation
         CURRENT, FIXED seed 20260905, 1,000 iterations, a known active bank
      3. PCCM_RunAnnualStochastic SUCCEEDS
      4. the annual stamp binds to that exact run - every projected stamp field
      5. one indexed annual record per project year, both percentile ladders
         complete and numeric, and ordered where the contract requires it
      6. the selected-Px profiles published for every year, in both measures
      7. sum_y Profile_Px(y) reconciles to the reported total Px, nominal and PV
      8. the four Phase-8 accessors report CURRENT, the selected identity and 4
      9. the simulation identity AND the iteration columns are value-identical
         before and after: annual is observational post-processing, not a second
         stochastic run
     10. the persisted shape is COMPACT: one record per project year and nothing
         beyond, so no N x Y matrix and no iteration-level workspace was left
         behind as a published record set

    W5 STOPS THERE. The reporting selector is not moved - that is W6 - the
    duration is not shrunk, and no stale or invalid refusal is provoked.

    THE SHAPE IS W2's, W3's AND W5's, and the functions the runners share are
    pinned to W2's so the mandated reuse cannot drift.

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
$script:W5Lines = New-Object System.Collections.ArrayList
$script:W5Path = ''
$script:W5Checks = New-Object System.Collections.ArrayList

function Write-W5Line {
    param([string]$Text = '')
    $null = $script:W5Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W5Path)) {
        try {
            Set-Content -LiteralPath $script:W5Path `
                -Value ($script:W5Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W5 had to perform to reach
# its subject - from a RESULT, which is W5's own claim about the calculation. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W5Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W5Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W5Line $line
    return $Ok
}

function Format-W5Value {
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
$script:W5Residual = New-Object System.Collections.ArrayList

function Invoke-W5Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W5Residual.Add(
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
function Get-W5SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W5 result could not be attributed to a source revision')
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
function Compare-W5Cell {
    param($Got, $Expect, [double]$Allowance)
    if ($null -eq $Expect) {
        if (($null -eq $Got) -or ($Got -is [System.DBNull]) -or
            (($Got -is [string]) -and ([string]$Got).Length -eq 0)) { return '' }
        return ('published ' + (Format-W5Value $Got) + ' where the oracle has none')
    }
    if ($Expect -is [string]) {
        if (($Got -is [string]) -and (([string]$Got) -ceq ([string]$Expect))) { return '' }
        return ((Format-W5Value $Got) + ' vs ' + [string]$Expect)
    }
    if ($Got -isnot [double]) {
        return ((Format-W5Value $Got) + ' is not a number')
    }
    $difference = [Math]::Abs([double]$Got - [double]$Expect)
    if ($difference -le $Allowance) { return '' }
    return ([string]$Got + ' vs ' + [string]$Expect + ' (difference ' + [string]$difference + ')')
}

# THE PUBLISHED TABLE, ROW FOR ROW, EVERY PROJECTED COLUMN. A failing table
# reports the first twelve disagreements and the count: the count says how bad
# it is, the examples make it diagnosable, and 6,300 lines would do neither.
function Compare-W5Table {
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
            $problem = Compare-W5Cell -Got $Live[$index][$at] -Expect $expect -Allowance $Allowance
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
    return (Add-W5Check ($Label + ': ' + $TableKey +
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
function Get-W5AnnualStamp {
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
function Get-W5AnnualFirstRecord {
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

function Get-W5Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $out.Add([string]$accessor, $Excel.Run([string]$accessor))
    }
    return $out
}

# THE ANNUAL ATTEMPT, THROUGH THE PUBLISHED ENDPOINT AND ITS PUBLISHED ANSWER.
# The runner never judges the outcome itself: it asks production what happened.
function Invoke-W5Annual {
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
$script:W5DerivedRows = @('simulation_status', 'status_evaluated_at')

function Get-W5RunInvariants {
    param($Workbook, $Inspection)
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($script:W5DerivedRows -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-W5InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ([string]$key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-W5Check ($Label + ': no run identity, nonce, pending marker or digest moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# A defined name that carries TEXT. Set-NamedValue writes a Double - PowerShell
# binds a COM property call site per argument type, so one polymorphic
# assignment silently fails the second time it is reached with another type -
# and the selected confidence level is a label, not a number.
function Set-W5NamedText {
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
function Get-W5IterationBlock {
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
        if ($null -ne $range)  { Invoke-W5Release $Ledger $range  'Range(iterations)';   $range  = $null }
        if ($null -ne $sheet)  { Invoke-W5Release $Ledger $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W5Release $Ledger $sheets 'Worksheets';          $sheets = $null }
    }
}


# ===========================================================================
# COLUMN ARITHMETIC, AND THE ANNUAL LADDER IT REACHES
# ===========================================================================
# The projection gives the FIRST ladder column per bank per measure and the
# ladder's LENGTH; the eleven columns follow it. That is the contract's shape,
# so the offset is computed rather than eleven letters being typed in.
# THE DEFECT THAT STOPPED THE FIRST W5 RUN, and the reason it survived every
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
function ConvertTo-W5ColumnNumber {
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
function ConvertFrom-W5ColumnNumber {
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
function Get-W5AnnualRecord {
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
        $first = ConvertTo-W5ColumnNumber -Letters ([string]$records.quantile_first_column.$Bank.$measure)
        $ladder = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt $count; $index++) {
            $column = ConvertFrom-W5ColumnNumber -Number ($first + $index)
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
function Get-W5AnnualRegionIndex {
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
        if ($null -ne $range)  { Invoke-W5Release $Ledger $range  'Range(annual region)'; $range  = $null }
        if ($null -ne $sheet)  { Invoke-W5Release $Ledger $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Invoke-W5Release $Ledger $sheets 'Worksheets';           $sheets = $null }
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
function Get-W5Type7Value {
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
function Get-W5IdentityAllowance {
    param($Provenance, [double]$ConditioningScale)
    $floor = [double]$Provenance.identity_absolute_floor
    $coefficient = [double]$Provenance.identity_relative_coefficient
    $scaleFloor = [double]$Provenance.conditioning_scale_floor
    $scale = [Math]::Max($scaleFloor, [Math]::Abs($ConditioningScale))
    return [Math]::Max($floor, $coefficient * $scale)
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W5 (the first successful annual run)' -ForegroundColor Cyan
Write-Host '==================================================' -ForegroundColor Cyan
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
try { $revision = Get-W5SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W5 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w5-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $gateBCasePath,
                        $p7InspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W5Path = Join-Path $tempRoot 'phase7_w5_annual_success.txt'
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
$selectedLabel = [string]$case.selected_confidence_level
$quantileLabels = @($gateBCases.vocabulary.quantile_labels | ForEach-Object { [string]$_ })
$selectedIndex = [array]::IndexOf($quantileLabels, $selectedLabel)
# THE PROBABILITY A Pxx LABEL NAMES. It is the only reading of the label, and it
# is checked against the workbook's own resolved value below rather than trusted.
$selectedProbability = 0.0
if ($selectedLabel -match '^P(\d+)$') { $selectedProbability = [double]$Matches[1] / 100.0 }

Write-W5Line 'PCCM - PHASE 7 W5: THE FIRST SUCCESSFUL ANNUAL RUN'
Write-W5Line '=================================================='
Write-W5Line ''
Write-W5Line 'This is the MINIMAL W5 runner. It rebuilds the accepted W4 baseline in its'
Write-W5Line 'own disposable session and then runs the annual endpoint successfully for'
Write-W5Line 'the first time. The reporting selector is NOT moved (W6), the duration is'
Write-W5Line 'NOT shrunk (W7), and no refusal is provoked (W8). No Gate-B result is'
Write-W5Line 'produced or implied.'
Write-W5Line ''
Write-W5Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W5Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W5Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W5Line ('git HEAD               : ' + [string]$revision.Head)
Write-W5Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W5Line ('model version          : ' + [string]$manifest.model_version)
Write-W5Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-W5Line ('build directory        : ' + $BuildDir)
Write-W5Line ('working copy           : ' + $tempRoot)
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
    Write-W5Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W5Line ''
Write-W5Line 'THE FIXTURE AND THE AUTHORITIES'
Write-W5Line '-------------------------------'
Write-W5Line ('  drivers              : ' + [string]$driverCount)
Write-W5Line ('  project years        : ' + [string]$yearCount)
Write-W5Line ('  seed mode            : ' + [string]$case.seed_mode +
              ', supplied seed ' + [string]$case.supplied_seed)
Write-W5Line ('  iterations           : ' + [string]$iterations)
Write-W5Line ('  selected Px          : ' + $selectedLabel + ' (ladder position ' +
              [string]($selectedIndex + 1) + ' of ' + [string]$quantileLabels.Count + ')')
Write-W5Line ('  deterministic oracle : ' + [string]$cases.provenance.expectation_source)
Write-W5Line ('  identity rule        : ' + [string]$cases.provenance.identity_rule)
Write-W5Line ''
Write-W5Line 'THERE IS NO INDEPENDENT ANNUAL WINDOWS ORACLE, and W5 does not claim one.'
Write-W5Line 'The annual COMPUTATION is proved independently on Linux, where the Python'
Write-W5Line 'implementation and the VBA are checked against each other statically. What'
Write-W5Line 'only a Windows run can prove is what W5 proves: that the endpoint executes,'
Write-W5Line 'persists a compact result, stamps it to the run that produced it, hands it'
Write-W5Line 'to Phase 8 correctly, and reconciles to the total percentile derived from'
Write-W5Line ('the published iteration column by the contract' + [char]39 +
              's own Type-7 definition.')
Write-W5Line ''

$null = Add-W5Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W5Line ''
    Write-W5Line 'STOP. Excel was never started for the W5 session, and nothing was accepted.'
    Write-W5Line ('report                 : ' + $script:W5Path)
    Write-W5Line 'W5: FAIL'
    exit 1
}

$null = Add-W5Check 'the W5 fixture is the accepted W4 baseline' `
    (($driverCount -eq 5) -and ($yearCount -eq 4) -and ($iterations -eq 1000) -and
     (([string]$case.seed_mode) -ceq 'FIXED') -and ($suppliedSeed -eq 20260905) -and
     ($selectedLabel -ceq 'P80')) `
    ([string]$driverCount + ' drivers over ' + [string]$yearCount + ' years, ' +
     [string]$iterations + ' iterations, ' + [string]$case.seed_mode + ' seed ' +
     [string]$case.supplied_seed + ', selected ' + $selectedLabel) 'PREREQUISITE'
$null = Add-W5Check 'the selected label is a position on the projected quantile ladder' `
    (($selectedIndex -ge 0) -and ($quantileLabels.Count -eq [int]$p7.annual_records.quantile_count)) `
    ([string]$quantileLabels.Count + ' projected labels, ladder position ' +
     [string]($selectedIndex + 1)) 'PREREQUISITE'
$null = Add-W5Check 'the identity rule reached this runner from the corpus' `
    (([double]$cases.provenance.identity_absolute_floor -gt 0) -and
     ([double]$cases.provenance.identity_relative_coefficient -gt 0) -and
     ([double]$cases.provenance.conditioning_scale_floor -gt 0)) `
    ('floor ' + [string]$cases.provenance.identity_absolute_floor + ', coefficient ' +
     [string]$cases.provenance.identity_relative_coefficient + ', scale floor ' +
     [string]$cases.provenance.conditioning_scale_floor) 'PREREQUISITE'

# ===========================================================================
# THE W5 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

$rel = New-ReleaseLedger 'phase 7 W5 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W5Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W5Line 'ever terminated, and the workbook is never saved.'
    Write-W5Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    Write-W5Line 'PREREQUISITES - THE W4 BASELINE, REBUILT AND RE-PROVED'
    Write-W5Line '-----------------------------------------------------'

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $compiled = [string]::IsNullOrWhiteSpace($compileFailure)
    $null = Add-W5Check 'the current VBAProject compiles in real Excel' `
        $compiled $compileFailure 'PREREQUISITE'
    if (-not $compiled) { throw ('the VBAProject does not compile: ' + $compileFailure) }

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
        Set-W5NamedText -Workbook $wb `
            -DefinedName ([string]$inspection.inputs.selected_confidence_level.defined_name) `
            -Value $selectedLabel
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-W5Check 'the behavioural model and the simulation request were applied' `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) { throw ('the W5 model was not applied: ' + $applied) }

    $calcFailure = ''
    try {
        $null = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'W5 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-W5Check 'PCCM_Calculate succeeded' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) $calcFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) { throw $calcFailure }
    $calcStatus = [string]$excel.Run('PCCM_CalculationStatus')
    $null = Add-W5Check 'the calculation reports CURRENT before the simulation' `
        ($calcStatus -ceq 'CURRENT') $calcStatus 'PREREQUISITE'

    # THE DETERMINISTIC ANCHOR, against the independent Phase-5 oracle. The
    # annual answer is built on this calculation, so a wrong base would make
    # every reconciliation below meaningless.
    $totalProblems = @()
    foreach ($key in $case.expected.totals.PSObject.Properties.Name) {
        $got = Get-CalcScalar -Workbook $wb -Inspection $inspection `
            -Block 'calc_totals' -FieldKey $key
        $problem = Compare-W5Cell -Got $got -Expect $case.expected.totals.$key -Allowance $allowance
        if (-not [string]::IsNullOrWhiteSpace($problem)) {
            $totalProblems += ([string]$key + ': ' + $problem)
        }
    }
    $null = Add-W5Check 'the deterministic totals match the independent Phase-5 oracle' `
        ($totalProblems.Count -eq 0) ($totalProblems -join '; ') 'PREREQUISITE'

    $simResult = Invoke-Phase6Simulation -Excel $excel
    $null = Add-W5Check 'PCCM_RunSimulation succeeded' `
        (Test-Phase6Announced -Result $simResult -Kind 'OK') $simResult 'PREREQUISITE'
    if (-not (Test-Phase6Announced -Result $simResult -Kind 'OK')) { throw $simResult }

    $simStatus = [string]$excel.Run('PCCM_SimulationStatus')
    $null = Add-W5Check 'the simulation reports CURRENT' `
        (($simStatus -ceq 'CURRENT') -and
         (@($gateBCases.vocabulary.sim_states) -contains $simStatus)) $simStatus 'PREREQUISITE'

    # THE IDENTITY THE ANNUAL RESULT MUST BIND TO, CAPTURED BEFORE IT RUNS.
    $stateBefore = Get-Phase6State -Workbook $wb -Inspection $simInspection
    $bank = Get-Phase6ActiveBank -State $stateBefore
    $null = Add-W5Check 'a publication bank is active' `
        ((-not [string]::IsNullOrEmpty($bank)) -and
         (@($simInspection.publication.bank_labels) -contains $bank)) $bank 'PREREQUISITE'
    if ([string]::IsNullOrEmpty($bank)) { throw 'no publication bank was activated' }
    $simBlock = $stateBefore[('bank_' + $bank)]
    $null = Add-W5Check 'the FIXED request was honoured before the annual run' `
        ((Test-SimExactDouble -Actual $simBlock['effective_seed'] -Expected $suppliedSeed) -and
         (Test-SimExactDouble -Actual $simBlock['iterations_run'] -Expected ([double]$iterations)) -and
         (Test-SimExactText -Actual $simBlock['seed_mode'] -Expected ([string]$case.seed_mode))) `
        ('seed ' + (Format-SimValue $simBlock['effective_seed']) + ', iterations ' +
         (Format-SimValue $simBlock['iterations_run']) + ', mode ' +
         (Format-SimValue $simBlock['seed_mode'])) 'PREREQUISITE'
    foreach ($field in @('run_id', 'request_fingerprint', 'result_digest')) {
        $null = Add-W5Check ('the simulation published a ' + $field + ' to bind to') `
            (-not (Test-SimBlank -Value $simBlock[$field])) (Format-SimValue $simBlock[$field]) `
            'PREREQUISITE'
    }
    $invariantsBefore = Get-W5RunInvariants -Workbook $wb -Inspection $simInspection
    $iterationsBefore = Get-W5IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsBefore.Acquired

    Write-W5Line ''
    Write-W5Line (Format-Phase6State -State $stateBefore -Label '  the run identity BEFORE the annual step')

    # -------------------------------------------------------------------
    # THE ANNUAL RUN
    # -------------------------------------------------------------------
    Write-W5Line ''
    Write-W5Line 'THE ANNUAL RUN'
    Write-W5Line '--------------'
    $announcement = Invoke-W5Annual -Excel $excel -P7 $p7
    $annualOk = Add-W5Check 'PCCM_RunAnnualStochastic succeeded' `
        ($announcement -like 'OK|*') $announcement
    if (-not $annualOk) { throw ('the annual endpoint did not succeed: ' + $announcement) }

    # -------------------------------------------------------------------
    # THE STAMP, BOUND TO THE RUN THAT PRODUCED IT
    # -------------------------------------------------------------------
    $stamp = Get-W5AnnualStamp -Workbook $wb -Inspection $simInspection -P7 $p7 -Bank $bank
    $null = Add-W5Check 'the annual result is stamped in the active simulation bank' `
        (-not (Test-SimBlank -Value $stamp['published'])) (Format-SimValue $stamp['published'])
    $null = Add-W5Check 'the publication marker is the contracted one' `
        (Test-SimExactText -Actual $stamp['published'] `
            -Expected ([string]$p7.annual_records.stamp.published_marker)) `
        ((Format-SimValue $stamp['published']) + ' vs ' +
         [string]$p7.annual_records.stamp.published_marker)

    # EVERY STAMPED IDENTITY FIELD, AGAINST THE SIMULATION'S OWN.
    foreach ($pair in @(
        [pscustomobject]@{ Stamp = 'run_id';              Sim = 'run_id' },
        [pscustomobject]@{ Stamp = 'effective_seed';      Sim = 'effective_seed' },
        [pscustomobject]@{ Stamp = 'request_fingerprint'; Sim = 'request_fingerprint' },
        [pscustomobject]@{ Stamp = 'result_digest';       Sim = 'result_digest' },
        [pscustomobject]@{ Stamp = 'iterations';          Sim = 'iterations_run' })) {
        $stampKey = [string]$pair.Stamp
        $simKey = [string]$pair.Sim
        $null = Add-W5Check ('the stamped ' + $stampKey + ' is the simulation' + [char]39 + 's own') `
            (Test-SimSameValue -A $stamp[$stampKey] -B $simBlock[$simKey]) `
            ('stamp ' + (Format-SimValue $stamp[$stampKey]) + ', run ' +
             (Format-SimValue $simBlock[$simKey]))
    }
    $null = Add-W5Check 'the stamped year count is the project duration' `
        (Test-SimExactDouble -Actual $stamp['year_count'] -Expected ([double]$yearCount)) `
        (Format-SimValue $stamp['year_count'])
    $null = Add-W5Check 'the stamped selected Px label is the requested one' `
        (Test-SimExactText -Actual $stamp['selected_px_label'] -Expected $selectedLabel) `
        (Format-SimValue $stamp['selected_px_label'])
    $null = Add-W5Check 'the stamped selected Px probability is the one that label names' `
        (Test-SimExactDouble -Actual $stamp['selected_px_probability'] `
            -Expected $selectedProbability) `
        ('stamped ' + (Format-SimValue $stamp['selected_px_probability']) + ', ' +
         $selectedLabel + ' names ' + [string]$selectedProbability)

    # -------------------------------------------------------------------
    # THE ANNUAL RECORDS
    # -------------------------------------------------------------------
    Write-W5Line ''
    Write-W5Line 'THE ANNUAL DISTRIBUTION RECORDS'
    Write-W5Line '-------------------------------'
    $ladderCount = [int]$p7.annual_records.quantile_count
    $profileSums = @{ 'nominal' = 0.0; 'pv' = 0.0 }
    $profileScale = @{ 'nominal' = 0.0; 'pv' = 0.0 }
    $recordProblems = New-Object System.Collections.ArrayList
    $ladderProblems = New-Object System.Collections.ArrayList
    for ($offset = 0; $offset -lt $yearCount; $offset++) {
        $record = Get-W5AnnualRecord -Workbook $wb -Inspection $simInspection -P7 $p7 `
            -Bank $bank -Offset $offset
        $expectedIndex = [double]($offset + 1)
        if (-not (Test-SimExactDouble -Actual $record['project_index'] -Expected $expectedIndex)) {
            $null = $recordProblems.Add('row ' + [string]($offset + 1) + ' project_index ' +
                                        (Format-SimValue $record['project_index']))
        }
        $expectedYear = [double]([int]$model.timeline.start_year + $offset)
        if (-not (Test-SimExactDouble -Actual $record['calendar_year'] -Expected $expectedYear)) {
            $null = $recordProblems.Add('row ' + [string]($offset + 1) + ' calendar_year ' +
                                        (Format-SimValue $record['calendar_year']))
        }
        foreach ($measure in @('nominal', 'pv')) {
            $ladder = @($record[('ladder_' + $measure)])
            if ($ladder.Count -ne $ladderCount) {
                $null = $ladderProblems.Add('row ' + [string]($offset + 1) + ' ' + $measure +
                                            ' ladder has ' + [string]$ladder.Count + ' values')
                continue
            }
            $previous = [double]::NegativeInfinity
            for ($index = 0; $index -lt $ladderCount; $index++) {
                $value = $ladder[$index]
                if ($value -isnot [double]) {
                    $null = $ladderProblems.Add('row ' + [string]($offset + 1) + ' ' + $measure +
                                                ' ' + $quantileLabels[$index] + ' published ' +
                                                (Format-SimValue $value))
                    continue
                }
                # ORDERED BY PROBABILITY. The projected ladder is in ascending
                # probability order, and a percentile ladder over one sample can
                # never decrease as the probability rises.
                if ([double]$value -lt $previous) {
                    $null = $ladderProblems.Add('row ' + [string]($offset + 1) + ' ' + $measure +
                                                ' ' + $quantileLabels[$index] + ' (' +
                                                [string]$value + ') is below ' +
                                                $quantileLabels[$index - 1] + ' (' +
                                                [string]$previous + ')')
                }
                $previous = [double]$value
            }
            $profile = $record[('profile_' + $measure)]
            if ($profile -isnot [double]) {
                $null = $recordProblems.Add('row ' + [string]($offset + 1) + ' ' + $measure +
                                            ' profile published ' + (Format-SimValue $profile))
            } else {
                $profileSums[$measure] = $profileSums[$measure] + [double]$profile
                $profileScale[$measure] = $profileScale[$measure] + [Math]::Abs([double]$profile)
            }
        }
    }
    $null = Add-W5Check ('each of the ' + [string]$yearCount +
                         ' project years has one correctly indexed annual record') `
        ($recordProblems.Count -eq 0) ($recordProblems -join '; ')
    $null = Add-W5Check ('both percentile ladders are complete, numeric and ordered in every year') `
        ($ladderProblems.Count -eq 0) ($ladderProblems -join '; ')

    # -------------------------------------------------------------------
    # COMPACT PERSISTENCE
    # -------------------------------------------------------------------
    $region = Get-W5AnnualRegionIndex -Workbook $wb -Inspection $simInspection -P7 $p7 `
        -Bank $bank -Ledger $rel
    $comAcquired = $comAcquired + [int]$region.Acquired
    $regionValues = $region.Values
    $populated = 0
    $firstStray = ''
    $regionRows = 0
    if ($null -ne $regionValues) { $regionRows = [int]$regionValues.GetLength(0) }
    for ($row = 1; $row -le $regionRows; $row++) {
        if (Test-SimBlank -Value $regionValues[$row, 1]) { continue }
        $populated = $populated + 1
        if (($row -gt $yearCount) -and [string]::IsNullOrWhiteSpace($firstStray)) {
            $firstStray = ('record row ' + [string]$row + ' carries ' +
                           (Format-SimValue $regionValues[$row, 1]))
        }
    }
    # ONE RECORD PER PROJECT YEAR AND NOTHING BEYOND. This is what a compact
    # persisted shape means at runtime: no N x Y matrix, and no iteration-level
    # annual workspace materialised as a published record set - either would
    # leave far more than four rows behind in the contracted region.
    $null = Add-W5Check ('the annual region holds exactly ' + [string]$yearCount +
                         ' records, one per project year') `
        (($populated -eq $yearCount) -and [string]::IsNullOrWhiteSpace($firstStray)) `
        ('populated ' + [string]$populated + ' of ' + [string]$regionRows + ' available rows; ' +
         $firstStray)
    # THE ROW IMMEDIATELY AFTER THE AUTHORITATIVE RESULT. A region that held a
    # longer answer before would leave its tail here, and a reader has to be
    # able to see that the current answer ends where the stamp says it does.
    $boundaryRow = $yearCount + 1
    $boundaryValue = $null
    if ($boundaryRow -le $regionRows) { $boundaryValue = $regionValues[$boundaryRow, 1] }
    $null = Add-W5Check 'the row immediately after the annual result is not populated' `
        (Test-SimBlank -Value $boundaryValue) `
        ('row ' + [string]$boundaryRow + ' = ' + (Format-SimValue $boundaryValue))

    # -------------------------------------------------------------------
    # RECONCILIATION AGAINST THE AUTHORITATIVE TOTAL Px
    # -------------------------------------------------------------------
    Write-W5Line ''
    Write-W5Line 'RECONCILIATION: sum_y Profile_Px(y) = reported Px'
    Write-W5Line '------------------------------------------------'
    $grid = $iterationsBefore.Values
    $ladderRowKey = 'quantile_' + [string]($selectedIndex + 1)
    foreach ($measure in @('nominal', 'pv')) {
        $column = 2
        if ($measure -eq 'pv') { $column = 3 }
        $totals = New-Object System.Collections.ArrayList
        for ($row = 1; $row -le [int]$grid.GetLength(0); $row++) {
            $value = $grid[$row, $column]
            if ($value -is [double]) { $null = $totals.Add([double]$value) }
        }
        $derived = Get-W5Type7Value -Values ([double[]]@($totals)) -Probability $selectedProbability
        $published = Get-SimRawCell -Workbook $wb -Inspection $simInspection `
            -Address ([string]$simInspection.sim_data.contingency_ladder.bank_value_columns.$bank.$measure +
                      [string]([int]$simInspection.sim_data.contingency_ladder.rows.$ladderRowKey))

        # THE PUBLISHED LADDER IS THE AUTHORITATIVE TOTAL Px, and the contract's
        # own Type-7 over the published iteration column is the independent
        # route to the same number. They are compared to each other first, so a
        # reconciliation that passed against a wrong total would be caught.
        $ladderScale = 0.0
        if ($published -is [double]) { $ladderScale = [Math]::Abs([double]$published) }
        $ladderScale = $ladderScale + [Math]::Abs($derived)
        $ladderAllowance = Get-W5IdentityAllowance -Provenance $cases.provenance `
            -ConditioningScale $ladderScale
        $ladderDelta = [double]::PositiveInfinity
        if ($published -is [double]) { $ladderDelta = [Math]::Abs([double]$published - $derived) }
        $null = Add-W5Check ('the published ' + $measure + ' ' + $selectedLabel +
                             ' equals the contract' + [char]39 + 's Type-7 value over the iteration column') `
            (($published -is [double]) -and ($ladderDelta -le $ladderAllowance)) `
            ('published ' + (Format-SimValue $published) + ', derived ' + [string]$derived +
             ', delta ' + [string]$ladderDelta + ', allowance ' + [string]$ladderAllowance)

        # AND THE IDENTITY ITSELF. The conditioning scale names the magnitude of
        # the arithmetic performed - the annual terms summed, plus the aggregate
        # they are compared against - never the magnitude of the net result.
        $sum = [double]$profileSums[$measure]
        $scale = [double]$profileScale[$measure]
        if ($published -is [double]) { $scale = $scale + [Math]::Abs([double]$published) }
        $identityAllowance = Get-W5IdentityAllowance -Provenance $cases.provenance `
            -ConditioningScale $scale
        $delta = [double]::PositiveInfinity
        if ($published -is [double]) { $delta = [Math]::Abs($sum - [double]$published) }
        $null = Add-W5Check ('the ' + $measure + ' selected-Px profile sums to the reported ' +
                             $selectedLabel) `
            (($published -is [double]) -and ($delta -le $identityAllowance)) `
            ('sum ' + [string]$sum + ', reported ' + (Format-SimValue $published) +
             ', delta ' + [string]$delta + ', allowance ' + [string]$identityAllowance +
             ' (conditioning scale ' + [string]$scale + ')')
        Write-W5Line ('    ' + $measure.PadRight(8) + ' sum ' + [string]$sum +
                      '  reported ' + (Format-SimValue $published) +
                      '  delta ' + [string]$delta)
    }

    # -------------------------------------------------------------------
    # THE PHASE-8 HANDOFF
    # -------------------------------------------------------------------
    Write-W5Line ''
    Write-W5Line 'THE PHASE-8 HANDOFF'
    Write-W5Line '-------------------'
    $handoff = Get-W5Handoff -Excel $excel -P7 $p7
    $accessors = @($p7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
    $current = [string]$p7.handoff.distribution_states[1]
    $profileCurrent = [string]$p7.handoff.profile_states[1]
    $null = Add-W5Check ($accessors[0] + ' reports the projected CURRENT state') `
        ((([string]$handoff[$accessors[0]]) -ceq $current)) `
        ((Format-SimValue $handoff[$accessors[0]]) + ' vs ' + $current)
    $null = Add-W5Check ($accessors[1] + ' reports the projected CURRENT state') `
        ((([string]$handoff[$accessors[1]]) -ceq $profileCurrent)) `
        ((Format-SimValue $handoff[$accessors[1]]) + ' vs ' + $profileCurrent)
    $null = Add-W5Check ($accessors[2] + ' reports the selected Px identity') `
        ((([string]$handoff[$accessors[2]]) -match ([regex]::Escape($selectedLabel)))) `
        (Format-SimValue $handoff[$accessors[2]])
    $null = Add-W5Check ($accessors[3] + ' reports the project year count') `
        (([int]$handoff[$accessors[3]]) -eq $yearCount) (Format-SimValue $handoff[$accessors[3]])

    # -------------------------------------------------------------------
    # THE SIDE-EFFECT INVARIANTS
    # -------------------------------------------------------------------
    Write-W5Line ''
    Write-W5Line 'OBSERVATIONAL SIDE-EFFECT INVARIANTS'
    Write-W5Line '------------------------------------'
    $invariantsAfter = Get-W5RunInvariants -Workbook $wb -Inspection $simInspection
    $null = Add-W5InvariantChecks 'the successful annual run' $invariantsBefore $invariantsAfter

    $iterationsAfter = Get-W5IterationBlock -Workbook $wb -Inspection $simInspection `
        -Bank $bank -Count $iterations -Ledger $rel
    $comAcquired = $comAcquired + [int]$iterationsAfter.Acquired
    $before = $iterationsBefore.Values
    $after = $iterationsAfter.Values
    $moved = ''
    if (($null -eq $before) -or ($null -eq $after) -or
        ($before.GetLength(0) -ne $after.GetLength(0)) -or
        ($before.GetLength(1) -ne $after.GetLength(1))) {
        $moved = 'the iteration block changed shape'
    } else {
        for ($row = 1; $row -le [int]$before.GetLength(0); $row++) {
            for ($column = 1; $column -le [int]$before.GetLength(1); $column++) {
                if (-not (Test-SimSameValue -A $before[$row, $column] -B $after[$row, $column])) {
                    $moved = ('iteration row ' + [string]$row + ' column ' + [string]$column +
                              ': ' + (Format-SimValue $before[$row, $column]) + ' -> ' +
                              (Format-SimValue $after[$row, $column]))
                    break
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($moved)) { break }
        }
    }
    # ANNUAL IS OBSERVATIONAL POST-PROCESSING, NOT A SECOND STOCHASTIC RUN. Every
    # published iteration index, nominal total and PV total must be the same
    # value, of the same type, after the annual step as before it.
    $null = Add-W5Check 'every published iteration value is unchanged by the annual run' `
        ([string]::IsNullOrWhiteSpace($moved)) $moved

    Write-W5Line ''
    Write-W5Line (Format-Phase6State -State (Get-Phase6State -Workbook $wb `
        -Inspection $simInspection) -Label '  the run identity AFTER the annual step')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W5Line ''
    Write-W5Line ('THE W5 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W5Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W5Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W5Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W5Line ''
        Write-W5Line 'EXCEL SHUTDOWN'
        Write-W5Line '--------------'
        if ($naturalExit) {
            Write-W5Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W5'
            Write-W5Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W5Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W5Residual)) {
            Write-W5Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W5Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W5Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W5Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W5Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W5Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W5Check 'every COM release left 0 outstanding references' `
    (@($script:W5Residual).Count -eq 0) ((@($script:W5Residual)) -join '; ')
$null = Add-W5Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W5Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W5Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W5Line ''
Write-W5Line 'VERDICT'
Write-W5Line '-------'
Write-W5Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W5Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W5Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W5Line ''
if ($ok) {
    Write-W5Line 'W5: PASS'
} else {
    Write-W5Line 'W5: FAIL'
    Write-W5Line ''
    Write-W5Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W5Line ''
Write-W5Line ('report                 : ' + $script:W5Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W5Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
