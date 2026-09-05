<#
.SYNOPSIS
    PCCM Phase 7 - the MINIMAL W2 runner: many DriverFactors instances, against
    the accepted Phase-5 oracle.

.DESCRIPTION
    W2 IS AN INHERITED PHASE-5 REGRESSION SCENARIO. Phase 7 gave DriverFactors
    two dynamic arrays per driver. W2 asks one question about that change: with
    ~300 driver instances each carrying those arrays, does the inherited Phase-5
    calculation path still publish the answer the INDEPENDENT Phase-5 oracle
    computed for the very same model. Nothing else. No annual stochastic run, no
    sensitivity, no W3, and no repeat of the W1 surface matrix.

    THE DIMENSION IS DRIVER COUNT, ON PURPOSE. 300 drivers over 5 project years.
    The array LENGTH dimension is W3's and is not combined with this one: the
    two runtime risks are independent, and one Cartesian fixture would cost far
    more Windows time while proving neither more strongly.

    THE AUTHORITY IS NOT ANOTHER VBA CALCULATION. Both the fixture model and the
    expected answer are read from `build/phase7_acceptance_cases.json`, whose
    expectations come from `pccm_builder.calc_oracle.calculate` - the same
    independent Phase-5 implementation the accepted `phase5_cases.json` corpus
    is built from, which Phase 7 did not touch. The model the workbook is given
    and the model the expectation was computed from are the same bytes of the
    same artefact, so a disagreement is a disagreement about the calculation and
    not about the fixture.

    WHAT IT REUSES, AND WHY IT IS NOT A SECOND HARNESS
    -------------------------------------------------
    The workbook is populated through the ACCEPTED Phase-5 fixture - which
    drives production's own Add endpoints, so this runner writes no register row
    of its own - and read back through the ACCEPTED TYPED Phase-5 reader, so a
    number published as text FAILS instead of passing through a stringifying
    reader. Those live in `phase5_gate_b_scenarios.ps1`, and the value
    comparison lives in `phase6_gate_b_scenarios.ps1`. Both are definition-only
    at top level: dot-sourcing them defines functions and runs no scenario, so
    no Gate-B result is produced, recorded or implied.

    TEN HELPERS ARE COPIED RATHER THAN DOT-SOURCED, for the reason W1 defect #2
    established: the accepted fixture calls them, and their owner
    (`phase4_functional_test.ps1`) is NOT definition-only - dot-sourcing it would
    run the entire Phase-4 matrix. They are copied byte for byte from
    `phase7_timing_scenarios.ps1` and a control pins them to it, and a second
    control walks the whole call closure so a seventh missing helper is found
    here rather than by a failed Windows session.

    WHAT IT PROVES
    --------------
      1. the candidate identity: git HEAD, with pccm/src, pccm/spec and
         pccm/builder clean, proved BEFORE Excel is started
      2. a Stage-B workbook generated from the CURRENT Stage-A build
      3. the current project compiles (one lightweight read-only trigger; the
         full surface matrix is W1's and is not repeated)
      4. the 300-driver model is applied through the accepted input paths
      5. PCCM_Calculate succeeds and the workbook reports CURRENT
      6. all 300 drivers are published, in supply order, once each
      7. every published driver matches the oracle BY ITS OWN ID - the
         misalignment control - and again ROW FOR ROW - the ordering control
      8. the 5-year axis is 5 years long, and the per-year arrays behind it
         match: calc_years, calc_inflation_factors and calc_annual
      9. the ten published totals match
     10. the owned Excel process exits NATURALLY, with no emergency cleanup

    LIFECYCLE. The discipline W1 closed on: this runner's own acquisitions are
    counted where they happen, released through one ledger that exists before
    Excel does, and every ReleaseComObject count is recorded - a non-zero
    residual is a FAILING check that names the object. The accepted Phase-5 and
    Phase-6 readers own their own transients under the accepted
    `com_lifecycle.ps1` policy; extending the residual accounting into them
    would mean editing accepted files, so instead their own gate is reported:
    `Get-TransientFailures` must be empty.

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
$script:W2Lines = New-Object System.Collections.ArrayList
$script:W2Path = ''
$script:W2Checks = New-Object System.Collections.ArrayList

function Write-W2Line {
    param([string]$Text = '')
    $null = $script:W2Lines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:W2Path)) {
        try {
            Set-Content -LiteralPath $script:W2Path `
                -Value ($script:W2Lines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step W2 had to perform to reach
# its subject - from a RESULT, which is W2's own claim about the calculation. A
# failed prerequisite still fails the invocation; it is just not evidence about
# the property under test, and the report says which it was.
function Add-W2Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:W2Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-W2Line $line
    return $Ok
}

function Format-W2Value {
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
$script:W2Residual = New-Object System.Collections.ArrayList

function Invoke-W2Release {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:W2Residual.Add(
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
function Get-W2SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a W2 result could not be attributed to a source revision')
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
function Compare-W2Cell {
    param($Got, $Expect, [double]$Allowance)
    if ($null -eq $Expect) {
        if (($null -eq $Got) -or ($Got -is [System.DBNull]) -or
            (($Got -is [string]) -and ([string]$Got).Length -eq 0)) { return '' }
        return ('published ' + (Format-W2Value $Got) + ' where the oracle has none')
    }
    if ($Expect -is [string]) {
        if (($Got -is [string]) -and (([string]$Got) -ceq ([string]$Expect))) { return '' }
        return ((Format-W2Value $Got) + ' vs ' + [string]$Expect)
    }
    if ($Got -isnot [double]) {
        return ((Format-W2Value $Got) + ' is not a number')
    }
    $difference = [Math]::Abs([double]$Got - [double]$Expect)
    if ($difference -le $Allowance) { return '' }
    return ([string]$Got + ' vs ' + [string]$Expect + ' (difference ' + [string]$difference + ')')
}

# THE PUBLISHED TABLE, ROW FOR ROW, EVERY PROJECTED COLUMN. A failing table
# reports the first twelve disagreements and the count: the count says how bad
# it is, the examples make it diagnosable, and 6,300 lines would do neither.
function Compare-W2Table {
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
            $problem = Compare-W2Cell -Got $Live[$index][$at] -Expect $expect -Allowance $Allowance
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
    return (Add-W2Check ($Label + ': ' + $TableKey +
                         ' matches the accepted Phase-5 oracle in every projected column') `
        ($problems.Count -eq 0) $detail)
}

# ===========================================================================
# PREFLIGHT
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 W2 (many DriverFactors instances)' -ForegroundColor Cyan
Write-Host '===============================================' -ForegroundColor Cyan
Write-Host ''

$manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath  = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$casesPath    = Join-Path $BuildDir 'phase7_acceptance_cases.json'
foreach ($required in @($manifestPath, $inspectPath, $casesPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ($required + ' not found. Run the Stage-A build first: ' +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest   = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath $inspectPath  -Raw | ConvertFrom-Json
$cases      = Get-Content -LiteralPath $casesPath    -Raw | ConvertFrom-Json

$case = $cases.scenarios | Where-Object { [string]$_.id -ceq 'W2' }
if ($null -eq $case) {
    Write-Host 'the acceptance corpus carries no W2 scenario.' -ForegroundColor Red
    exit 1
}

$revision = $null
try { $revision = Get-W2SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a W2 result could ' +
                'not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-w2-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $casesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:W2Path = Join-Path $tempRoot 'phase7_w2_many_drivers.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

# INVOKED AS A CHILD SCRIPT, never dot-sourced: build_stage_b.ps1 has
# executable top-level code, and this runner takes on none of it.
$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
$bootstrapOk = (($bootstrapExit -eq 0) -and (Test-Path -LiteralPath $stageBPath))

$model = $case.model
$driverCount = @($model.cost_lines).Count + @($model.risks).Count
$yearCount = [int]$model.timeline.duration
$allowance = [double]$cases.provenance.comparison_absolute_floor

Write-W2Line 'PCCM - PHASE 7 W2: MANY DRIVERFACTORS INSTANCES'
Write-W2Line '=============================================='
Write-W2Line ''
Write-W2Line 'This is the MINIMAL W2 runner. It is an INHERITED PHASE-5 REGRESSION'
Write-W2Line 'scenario: no annual stochastic run, no sensitivity, no Gate-B result, and'
Write-W2Line 'no repeat of the W1 surface matrix. The historical Phase-6 runtime'
Write-W2Line 'authority remains Run 6 on its own closure commit.'
Write-W2Line ''
Write-W2Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-W2Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-W2Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-W2Line ('git HEAD               : ' + [string]$revision.Head)
Write-W2Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-W2Line ('model version          : ' + [string]$manifest.model_version)
Write-W2Line ('calc contract version  : ' + [string]$cases.provenance.calc_contract_version)
Write-W2Line ('build directory        : ' + $BuildDir)
Write-W2Line ('working copy           : ' + $tempRoot)
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
    Write-W2Line ('  ' + ([string]$artefact.Label).PadRight(20) + ' SHA-256 ' + $shown)
}
Write-W2Line ''
Write-W2Line 'THE SCENARIO AND ITS AUTHORITY'
Write-W2Line '------------------------------'
Write-W2Line ('  dimension under test : ' + [string]$case.dimension)
Write-W2Line ('  drivers              : ' + [string]$driverCount +
              ' (' + [string]@($model.cost_lines).Count + ' cost lines, ' +
              [string]@($model.risks).Count + ' risks)')
Write-W2Line ('  project years        : ' + [string]$yearCount)
Write-W2Line ('  expectation source   : ' + [string]$cases.provenance.expectation_source)
Write-W2Line ('  comparison allowance : ' + [string]$allowance +
              ' (the accepted Phase-5 identity absolute floor; no tolerance is invented here)')
Write-W2Line ''

$null = Add-W2Check 'the Stage-B workbook was generated from the current Stage-A build' `
    $bootstrapOk ('bootstrap exit ' + [string]$bootstrapExit) 'PREREQUISITE'
if (-not $bootstrapOk) {
    Write-W2Line ''
    Write-W2Line 'STOP. Excel was never started for the W2 session, and nothing was accepted.'
    Write-W2Line ('report                 : ' + $script:W2Path)
    Write-W2Line 'W2: FAIL'
    exit 1
}

# THE CORPUS IS THE AUTHORITY, so its identity is checked before it is trusted.
$null = Add-W2Check 'the expectations come from the independent Phase-5 oracle' `
    (([string]$cases.provenance.expectation_source) -ceq 'pccm_builder.calc_oracle.calculate') `
    ([string]$cases.provenance.expectation_source) 'PREREQUISITE'
$null = Add-W2Check 'the W2 case is the driver-count dimension' `
    (([string]$case.dimension) -ceq 'driver_count') ([string]$case.dimension) 'PREREQUISITE'
$null = Add-W2Check 'the W2 fixture is the authorised size' `
    (($driverCount -ge 250) -and ($driverCount -le 350) -and ($yearCount -eq 5)) `
    ([string]$driverCount + ' drivers over ' + [string]$yearCount + ' years') 'PREREQUISITE'
$null = Add-W2Check 'the oracle carries one expected driver record per model driver' `
    (@($case.expected.drivers).Count -eq $driverCount) `
    ([string]@($case.expected.drivers).Count + ' expected, ' + [string]$driverCount + ' supplied') `
    'PREREQUISITE'

# ===========================================================================
# THE W2 SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$naturalExit = $false
$emergencyRequired = $false
$fatal = ''
$comAcquired = 0

# THE LEDGER EXISTS BEFORE EXCEL DOES, so every release of the run lands in one
# record with its count.
$rel = New-ReleaseLedger 'phase 7 W2 session'

try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-W2Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is')
    Write-W2Line 'ever terminated, and the workbook is never saved.'
    Write-W2Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    Write-W2Line 'PREREQUISITES'
    Write-W2Line '-------------'

    # THE COMPILE TRIGGER, DELIBERATELY LIGHT. VBA compiles the whole project
    # before it executes any statement, so the first Application.Run IS a
    # full-project compile - and a published READ accessor runs nothing, writes
    # nothing and consumes no run identity. W1 owns the surface matrix; W2 only
    # needs to know that the project Excel is about to calculate with builds.
    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $compiled = [string]::IsNullOrWhiteSpace($compileFailure)
    $null = Add-W2Check 'the current VBAProject compiles in real Excel' `
        $compiled $compileFailure 'PREREQUISITE'
    if (-not $compiled) {
        Write-W2Line ''
        Write-W2Line 'STOP. Nothing observed after a compile failure is evidence about anything.'
        throw ('the VBAProject does not compile: ' + $compileFailure)
    }

    # The accepted fixture restores the locked FX seed rather than keeping
    # whatever row 1 happens to hold, so the seed is captured before it runs.
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # THE FIXTURE IS THE MODEL FROM THE ARTEFACT, applied through the accepted
    # choreography, which drives production's own Add endpoints. This runner
    # writes no register row, no grid cell and no timeline of its own.
    $applied = ''
    $fixtureFailure = ''
    try {
        $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
            -Inspection $inspection -Model $model)
    } catch { $fixtureFailure = (Format-Err $_) }
    $null = Add-W2Check ('the ' + [string]$driverCount +
                         '-driver model was applied through the accepted fixture') `
        (($applied -like 'OK|*') -and [string]::IsNullOrWhiteSpace($fixtureFailure)) `
        ($applied + $fixtureFailure) 'PREREQUISITE'
    if (-not ($applied -like 'OK|*')) {
        throw ('the W2 model was not applied: ' + $applied + $fixtureFailure)
    }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

    # -------------------------------------------------------------------
    # THE INHERITED PHASE-5 CALCULATION, THROUGH ITS OWN PATH
    # -------------------------------------------------------------------
    $result = ''
    $calcFailure = ''
    try {
        $result = [string](Invoke-Phase5ProductionOperation -Excel $excel `
            -Operation 'PCCM_Calculate' -Stage 'W2 calculate')
    } catch { $calcFailure = (Format-Err $_) }
    $null = Add-W2Check 'PCCM_Calculate succeeded on the 300-driver model' `
        ([string]::IsNullOrWhiteSpace($calcFailure)) ($result + $calcFailure)
    if (-not [string]::IsNullOrWhiteSpace($calcFailure)) {
        throw ('PCCM_Calculate did not succeed: ' + $calcFailure)
    }

    $status = [string]$excel.Run('PCCM_CalculationStatus')
    $null = Add-W2Check 'the calculation reports CURRENT' ($status -ceq 'CURRENT') $status

    # THE TIMELINE THE CALCULATION ACTUALLY APPLIED, from the workbook's own
    # state block. This is the first thing that would move if the 5-year axis
    # had been truncated or extended anywhere in the path.
    $appliedTimeline = Get-CalcScalar -Workbook $wb -Inspection $inspection `
        -Block 'calc_state' -FieldKey 'last_successful_applied_timeline'
    $null = Add-W2Check 'the applied timeline is the one the oracle computed against' `
        ((([string]$appliedTimeline) -ceq ([string]$case.expected.applied_timeline))) `
        ((Format-W2Value $appliedTimeline) + ' vs ' + [string]$case.expected.applied_timeline)

    Write-W2Line ''
    Write-W2Line 'THE PUBLISHED ANSWER AGAINST THE INDEPENDENT ORACLE'
    Write-W2Line '---------------------------------------------------'

    # -------------------------------------------------------------------
    # ALL 300 DRIVERS, PRESENT, ONCE EACH, IN SUPPLY ORDER
    # -------------------------------------------------------------------
    $liveDrivers = @(Get-CalcTableRows -Workbook $wb -Inspection $inspection -TableKey 'calc_drivers')
    $expectedDrivers = @($case.expected.drivers)
    $null = Add-W2Check ('calc_drivers publishes one row per supplied driver (' +
                         [string]$driverCount + ')') `
        ($liveDrivers.Count -eq $expectedDrivers.Count) `
        ('published ' + [string]$liveDrivers.Count + ', expected ' + [string]$expectedDrivers.Count)

    $idAt = Get-CalcTableColumnIndex -Inspection $inspection -TableKey 'calc_drivers' `
        -ColumnKey 'permanent_id'
    if (($idAt -ge 0) -and ($liveDrivers.Count -eq $expectedDrivers.Count)) {
        # SUPPLY ORDER, EXACTLY: the ids the oracle holds, in the order the
        # fixture supplied them, with nothing dropped, repeated or transposed.
        $liveIds = @()
        foreach ($row in $liveDrivers) { $liveIds += [string]$row[$idAt] }
        $wantIds = @()
        foreach ($record in $expectedDrivers) { $wantIds += [string]$record.permanent_id }
        $firstDivergence = ''
        for ($index = 0; $index -lt $wantIds.Count; $index++) {
            if ($liveIds[$index] -cne $wantIds[$index]) {
                $firstDivergence = ('row ' + [string]($index + 1) + ': ' + $liveIds[$index] +
                                    ' where the model supplied ' + $wantIds[$index])
                break
            }
        }
        $null = Add-W2Check 'the published drivers are in supply order, one row per driver' `
            ([string]::IsNullOrWhiteSpace($firstDivergence)) $firstDivergence
        $distinct = @($liveIds | Sort-Object -Unique)
        $null = Add-W2Check 'no driver is published twice and none is missing' `
            ($distinct.Count -eq $wantIds.Count) `
            ([string]$distinct.Count + ' distinct ids across ' + [string]$liveIds.Count + ' rows')

        # THE MISALIGNMENT CONTROL. Every published row is matched to the oracle
        # record its OWN id names, not to the record in the same position. If the
        # factors of one driver were ever copied onto another - the failure a
        # shared or reused DriverFactors instance would produce - this fails
        # while the ordering check above still passes, and the report says so.
        $byId = @{}
        foreach ($record in $expectedDrivers) { $byId[[string]$record.permanent_id] = $record }
        $misaligned = New-Object System.Collections.ArrayList
        $columns = @($inspection.calc.tables.calc_drivers.columns)
        for ($index = 0; $index -lt $liveDrivers.Count; $index++) {
            $id = [string]$liveDrivers[$index][$idAt]
            if (-not $byId.ContainsKey($id)) {
                $null = $misaligned.Add('row ' + [string]($index + 1) + ': ' + $id +
                                        ' is not a driver the model supplied')
                continue
            }
            $record = $byId[$id]
            foreach ($column in $columns) {
                $at = Get-CalcTableColumnIndex -Inspection $inspection -TableKey 'calc_drivers' `
                    -ColumnKey $column
                $problem = Compare-W2Cell -Got $liveDrivers[$index][$at] `
                    -Expect $record.$column -Allowance $allowance
                if (-not [string]::IsNullOrWhiteSpace($problem)) {
                    $null = $misaligned.Add($id + ' ' + [string]$column + ': ' + $problem)
                }
            }
            if ($misaligned.Count -gt 12) { break }
        }
        $detail = ''
        if ($misaligned.Count -gt 0) {
            $detail = ($misaligned -join '; ')
            if ($misaligned.Count -gt 12) { $detail = $detail + ' ... (truncated)' }
        }
        $null = Add-W2Check ('every driver carries the factors of the driver its own id names') `
            ($misaligned.Count -eq 0) $detail

        # AND THE SAME COMPARISON BY POSITION. Both passing is the answer; this
        # one failing while the one above passes means the numbers are right and
        # the ORDER is wrong, which is a different defect and a different fix.
        $null = Compare-W2Table -Live $liveDrivers -Expected $expectedDrivers `
            -Inspection $inspection -TableKey 'calc_drivers' -Allowance $allowance -Label 'W2'
    }

    # -------------------------------------------------------------------
    # THE FIVE-YEAR AXIS AND THE PER-YEAR ARRAYS BEHIND IT
    # -------------------------------------------------------------------
    # Knom and Kpv are built from these arrays, so a truncated, shared or
    # misindexed year array shows up here as well as in every driver above.
    $liveYears = @(Get-CalcTableRows -Workbook $wb -Inspection $inspection -TableKey 'calc_years')
    $null = Add-W2Check ('calc_years publishes exactly ' + [string]$yearCount + ' project years') `
        ($liveYears.Count -eq @($case.expected.calc_years).Count) `
        ('published ' + [string]$liveYears.Count + ', expected ' +
         [string]@($case.expected.calc_years).Count)
    if ($liveYears.Count -eq @($case.expected.calc_years).Count) {
        $null = Compare-W2Table -Live $liveYears -Expected $case.expected.calc_years `
            -Inspection $inspection -TableKey 'calc_years' -Allowance $allowance -Label 'W2'
    }

    # THE INFLATION FACTOR ARRAY ITSELF, one row per profile per year. The
    # corpus names two of its fields differently from the projected columns, so
    # the two names are mapped rather than assumed equal.
    $inflationMap = @{ 'inflation_profile' = 'profile'; 'cumulative_inflation_factor' = 'cumulative_factor' }
    $liveInflation = @(Get-CalcTableRows -Workbook $wb -Inspection $inspection `
        -TableKey 'calc_inflation_factors')
    $wantInflation = @($case.expected.inflation_factors)
    $null = Add-W2Check 'calc_inflation_factors publishes one row per profile per project year' `
        ($liveInflation.Count -eq $wantInflation.Count) `
        ('published ' + [string]$liveInflation.Count + ', expected ' + [string]$wantInflation.Count)
    if ($liveInflation.Count -eq $wantInflation.Count) {
        $null = Compare-W2Table -Live $liveInflation -Expected $wantInflation `
            -Inspection $inspection -TableKey 'calc_inflation_factors' -Allowance $allowance `
            -Label 'W2' -FieldMap $inflationMap
    }

    # THE ANNUAL DECOMPOSITION Phase 5 publishes: the per-year spread of the
    # same arrays, summed across all 300 drivers.
    $liveAnnual = @(Get-CalcTableRows -Workbook $wb -Inspection $inspection -TableKey 'calc_annual')
    $null = Add-W2Check ('calc_annual publishes exactly ' + [string]$yearCount + ' years') `
        ($liveAnnual.Count -eq @($case.expected.annual).Count) `
        ('published ' + [string]$liveAnnual.Count + ', expected ' +
         [string]@($case.expected.annual).Count)
    if ($liveAnnual.Count -eq @($case.expected.annual).Count) {
        $null = Compare-W2Table -Live $liveAnnual -Expected $case.expected.annual `
            -Inspection $inspection -TableKey 'calc_annual' -Allowance $allowance -Label 'W2'
    }

    # -------------------------------------------------------------------
    # THE TEN PUBLISHED TOTALS
    # -------------------------------------------------------------------
    $totalProblems = @()
    foreach ($key in $case.expected.totals.PSObject.Properties.Name) {
        $got = Get-CalcScalar -Workbook $wb -Inspection $inspection `
            -Block 'calc_totals' -FieldKey $key
        $problem = Compare-W2Cell -Got $got -Expect $case.expected.totals.$key -Allowance $allowance
        if (-not [string]::IsNullOrWhiteSpace($problem)) {
            $totalProblems += ([string]$key + ': ' + $problem)
        }
    }
    $null = Add-W2Check 'the ten published totals match the accepted Phase-5 oracle' `
        ($totalProblems.Count -eq 0) ($totalProblems -join '; ')

    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

    # The automation guard is ended the way the proven driver ends it.
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-W2Line ''
    Write-W2Line ('THE W2 SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-W2Release $rel $wb        'Workbook';  $wb        = $null
        Invoke-W2Release $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-W2Release $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        # CLEARED BEFORE THE COLLECT. An ErrorRecord from a failed COM call still
        # references the object the call was made on, and a rooted RCW is one the
        # collects cannot reclaim.
        $Error.Clear()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        # 90 SECONDS, EXPLICITLY. The accepted default is 25; a workbook holding
        # 300 driver rows is the slowest close in this repo, and a bound that
        # cannot tell "still referenced" from "still closing" decides the verdict
        # for the wrong reason.
        if ($null -ne $excelIdentity) {
            $naturalExit = Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90
        }
        $rel.NaturalExit = $naturalExit
        Write-W2Line ''
        Write-W2Line 'EXCEL SHUTDOWN'
        Write-W2Line '--------------'
        if ($naturalExit) {
            Write-W2Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            # THE SAFETY NET, WHICH IS NEVER A PASS.
            $emergencyRequired = $true
            $rel.EmergencyRequired = $true
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity -Label 'W2'
            Write-W2Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
        Write-W2Line (Format-ReleaseLedger $rel)
        foreach ($residual in @($script:W2Residual)) {
            Write-W2Line ('      OUTSTANDING: ' + [string]$residual)
        }
        foreach ($transient in @(Get-TransientFailures)) {
            Write-W2Line ('      transient release FAILED: ' + [string]$transient)
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$null = Add-W2Check 'the owned Excel process exited naturally' $naturalExit
$null = Add-W2Check 'no emergency cleanup was required' (-not $emergencyRequired)
$null = Add-W2Check 'every COM object this runner acquired was released' `
    ([int]$rel.Attempted -eq [int]$comAcquired) `
    ([string]$comAcquired + ' acquired, ' + [string]$rel.Attempted + ' released')
$null = Add-W2Check 'every COM release succeeded' ($rel.Failed.Count -eq 0) ($rel.Failed -join ', ')
$null = Add-W2Check 'every COM release left 0 outstanding references' `
    (@($script:W2Residual).Count -eq 0) ((@($script:W2Residual)) -join '; ')
# The accepted Phase-5 and Phase-6 readers own their own transients under the
# accepted lifecycle policy. This is that policy's own gate, reported here.
$null = Add-W2Check 'every transient release inside the accepted readers succeeded' `
    (@(Get-TransientFailures).Count -eq 0) ((@(Get-TransientFailures)) -join '; ')

$results = @($script:W2Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:W2Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-W2Line ''
Write-W2Line 'VERDICT'
Write-W2Line '-------'
Write-W2Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-W2Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    $suffix = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$failure.Detail)) {
        $suffix = ' -- ' + [string]$failure.Detail
    }
    Write-W2Line ('  FAILED [' + [string]$failure.Kind + '] ' + [string]$failure.Label + $suffix)
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-W2Line ''
if ($ok) {
    Write-W2Line 'W2: PASS'
} else {
    Write-W2Line 'W2: FAIL'
    Write-W2Line ''
    Write-W2Line 'STOP AND REVIEW. Do not run any later scenario until this is understood.'
}
Write-W2Line ''
Write-W2Line ('report                 : ' + $script:W2Path)
Write-Host ''
Write-Host ('The report is at ' + $script:W2Path) -ForegroundColor Cyan
Write-Host 'The working copy is left in place so the report survives; delete it when done.'
if ($ok) { exit 0 } else { exit 1 }
