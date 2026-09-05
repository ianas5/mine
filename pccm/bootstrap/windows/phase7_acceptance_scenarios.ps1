<#
.SYNOPSIS
    PCCM Phase 7 - the WINDOWS ACCEPTANCE harness (W1-W8).

.DESCRIPTION
    THIS IS THE PHASE-7 ACCEPTANCE SESSION. It is not Gate B, it is not the
    Phase-7 timing harness, and it is not a second static-proof campaign. It
    establishes LIVE Excel/VBA evidence for the material runtime risks Phase 7
    introduced since the accepted Phase-6 Run 6, and nothing else.

    IT IS ADDITIVE. `phase5_gate_b_scenarios.ps1` and `phase6_gate_b_scenarios.ps1`
    are dot-sourced for their PRIMITIVES - the fixture choreography, the _SimData
    readers, the comparison helpers - and neither is modified. Nothing here calls
    Invoke-Phase5GateBScenarios or Invoke-Phase6GateBScenarios, so no Gate-B
    result is produced, recorded or implied, and the historical Phase-6 harness
    identity is untouched.

    ONE SCENARIO PER INVOCATION, ON PURPOSE
    ---------------------------------------
    P7-7 is executed one Windows step at a time: run one, read the report,
    decide whether to proceed. So `-Scenario` takes exactly one id and this
    script runs exactly that scenario. A scenario that needs earlier state
    ESTABLISHES IT ITSELF, through the same production paths, and reports those
    steps as PREREQUISITES rather than as results - so every invocation is
    self-contained and every report can be read on its own.

    EVERY SCENARIO COMPILES THE PROJECT FIRST. W1 is the compile scenario, but
    the compile check runs at the start of all of them: behavioural evidence
    taken from a project that does not compile is not evidence, and making the
    ordering requirement structural is better than making it procedural.

    WHAT IT NEVER DOES
    ------------------
      * It does not modify production VBA, and it imports no diagnostic module.
        Every observation is a read of a published cell or a call of a published
        endpoint.
      * It does not write a register row, a grid cell or a timeline of its own.
        Models are established through `Set-Phase5Fixture`, which drives
        production's own Add endpoints.
      * It does not save the workbook. The real build output is copied to a
        temporary directory and only the copy is ever opened.
      * It does not alter a security setting, touch the registry, add a Trusted
        Location, or terminate an Excel process it did not create.
      * It reaches STALE and INVALID through ACCEPTED INPUT PATHS - the Monte
        Carlo iteration control and a driver register edit - never by writing to
        a hidden machine sheet.

    THE TWO INHERITED-BEHAVIOUR SCENARIOS COMPARE AGAINST PRE-PHASE-7 AUTHORITY.
    W2 and W3 compare the live `_Calc` output against
    `build/phase7_acceptance_cases.json`, whose expectations are produced by
    `pccm_builder.calc_oracle.calculate` - the independent Phase-5 oracle the
    accepted `phase5_cases.json` corpus is built from, and which Phase 7 did not
    touch. The FIXTURE MODEL is read from that same artefact, so the model the
    workbook is given and the model the expectation was computed from are the
    same bytes.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.PARAMETER Scenario
    Exactly one of W1..W8. There is no 'All': see ONE SCENARIO PER INVOCATION.

.PARAMETER KeepArtifacts
    Leave the temporary workbook and the report on disk.

.NOTES
    SHUTDOWN is the accepted `com_lifecycle.ps1` path: Workbook.Close,
    Application.Quit, named releases leaf-before-parent, Wait-ExcelExit, and
    Invoke-EmergencyExcelCleanup ONLY for a process whose identity is still
    positively verified.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [Parameter(Mandatory = $true)]
    [ValidateSet('W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8')]
    [string]$Scenario,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# THE ACCEPTED FILES ARE REUSED, NOT REIMPLEMENTED. All three are
# definition-only at top level: dot-sourcing them defines functions and script
# variables and runs no scenario.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

# WINDOWS POWERSHELL 5.1. `Join-Path a b c` - a child per positional argument -
# is PowerShell 6+ only; on 5.1 Join-Path takes exactly -Path and -ChildPath and
# the third argument is refused with "A positional parameter cannot be found".
# The acceptance machine runs 5.1, so this is the form all four accepted
# harnesses already use, and it is used here for the same reason.
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# ===========================================================================
# THE DECLARED MATRIX
# ===========================================================================
# Purpose and PREREQUISITES, declared once. The prerequisite chain is data, so
# the report can state what a scenario had to establish before it could begin
# and a reader never has to infer it.
function Get-Phase7AcceptanceScenarios {
    return @(
        [pscustomobject]@{
            Id = 'W1'; Prerequisites = @()
            Title = 'compile and public surface'
            Purpose = ('the complete current VBAProject compiles in real Excel, and the ' +
                       'expected public command surface is present')
        },
        [pscustomobject]@{
            Id = 'W2'; Prerequisites = @()
            Title = 'many DriverFactors instances'
            Purpose = ('300 drivers over 5 project years: many UDT instances each carrying ' +
                       'dynamic arrays, compared against the accepted Phase-5 oracle')
        },
        [pscustomobject]@{
            Id = 'W3'; Prerequisites = @()
            Title = 'maximum year-array length'
            Purpose = ('10 drivers over 200 project years: dynamic arrays inside the UDT at ' +
                       'the structural maximum, compared against the accepted Phase-5 oracle')
        },
        [pscustomobject]@{
            Id = 'W4'; Prerequisites = @()
            Title = 'base simulation and the no-run refusal'
            Purpose = ('the annual endpoint refuses with no successful simulation, then one ' +
                       'deterministic FIXED-seed run the annual scenarios bind to')
        },
        [pscustomobject]@{
            Id = 'W5'; Prerequisites = @('W4')
            Title = 'annual current-run success'
            Purpose = 'the annual answer published and bound to the current run'
        },
        [pscustomobject]@{
            Id = 'W6'; Prerequisites = @('W4', 'W5')
            Title = 'selector move without re-simulation'
            Purpose = ('the ladders survive a moved reporting selector and the profile does ' +
                       'not, and is never relabelled')
        },
        [pscustomobject]@{
            Id = 'W7'; Prerequisites = @()
            Title = 'bank alternation and duration shrink'
            Purpose = 'A to B and back to A, the reused bank publishing a shorter answer'
        },
        [pscustomobject]@{
            Id = 'W8'; Prerequisites = @('W4', 'W5')
            Title = 'live state refusal: STALE and INVALID'
            Purpose = 'both refusals through accepted input paths, with nothing published'
        }
    )
}

# ===========================================================================
# SEVEN HELPERS COPIED VERBATIM FROM THE ACCEPTED PHASE-7 TIMING HARNESS
# ===========================================================================
# WHY THEY ARE COPIED RATHER THAN DOT-SOURCED. `phase4_functional_test.ps1` and
# `phase7_timing_scenarios.ps1` both define them, and NEITHER is definition-only
# at top level: dot-sourcing either would RUN it. The three files this harness
# does dot-source are definition-only, which is what makes dot-sourcing them
# safe and what keeps the historical Gate-B identity untouched.
#
# They are copied BYTE FOR BYTE from `phase7_timing_scenarios.ps1` so there is
# one behaviour rather than two, and a source control pins them to it.
#
# `Write-RowObject` IS HERE BECAUSE OF A TRANSITIVE DEPENDENCY, and it is worth
# naming: nothing in this file calls it. `Get-Phase5TypedTableBody` does - and
# that function lives in `phase5_gate_b_scenarios.ps1`, which this harness DOES
# dot-source, while its own helper lives in `phase4_functional_test.ps1`, which
# this harness deliberately does not. In the accepted Gate-B runs the Phase-4
# driver dot-sources the scenarios file, so the helper is in scope; reaching the
# scenarios file directly leaves that dependency unmet, and W1 died on it before
# a single check was recorded. The timing harness met the same gap the same way.
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

function Get-Phase7SourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; the timing workbook cannot be attributed to a source revision')
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

function Get-Phase7ModuleIdentities {
    param($Manifest, [string]$PccmRoot, [string]$TempRoot)
    $sourceDir = Join-Path $PccmRoot ([string]$Manifest.vba.source_dir)
    $generatedDir = Join-Path $TempRoot (Split-Path -Leaf ([string]$Manifest.vba.generated_dir))
    $out = @()
    foreach ($module in @($Manifest.vba.modules)) {
        $generated = [bool]$module.generated
        $directory = $sourceDir
        $origin = 'repository ' + [string]$Manifest.vba.source_dir
        if ($generated) {
            $directory = $generatedDir
            $origin = 'disposable build copy'
        }
        $path = Join-Path $directory ([string]$module.name + '.bas')
        $hash = ''
        $problem = ''
        try {
            # The SAME canonicalisation the accepted Phase-6 projection identity
            # uses: line endings normalised, a bare CR refused, a BOM refused.
            $hash = Get-Phase6CanonicalModuleHash -Path $path
        } catch {
            $problem = [string]$_.Exception.Message
        }
        $out += [pscustomobject]@{
            Name = [string]$module.name; Generated = $generated; Origin = $origin
            Path = $path; Hash = $hash; Problem = $problem
        }
    }
    return ,$out
}

# AND FIVE MORE FOR THE SAME TRANSITIVE REASON. None of them is called from
# this file either: the accepted Phase-5 fixture choreography calls them, and
# its own file does not define them. They are copied together rather than one
# at a time because the closure control below proves the set is COMPLETE - a
# helper found by running Windows and failing is a helper found too late.
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

# THE SELECTOR IS TEXT, AND Set-NamedValue WRITES NUMBERS.
#
# `$rng.Value2 = [double]$Value` is right for every control the accepted helper
# was written for and wrong for a confidence level, which is a LABEL. Widening
# that helper to switch on type would put a polymorphic COM assignment at a
# single call site - the exact binding defect Set-SimRawCell exists to avoid,
# where PowerShell binds the site per argument type and the second type silently
# fails. So the text writer is its own procedure with its own single-typed
# assignment.
function Set-P7NamedText {
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

# ===========================================================================
# THE REPORT AND THE CHECK LEDGER
# ===========================================================================
$script:Phase7AcceptanceLines = New-Object System.Collections.ArrayList
$script:Phase7AcceptancePath = ''
$script:Phase7Checks = New-Object System.Collections.ArrayList

function Write-P7Line {
    param([string]$Text = '')
    $null = $script:Phase7AcceptanceLines.Add($Text)
    Write-Host $Text
    # WRITTEN THROUGH ON EVERY LINE, so a run that is stopped or killed still
    # leaves every observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:Phase7AcceptancePath)) {
        try {
            Set-Content -LiteralPath $script:Phase7AcceptancePath `
                -Value ($script:Phase7AcceptanceLines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

# ONE CHECK. `Kind` separates a PREREQUISITE - a step the scenario had to
# perform to reach its subject - from a RESULT, which is the scenario's own
# claim. A failed prerequisite is still a failure of the invocation; it is just
# not evidence about the property under test, and the report says which it was.
function Add-P7Check {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:Phase7Checks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = $(if ($Ok) { 'PASS' } else { 'FAIL' })
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line += ' -- ' + $Detail }
    Write-P7Line $line
    return $Ok
}

function Get-P7FailureCount {
    return @($script:Phase7Checks | Where-Object { -not $_.Ok }).Count
}

function Format-P7Err { param($ErrorRecord) return (Format-Err $ErrorRecord) }

# ===========================================================================
# THE PHASE-7 ANNUAL READERS
# ===========================================================================
# EVERY ADDRESS COMES FROM THE PROJECTION. Not one column letter or row number
# is written in this file: they are read from phase7_acceptance_inspection.json,
# which is projected from sim_contract.yaml by the same build that projects them
# into modSimContract.bas. A contract move therefore moves the harness too.
function Get-P7AnnualStamp {
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

function Get-P7ColumnNumber {
    param([string]$Letters)
    $total = 0
    foreach ($ch in $Letters.ToUpperInvariant().ToCharArray()) {
        $total = ($total * 26) + ([int][char]$ch - 64)
    }
    return $total
}

function Get-P7ColumnLetters {
    param([int]$Number)
    $out = ''
    $n = $Number
    while ($n -gt 0) {
        $remainder = ($n - 1) % 26
        $out = [string][char](65 + $remainder) + $out
        $n = [int](($n - $remainder - 1) / 26)
    }
    return $out
}

# THE WHOLE RECORD BLOCK IN ONE READ. 200 project years is 5,200 cells, and a
# per-cell COM round trip for each would dominate the scenario's runtime without
# observing anything a single Value2 read does not.
function Get-P7AnnualBlock {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$RowCount)
    $annual = $P7.annual_records
    $first = Get-P7ColumnNumber -Letters ([string]$annual.index_columns.$Bank.project_index)
    $last = Get-P7ColumnNumber -Letters ([string]$annual.selected_px_profile_columns.$Bank.pv)
    $firstRow = [int]$annual.first_record_row
    $address = (Get-P7ColumnLetters -Number $first) + [string]$firstRow + ':' +
               (Get-P7ColumnLetters -Number $last) + [string]($firstRow + $RowCount - 1)
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $range = $sheet.Range($address)
        return ,$range.Value2
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(annual block)';  $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';           $sheets = $null }
    }
}

# The zero-based column OFFSETS of every contracted field inside the block,
# derived from the projected letters. Nothing here counts columns by hand.
function Get-P7AnnualOffsets {
    param($P7, [string]$Bank)
    $annual = $P7.annual_records
    $origin = Get-P7ColumnNumber -Letters ([string]$annual.index_columns.$Bank.project_index)
    return [pscustomobject]@{
        ProjectIndex   = 0
        CalendarYear   = (Get-P7ColumnNumber -Letters ([string]$annual.index_columns.$Bank.calendar_year)) - $origin
        NominalFirst   = (Get-P7ColumnNumber -Letters ([string]$annual.quantile_first_column.$Bank.nominal)) - $origin
        PvFirst        = (Get-P7ColumnNumber -Letters ([string]$annual.quantile_first_column.$Bank.pv)) - $origin
        NominalProfile = (Get-P7ColumnNumber -Letters ([string]$annual.selected_px_profile_columns.$Bank.nominal)) - $origin
        PvProfile      = (Get-P7ColumnNumber -Letters ([string]$annual.selected_px_profile_columns.$Bank.pv)) - $origin
        Width          = (Get-P7ColumnNumber -Letters ([string]$annual.selected_px_profile_columns.$Bank.pv)) - $origin + 1
    }
}

# One cell of the block, by year (1-based) and zero-based field offset. The
# Variant array Excel hands back is 1-based in both dimensions.
function Get-P7BlockCell {
    param($Block, [int]$Year, [int]$Offset)
    return $Block[$Year, $Offset + 1]
}

# THE FOUR HANDOFF ACCESSORS, called through Application.Run exactly as Phase 8
# will call them. Their NAMES come from the projection.
function Get-P7Handoff {
    param($Excel, $P7)
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($name in @($P7.handoff.accessors)) {
        $out.Add([string]$name, $Excel.Run([string]$name))
    }
    return $out
}

# ONE INVOCATION OF THE ANNUAL ENDPOINT, and its announcement. Deliberately not
# a helper that throws on refusal: three of these scenarios EXPECT a refusal,
# and throwing on the expected outcome would turn them into harness errors.
function Invoke-P7Annual {
    param($Excel, $P7)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run([string]$P7.command_surface.annual_endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

# ===========================================================================
# THE INVARIANTS EVERY ANNUAL SCENARIO MUST HOLD
# ===========================================================================
# The annual step is not a simulation. These are the published facts that must
# not move across ANY annual invocation, successful or refused: the run
# identity, the AUTO nonce and its durable marker, and the result digest.
#
# TWO ROWS ARE EXCLUDED, AND NOT QUIETLY. `simulation_status` and
# `status_evaluated_at` are DERIVED rows: PCCM_SimulationStatus recomputes the
# status and rewrites both every time it is asked, and the annual endpoint's own
# precondition asks it. Requiring the evaluation TIMESTAMP not to move would be
# requiring the annual step not to check whether it may run.
#
# So they are excluded from the frozen set and checked separately, on the terms
# that actually matter: the STATUS ITSELF must be unchanged, and only the
# timestamp beside it may move. `Add-P7StatusChecks` is that check, and every
# annual scenario calls it.
$script:Phase7DerivedRows = @('simulation_status', 'status_evaluated_at')

function Get-P7RunInvariants {
    param($Workbook, $Inspection)
    $state = Get-Phase6State -Workbook $Workbook -Inspection $Inspection
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in $state['shared'].Keys) {
        if ($script:Phase7DerivedRows -contains [string]$key) { continue }
        $out.Add(('shared.' + $key), $state['shared'][$key])
    }
    foreach ($bank in @($Inspection.publication.bank_labels)) {
        $block = $state[('bank_' + $bank)]
        foreach ($key in $block.Keys) { $out.Add(('bank_' + $bank + '.' + $key), $block[$key]) }
    }
    $out.Add('pending_auto_nonce', $state['pending_auto_nonce'])
    return $out
}

function Add-P7InvariantChecks {
    param([string]$Label, $Before, $After)
    $moved = @()
    foreach ($key in $Before.Keys) {
        if (-not (Test-SimSameValue -A $Before[$key] -B $After[$key])) {
            $moved += ($key + ': ' + (Format-SimValue $Before[$key]) + ' -> ' +
                       (Format-SimValue $After[$key]))
        }
    }
    return (Add-P7Check ($Label + ': no run identity, nonce, pending marker or digest moved') `
        ($moved.Count -eq 0) ($moved -join '; '))
}

# THE TWO DERIVED ROWS, on the terms that matter. The status must be the same
# word before and after; only the evaluation timestamp beside it may move.
function Add-P7StatusChecks {
    param($Workbook, $Inspection, [string]$Label, $BeforeState, $AfterState)
    $before = $BeforeState['shared']['simulation_status']
    $after = $AfterState['shared']['simulation_status']
    return (Add-P7Check ($Label + ': the derived simulation status is unchanged') `
        (Test-SimSameValue -A $before -B $after) `
        ((Format-SimValue $before) + ' -> ' + (Format-SimValue $after)))
}

# The iteration column, before and after: an annual run must not touch one
# published iteration total. Read as ONE range, not N round trips.
function Get-P7IterationTotals {
    param($Workbook, $Inspection, [string]$Bank, [int]$Count)
    $block = $Inspection.sim_data.iteration_records
    $nominal = [string]$block.banks.$Bank.total_nominal
    $pv = [string]$block.banks.$Bank.total_pv
    $firstRow = [int]$block.first_iteration_row
    $lastRow = $firstRow + $Count - 1
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        $range = $sheet.Range($nominal + [string]$firstRow + ':' + $pv + [string]$lastRow)
        return ,$range.Value2
    } finally {
        if ($null -ne $range)  { Release-Transient $range  'Range(iterations)';    $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_SimData)';  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';           $sheets = $null }
    }
}

function Test-P7SameGrid {
    param($A, $B)
    if (($null -eq $A) -or ($null -eq $B)) { return ($null -eq $A) -and ($null -eq $B) }
    if ($A.GetLength(0) -ne $B.GetLength(0)) { return $false }
    if ($A.GetLength(1) -ne $B.GetLength(1)) { return $false }
    for ($row = 1; $row -le $A.GetLength(0); $row++) {
        for ($col = 1; $col -le $A.GetLength(1); $col++) {
            if (-not (Test-SimSameValue -A $A[$row, $col] -B $B[$row, $col])) { return $false }
        }
    }
    return $true
}

# ===========================================================================
# W1 - COMPILE AND PUBLIC SURFACE
# ===========================================================================
# NO WINDOWS RUN HAS EVER COMPILED THIS PROJECT. The Stage-B bootstrap imports
# every module and reopens the workbook to prove they persisted, but importing a
# module is not compiling it: a reserved identifier, a declaration after the
# first executable procedure, or a UDT with a dynamic array the compiler will
# not accept all survive an import and fail the first time a procedure is
# entered. VBProject.Compile is the only thing that answers that question, and
# it answers it for the WHOLE project at once.
# HOW THE COMPILE IS FORCED, AND WHY NOT THE OBVIOUS WAY.
#
# The VBIDE object model exposes no `VBProject.Compile`. The usual substitute is
# the Debug > Compile menu control, `VBE.CommandBars.FindControl(,578).Execute`,
# and this harness REFUSES it: on a compile error that control raises a MODAL
# dialog in a hidden Excel that no one can dismiss, which hangs the run and
# leaves exactly the orphaned Excel the accepted lifecycle policy exists to
# prevent. It also needs the VBE window, which an automated session has not
# shown.
#
# WHAT IS USED INSTEAD IS NOT A WEAKER CHECK. VBA compiles the WHOLE PROJECT
# before it executes a single statement of it, so the first `Application.Run` of
# any procedure is a full-project compile - a reserved identifier in a module
# nothing calls, a declaration after the first executable procedure, a UDT the
# compiler will not accept, all fail it. The difference is only that the failure
# arrives as a trappable COM error on the calling thread instead of as a dialog.
function Invoke-P7Compile {
    param($Excel, $P7)
    try {
        # A published READ accessor: it runs no simulation, writes nothing and
        # consumes no run identity, so forcing the compile through it cannot
        # change the state the later scenarios observe.
        $null = $Excel.Run([string]$P7.command_surface.handoff_accessors[0])
        return ''
    } catch {
        return (Format-P7Err $_)
    }
}

# THE DECLARED PROCEDURES, READ OUT OF THE COMPILED PROJECT. Finding a procedure
# in its declared module is what makes the surface claim about THIS project
# rather than about the source tree the workbook was built from.
function Get-P7ProjectProcedures {
    param($Workbook, [string]$ModuleName)
    $vbproj = $null; $comps = $null; $comp = $null; $code = $null
    try {
        $vbproj = $Workbook.VBProject
        $comps = $vbproj.VBComponents
        $comp = $comps.Item($ModuleName)
        $code = $comp.CodeModule
        $names = New-Object System.Collections.ArrayList
        $line = 1
        $total = [int]$code.CountOfLines
        while ($line -le $total) {
            $name = [string]$code.ProcOfLine($line, 0)
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                if (-not $names.Contains($name)) { $null = $names.Add($name) }
                $body = [int]$code.ProcStartLine($line, 0) + [int]$code.ProcCountLines($line, 0)
                if ($body -gt $line) { $line = $body } else { $line++ }
            } else {
                $line++
            }
        }
        return @($names)
    } finally {
        if ($null -ne $code)   { Release-Transient $code   'CodeModule';   $code   = $null }
        if ($null -ne $comp)   { Release-Transient $comp   'VBComponent';  $comp   = $null }
        if ($null -ne $comps)  { Release-Transient $comps  'VBComponents'; $comps  = $null }
        if ($null -ne $vbproj) { Release-Transient $vbproj 'VBProject';    $vbproj = $null }
    }
}

function Invoke-P7W1 {
    param($Excel, $Workbook, $Manifest, $P7)

    Write-P7Line 'W1 - COMPILE AND PUBLIC SURFACE'
    Write-P7Line '-------------------------------'

    # 1. EVERY DECLARED MODULE IS PRESENT, including the two the annual step
    #    added. A compile that passed because a module never loaded would prove
    #    nothing about that module.
    $vbproj = $null; $comps = $null
    $present = @()
    try {
        $vbproj = $Workbook.VBProject
        $comps = $vbproj.VBComponents
        for ($index = 1; $index -le [int]$comps.Count; $index++) {
            $item = $null
            try { $item = $comps.Item($index); $present += [string]$item.Name }
            finally { if ($null -ne $item) { Release-Transient $item 'VBComponent(enum)'; $item = $null } }
        }
    } finally {
        if ($null -ne $comps)  { Release-Transient $comps  'VBComponents'; $comps  = $null }
        if ($null -ne $vbproj) { Release-Transient $vbproj 'VBProject';    $vbproj = $null }
    }
    $declared = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })
    $missing = @($declared | Where-Object { $present -notcontains $_ })
    $null = Add-P7Check 'every manifest-declared module is loaded' `
        ($missing.Count -eq 0) ('missing: ' + ($missing -join ', '))
    foreach ($required in @('modSimAnnual', 'modSimAnnualRun', 'modSimAnnualStore')) {
        $null = Add-P7Check ('the Phase-7 module ' + $required + ' is loaded') `
            ($present -contains $required)
    }

    # 2. THE COMPILE ITSELF.
    $failure = Invoke-P7Compile -Excel $Excel -P7 $P7
    $compiled = [string]::IsNullOrWhiteSpace($failure)
    $null = Add-P7Check 'the complete VBAProject compiles' $compiled $failure
    if (-not $compiled) {
        Write-P7Line ''
        Write-P7Line 'STOP. A compile failure invalidates every behavioural scenario, so W2-W8'
        Write-P7Line 'must not be run until it is fixed.'
        return
    }

    # 3. THE PUBLIC COMMAND SURFACE, module by module. Compiling proves the
    #    project is well formed; this proves it is the surface Phase 8 and the
    #    later scenarios will call.
    $expected = @{
        'modSimAnnualRun'   = @([string]$P7.command_surface.annual_endpoint)
        'modSimAnnualStore' = @($P7.command_surface.handoff_accessors | ForEach-Object { [string]$_ })
        'modSimPostReport'  = @('PCCM_RunSensitivity')
        'modSimReport'      = @('PCCM_RunSimulation', 'PCCM_SimulationStatus')
    }
    foreach ($moduleName in ($expected.Keys | Sort-Object)) {
        $procedures = @()
        $problem = ''
        try { $procedures = @(Get-P7ProjectProcedures -Workbook $Workbook -ModuleName $moduleName) }
        catch { $problem = (Format-P7Err $_) }
        foreach ($wanted in $expected[$moduleName]) {
            $null = Add-P7Check ($moduleName + '.' + $wanted + ' exists in the compiled project') `
                (($procedures -contains $wanted) -and [string]::IsNullOrWhiteSpace($problem)) $problem
        }
    }

    # 4. AND THE ENDPOINTS ARE REACHABLE THROUGH Application.Run, which is how
    #    every later scenario and Phase 8 will reach them. A procedure that
    #    exists but is not callable by name would fail here and nowhere else.
    foreach ($accessor in @($P7.command_surface.handoff_accessors)) {
        $value = $null
        $problem = ''
        try { $value = $Excel.Run([string]$accessor) } catch { $problem = (Format-P7Err $_) }
        $null = Add-P7Check ($accessor + ' is callable and answers on an unrun workbook') `
            ([string]::IsNullOrWhiteSpace($problem)) `
            ('returned ' + (Format-SimValue $value) + $problem)
    }
    # THE HANDOFF ANSWERS "NOT PRODUCED" BEFORE ANYTHING HAS RUN, which is the
    # first state of the settlement and the cheapest place to observe it.
    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $notProduced = [string]$P7.handoff.distribution_states[0]
    $null = Add-P7Check 'the annual distribution state on a fresh workbook is NOT PRODUCED' `
        (([string]$handoff['PCCM_AnnualDistributionState']) -ceq $notProduced) `
        ('returned ' + (Format-SimValue $handoff['PCCM_AnnualDistributionState']))
    $null = Add-P7Check 'the annual profile state on a fresh workbook is NOT PRODUCED' `
        (([string]$handoff['PCCM_AnnualProfileState']) -ceq $notProduced) `
        ('returned ' + (Format-SimValue $handoff['PCCM_AnnualProfileState']))
}

# ===========================================================================
# W2 AND W3 - THE INHERITED PHASE-5 PATH, TWO DIMENSIONS, SEPARATELY
# ===========================================================================
# The Phase-7 change to the accepted Phase-5 path is that `DriverFactors` now
# carries two dynamic arrays of length Y per driver, filled by BuildDriverFactors
# and copied across the accepted bridge. Its two runtime risks are INDEPENDENT -
# many UDT instances each holding arrays, and long arrays inside a UDT - so they
# are proved by two small orthogonal models rather than by one large Cartesian
# one. Nothing about the change couples the dimensions, and a 300 x 200 case
# would cost far more Windows time while proving neither more strongly.
#
# THE EXPECTATION IS PRE-PHASE-7 AUTHORITY: the accepted Phase-5 oracle's answer
# for the SAME payload the fixture is built from, both read from
# build/phase7_acceptance_cases.json. This is not a Phase-7 workbook compared
# against a Phase-7 recalculation.
#
# THE READERS ARE THE ACCEPTED ONES. `Get-CalcTableRows` is the TYPED Phase-5
# reader, so a workbook that published a number as text fails rather than
# passing through a stringifying reader, and `Get-CalcTableColumnIndex` resolves
# every column through the projection rather than by position.
function Compare-P7CalcTable {
    param($Workbook, $Inspection, [string]$TableKey, $Expected, [string[]]$Numeric,
          [string[]]$Exact, [double]$Allowance, [string]$Label)
    $live = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey $TableKey)
    $want = @($Expected)
    if ($live.Count -ne $want.Count) {
        return (Add-P7Check ($Label + ': ' + $TableKey + ' publishes one row per expected row') `
            $false ('published ' + [string]$live.Count + ', expected ' + [string]$want.Count))
    }
    $problems = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $want.Count; $index++) {
        foreach ($column in $Exact) {
            $at = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey $TableKey -ColumnKey $column
            if ($at -lt 0) { $null = $problems.Add($column + ': not a projected column'); continue }
            $got = $live[$index][$at]
            $expect = $want[$index].$column
            if (-not (Test-SimSameValue -A $got -B $expect)) {
                $null = $problems.Add('row ' + [string]($index + 1) + ' ' + $column + ': ' +
                                      (Format-SimValue $got) + ' vs ' + (Format-SimValue $expect))
            }
        }
        foreach ($column in $Numeric) {
            $at = Get-CalcTableColumnIndex -Inspection $Inspection -TableKey $TableKey -ColumnKey $column
            if ($at -lt 0) { $null = $problems.Add($column + ': not a projected column'); continue }
            $got = $live[$index][$at]
            $expect = $want[$index].$column
            if ($null -eq $expect) {
                if ($null -ne $got) {
                    $null = $problems.Add('row ' + [string]($index + 1) + ' ' + $column +
                                          ': published ' + (Format-SimValue $got) + ' where the oracle has none')
                }
                continue
            }
            if ($got -isnot [double]) {
                $null = $problems.Add('row ' + [string]($index + 1) + ' ' + $column + ': ' +
                                      (Format-SimValue $got) + ' is not a number')
                continue
            }
            $difference = [Math]::Abs([double]$got - [double]$expect)
            if ($difference -gt $Allowance) {
                $null = $problems.Add('row ' + [string]($index + 1) + ' ' + $column + ': ' +
                                      [string]$got + ' vs ' + [string]$expect +
                                      ' (difference ' + [string]$difference + ')')
            }
        }
        # A FAILING SCENARIO REPORTS THE FIRST FEW AND THE COUNT, not 6,300
        # lines. The count is what says how bad it is; the examples are what
        # make it diagnosable.
        if ($problems.Count -gt 12) { break }
    }
    $detail = ''
    if ($problems.Count -gt 0) {
        $detail = ($problems -join '; ')
        if ($problems.Count -gt 12) { $detail += ' ... (truncated)' }
    }
    return (Add-P7Check ($Label + ': ' + $TableKey + ' matches the accepted Phase-5 oracle row for row') `
        ($problems.Count -eq 0) $detail)
}

function Compare-P7CalcTotals {
    param($Workbook, $Inspection, $Expected, [double]$Allowance, [string]$Label)
    $problems = @()
    foreach ($key in $Expected.PSObject.Properties.Name) {
        $got = Get-CalcScalar -Workbook $Workbook -Inspection $Inspection `
            -Block 'calc_totals' -FieldKey $key
        $expect = [double]$Expected.$key
        if ($got -isnot [double]) {
            $problems += ($key + ': ' + (Format-SimValue $got) + ' is not a number'); continue
        }
        $difference = [Math]::Abs([double]$got - $expect)
        if ($difference -gt $Allowance) {
            $problems += ($key + ': ' + [string]$got + ' vs ' + [string]$expect +
                          ' (difference ' + [string]$difference + ')')
        }
    }
    return (Add-P7Check ($Label + ': the ten published totals match the accepted Phase-5 oracle') `
        ($problems.Count -eq 0) ($problems -join '; '))
}

# THE SCENARIO ID IS PASSED, NOT READ OFF THE CASE. The Phase-5 harness
# retired reaching into a case object for its id, subtree-wide, after Run 10,
# where a fixture that had no such field was asked for one under StrictMode
# 2.0. The corpus case here does carry one, but the ban is deliberately blunt
# and its reasoning holds: the caller already knows which scenario it
# dispatched, so naming it at the call site is clearer and reaches for nothing.
function Invoke-P7InheritedScenario {
    param($Excel, $Workbook, $Manifest, $Inspection, $Case, $Cases, [string]$Label)

    $label = $Label
    Write-P7Line ($label + ' - ' + [string]$Case.title)
    Write-P7Line ('-' * (5 + ([string]$Case.title).Length))
    $model = $Case.model
    $drivers = @($model.cost_lines).Count + @($model.risks).Count
    Write-P7Line ('  dimension under test : ' + [string]$Case.dimension)
    Write-P7Line ('  drivers              : ' + [string]$drivers)
    Write-P7Line ('  project years        : ' + [string]$model.timeline.duration)
    Write-P7Line ('  expectation source   : ' + [string]$Cases.provenance.expectation_source)
    Write-P7Line ''

    # THE FIXTURE IS THE MODEL FROM THE ARTEFACT, not a second copy of it built
    # here. That is what makes the comparison below mean anything.
    $applied = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
        -Inspection $Inspection -Model $model
    $null = Add-P7Check 'the acceptance model was applied through the accepted fixture' `
        ($applied -like 'OK|*') $applied 'PREREQUISITE'

    $result = ''
    $failed = ''
    try {
        $result = [string](Invoke-Phase5ProductionOperation -Excel $Excel `
            -Operation 'PCCM_Calculate' -Stage ($label + ' calculate'))
    } catch { $failed = (Format-P7Err $_) }
    $null = Add-P7Check 'PCCM_Calculate succeeded on the acceptance model' `
        ([string]::IsNullOrWhiteSpace($failed)) ($result + $failed)
    if (-not [string]::IsNullOrWhiteSpace($failed)) { return }

    $status = [string]$Excel.Run('PCCM_CalculationStatus')
    $null = Add-P7Check 'the calculation reports CURRENT' ($status -ceq 'CURRENT') $status

    # THE ALLOWANCE IS THE PROJECT'S OWN Phase-5 identity floor. No broader
    # tolerance is invented here.
    $allowance = [double]$Cases.provenance.comparison_absolute_floor
    Write-P7Line ('  comparison allowance : ' + [string]$allowance +
                  ' (the accepted Phase-5 identity absolute floor)')

    $null = Compare-P7CalcTotals -Workbook $Workbook -Inspection $Inspection `
        -Expected $Case.expected.totals -Allowance $allowance -Label $label

    # THE YEAR AXIS. This is the array whose LENGTH W3 exists to exercise, and
    # the discount factor is the number the PV arrays inside the UDT carry.
    $null = Compare-P7CalcTable -Workbook $Workbook -Inspection $Inspection `
        -TableKey 'calc_years' -Expected $Case.expected.calc_years `
        -Exact @('project_index', 'calendar_year') -Numeric @('discount_factor') `
        -Allowance $allowance -Label $label

    # THE PER-DRIVER FACTORS. Knom and Kpv are built from the very arrays the UDT
    # now carries, so this is where a copy that lost, shared or truncated an
    # array shows up - one driver at a time, for every driver.
    $null = Compare-P7CalcTable -Workbook $Workbook -Inspection $Inspection `
        -TableKey 'calc_drivers' -Expected $Case.expected.drivers `
        -Exact @('permanent_id', 'currency', 'inflation_profile') `
        -Numeric @('knom', 'kpv', 'mean_value', 'deterministic_nominal', 'deterministic_pv',
                   'mean_basis_nominal', 'mean_basis_pv', 'expected_risk_nominal',
                   'expected_risk_pv') `
        -Allowance $allowance -Label $label

    # AND THE ANNUAL DECOMPOSITION Phase 5 publishes, which is the per-year
    # spread of the same arrays.
    $null = Compare-P7CalcTable -Workbook $Workbook -Inspection $Inspection `
        -TableKey 'calc_annual' -Expected $Case.expected.annual `
        -Exact @('project_index', 'calendar_year') `
        -Numeric @('base_cost_nominal', 'expected_risk_nominal', 'total_nominal',
                   'base_cost_pv', 'expected_risk_pv', 'total_pv') `
        -Allowance $allowance -Label $label
}

# ===========================================================================
# THE BEHAVIOURAL FIXTURE, SHARED BY W4, W5, W6 AND W8
# ===========================================================================
# Small, deterministic and FIXED-seed, so every annual claim binds to one
# identity a reader can name. Established through the accepted fixture and the
# accepted Setup controls - the harness writes no register row of its own.
function Set-P7BehaviouralModel {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $Case)
    $applied = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
        -Inspection $Inspection -Model $Case.model
    Set-NamedValue -Workbook $Workbook `
        -DefinedName ([string]$SimInspection.controls.monte_carlo_iterations.defined_name) `
        -Value ([double]$Case.iterations)
    # A NAMED SEED IS THE FIXED REQUEST. FIXED is used deliberately: the
    # effective seed is then a number the report can state and a later scenario
    # can require to be unchanged, and no AUTO nonce is consumed by the fixture
    # itself.
    Set-NamedValue -Workbook $Workbook `
        -DefinedName ([string]$SimInspection.controls.random_seed.defined_name) `
        -Value ([double]$Case.supplied_seed)
    Set-P7NamedText -Workbook $Workbook `
        -DefinedName ([string]$Inspection.inputs.selected_confidence_level.defined_name) `
        -Value ([string]$Case.selected_confidence_level)
    return $applied
}

# THE PREREQUISITE CHAIN, PERFORMED AND REPORTED AS SUCH.
# A scenario that needs a successful run establishes it here, through the same
# production endpoints, and every step is recorded as a PREREQUISITE so a reader
# never mistakes the setup for the subject.
function Initialize-P7Behavioural {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $Case)
    $applied = Set-P7BehaviouralModel -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
        -Inspection $Inspection -SimInspection $SimInspection -Case $Case
    $null = Add-P7Check 'W4 prerequisite: the behavioural model was applied' `
        ($applied -like 'OK|*') $applied 'PREREQUISITE'
    $calc = ''
    try {
        $calc = [string](Invoke-Phase5ProductionOperation -Excel $Excel `
            -Operation 'PCCM_Calculate' -Stage 'behavioural fixture')
    } catch { $calc = (Format-P7Err $_) }
    $null = Add-P7Check 'W4 prerequisite: PCCM_Calculate succeeded' `
        ($calc -like 'OK|*') $calc 'PREREQUISITE'
    return ($calc -like 'OK|*')
}

function Invoke-P7RunSimulation {
    param($Excel, [string]$Label, [string]$Kind = 'PREREQUISITE')
    $result = Invoke-Phase6Simulation -Excel $Excel
    $null = Add-P7Check ($Label + ': PCCM_RunSimulation succeeded') `
        (Test-Phase6Announced -Result $result -Kind 'OK') $result $Kind
    return $result
}

# ===========================================================================
# W4 - THE NO-RUN REFUSAL, THEN THE BASE SIMULATION
# ===========================================================================
# THE REFUSAL COMES FIRST AND IS NOT A SEPARATE WINDOWS SCENARIO. The state it
# needs - a calculated model with no successful simulation - exists exactly once
# per session, immediately before the first run, so asking for it as its own
# scenario would mean building the whole fixture twice to observe the same
# thing.
function Invoke-P7W4 {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $P7, $Case)

    Write-P7Line 'W4 - THE NO-RUN REFUSAL, THEN THE BASE SIMULATION'
    Write-P7Line '-------------------------------------------------'
    if (-not (Initialize-P7Behavioural -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
            -Inspection $Inspection -SimInspection $SimInspection -Case $Case)) { return $false }

    # --- PART A: no successful simulation exists ----------------------------
    $before = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection
    $status = [string]$Excel.Run('PCCM_SimulationStatus')
    $null = Add-P7Check 'no simulation has been published yet' `
        ([string]::IsNullOrEmpty($status)) ('PCCM_SimulationStatus returned ' + (Format-SimValue $status))

    $announcement = Invoke-P7Annual -Excel $Excel -P7 $P7
    $null = Add-P7Check 'the annual endpoint REFUSES with no successful simulation' `
        ($announcement -like 'FAIL|*') $announcement

    $after = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection
    $null = Add-P7InvariantChecks 'the refused annual run' $before $after

    foreach ($bank in @($SimInspection.publication.bank_labels)) {
        $stamp = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bank
        $blank = @()
        foreach ($key in $stamp.Keys) {
            if (-not (Test-SimBlank -Value $stamp[$key])) {
                $blank += ($key + ' = ' + (Format-SimValue $stamp[$key]))
            }
        }
        $null = Add-P7Check ('the refused run published no annual stamp in bank ' + $bank) `
            ($blank.Count -eq 0) ($blank -join '; ')
    }
    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $notProduced = [string]$P7.handoff.distribution_states[0]
    $null = Add-P7Check 'the annual distribution state is still NOT PRODUCED' `
        (([string]$handoff['PCCM_AnnualDistributionState']) -ceq $notProduced) `
        (Format-SimValue $handoff['PCCM_AnnualDistributionState'])
    $null = Add-P7Check 'no fake year count is reported' `
        (0 -eq [int]$handoff['PCCM_AnnualYearCount']) `
        (Format-SimValue $handoff['PCCM_AnnualYearCount'])

    # --- PART B: the deterministic run the annual scenarios bind to ---------
    Write-P7Line ''
    $result = Invoke-P7RunSimulation -Excel $Excel -Label 'the base simulation' -Kind 'RESULT'
    if (-not (Test-Phase6Announced -Result $result -Kind 'OK')) { return $false }

    $status = [string]$Excel.Run('PCCM_SimulationStatus')
    $null = Add-P7Check 'the simulation reports CURRENT' ($status -ceq 'CURRENT') $status

    $state = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bank = Get-Phase6ActiveBank -State $state
    $null = Add-P7Check 'a publication bank is active' (-not [string]::IsNullOrEmpty($bank)) $bank
    if ([string]::IsNullOrEmpty($bank)) { return $false }

    $seed = $state[('bank_' + $bank)]['effective_seed']
    $null = Add-P7Check 'the FIXED effective seed is the requested one' `
        (Test-SimExactDouble -Value $seed -Expected ([double]$Case.supplied_seed)) `
        ('published ' + (Format-SimValue $seed) + ', requested ' + [string]$Case.supplied_seed)
    $iterations = $state[('bank_' + $bank)]['iterations_run']
    $null = Add-P7Check 'the requested iteration count was run' `
        (Test-SimExactDouble -Value $iterations -Expected ([double]$Case.iterations)) `
        (Format-SimValue $iterations)
    foreach ($field in @('request_fingerprint', 'result_digest')) {
        $value = $state[('bank_' + $bank)][$field]
        $null = Add-P7Check ('the run published a ' + $field) `
            (-not (Test-SimBlank -Value $value)) (Format-SimValue $value)
    }
    Write-P7Line ''
    Write-P7Line (Format-Phase6State -State $state -Label '  the published run identity')
    return $true
}

# ===========================================================================
# W5 - THE ANNUAL ANSWER, PUBLISHED AND BOUND TO THE CURRENT RUN
# ===========================================================================
# THE INDEPENDENT CHECK AVAILABLE LIVE. The harness cannot recompute a per-year
# ladder from the sheet - the iteration-level annual values are deliberately not
# persisted - but it CAN recompute the profile's own identity from data it reads
# independently of the annual block: every iteration's annual vector sums to that
# iteration's total, so the profile must sum to the type-7 blend of the two
# published iteration TOTALS at the selected probability. That blend is computed
# here from the iteration column, and compared with the sum of the published
# profile. It shares no code and no cell with the annual block.
function Get-P7Type7Position {
    param([double[]]$Values, [double]$P)
    $sorted = @($Values | Sort-Object)
    $count = $sorted.Count
    $h = [double]($count - 1) * $P
    $low = [int][Math]::Floor($h)
    $high = $low + 1
    if ($high -gt ($count - 1)) { $high = $count - 1 }
    $fraction = $h - [double]$low
    return [pscustomobject]@{
        LowValue = [double]$sorted[$low]; HighValue = [double]$sorted[$high]
        Fraction = $fraction
    }
}

function Get-P7Type7Value {
    param([double[]]$Values, [double]$P)
    $position = Get-P7Type7Position -Values $Values -P $P
    if (($position.Fraction -eq 0.0) -or ($position.LowValue -eq $position.HighValue)) {
        return $position.LowValue
    }
    return ((1.0 - $position.Fraction) * $position.LowValue) +
           ($position.Fraction * $position.HighValue)
}

function Get-P7LabelProbability {
    param([string]$Label)
    if ($Label -notmatch '^P(\d+)$') {
        throw ('the confidence level ' + [char]39 + $Label + [char]39 + ' is not a P<number> label')
    }
    return ([double]$Matches[1]) / 100.0
}

# THE PUBLISHED ANNUAL ANSWER, READ BACK AND CHECKED AGAINST WHAT IT MUST BE.
function Test-P7AnnualAnswer {
    param($Workbook, $Inspection, $P7, [string]$Bank, [int]$YearCount, [int]$Iterations,
          [string]$PxLabel, [string]$Label)

    $offsets = Get-P7AnnualOffsets -P7 $P7 -Bank $Bank
    $rungs = [int]$P7.annual_records.quantile_count
    $block = Get-P7AnnualBlock -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank $Bank -RowCount $YearCount

    # 1. THE INDEX COLUMNS. One row per applied project year, in order.
    $problems = @()
    for ($year = 1; $year -le $YearCount; $year++) {
        $index = Get-P7BlockCell -Block $block -Year $year -Offset $offsets.ProjectIndex
        if ($index -isnot [double]) { $problems += ('year ' + [string]$year + ': project index ' + (Format-SimValue $index)) }
        elseif ([int]$index -ne $year) { $problems += ('year ' + [string]$year + ': project index ' + [string]$index) }
        $calendar = Get-P7BlockCell -Block $block -Year $year -Offset $offsets.CalendarYear
        if ($calendar -isnot [double]) { $problems += ('year ' + [string]$year + ': calendar year ' + (Format-SimValue $calendar)) }
    }
    $null = Add-P7Check ($Label + ': one record per applied project year, indexed in order') `
        ($problems.Count -eq 0) ($problems -join '; ')

    # 2. BOTH LADDERS ARE PRESENT AND NON-DECREASING ACROSS THE RUNGS. A
    #    percentile ladder that fell as p rose would not be a ladder, whatever
    #    the individual numbers were.
    foreach ($measure in @(@('nominal', $offsets.NominalFirst), @('pv', $offsets.PvFirst))) {
        $problems = @()
        for ($year = 1; $year -le $YearCount; $year++) {
            $previous = $null
            for ($rung = 0; $rung -lt $rungs; $rung++) {
                $value = Get-P7BlockCell -Block $block -Year $year -Offset ($measure[1] + $rung)
                if ($value -isnot [double]) {
                    $problems += ('year ' + [string]$year + ' rung ' + [string]($rung + 1) + ': ' +
                                  (Format-SimValue $value))
                    continue
                }
                if (($null -ne $previous) -and ([double]$value -lt [double]$previous)) {
                    $problems += ('year ' + [string]$year + ' rung ' + [string]($rung + 1) +
                                  ': ' + [string]$value + ' < ' + [string]$previous)
                }
                $previous = $value
            }
        }
        $null = Add-P7Check ($Label + ': the ' + $measure[0] +
                             ' annual ladder is complete and non-decreasing in p') `
            ($problems.Count -eq 0) ($problems -join '; ')
    }

    # 3. THE PROFILE RECONCILES TO THE AUTHORITATIVE TOTAL Px, for both measures,
    #    against the blend computed here from the published iteration totals.
    $totals = Get-P7IterationTotals -Workbook $Workbook -Inspection $Inspection `
        -Bank $Bank -Count $Iterations
    $probability = Get-P7LabelProbability -Label $PxLabel
    $measures = @(@('nominal', $offsets.NominalProfile, 1), @('pv', $offsets.PvProfile, 2))
    foreach ($measure in $measures) {
        $column = New-Object 'double[]' $Iterations
        for ($row = 1; $row -le $Iterations; $row++) { $column[$row - 1] = [double]$totals[$row, $measure[2]] }
        $expected = Get-P7Type7Value -Values $column -P $probability
        $sum = 0.0
        $bad = @()
        for ($year = 1; $year -le $YearCount; $year++) {
            $value = Get-P7BlockCell -Block $block -Year $year -Offset $measure[1]
            if ($value -isnot [double]) { $bad += ('year ' + [string]$year + ': ' + (Format-SimValue $value)); continue }
            $sum += [double]$value
        }
        # THE ALLOWANCE IS THE PROJECT'S OWN I3c/I4c SHAPE, conditioned on the
        # magnitude actually being summed. No new tolerance is invented: the
        # floor and the coefficient are the accepted Phase-5 ones.
        $scale = [Math]::Max(1.0, [Math]::Abs($expected) * [double]$YearCount)
        $allowance = [Math]::Max(1e-6, 1e-12 * $scale)
        $difference = [Math]::Abs($sum - $expected)
        $null = Add-P7Check ($Label + ': the ' + $measure[0] + ' ' + $PxLabel +
                             ' profile sums to the authoritative total Px') `
            (($bad.Count -eq 0) -and ($difference -le $allowance)) `
            ('sum ' + [string]$sum + ', total Px ' + [string]$expected +
             ', difference ' + [string]$difference + ', allowance ' + [string]$allowance +
             $(if ($bad.Count -gt 0) { '; ' + ($bad -join '; ') } else { '' }))
    }

    # 4. NOTHING BEYOND THE ANSWER IS AUTHORITATIVE. The row after the last
    #    project year must be blank, so a longer previous answer cannot be read
    #    as part of this one.
    $tail = Get-P7AnnualBlock -Workbook $Workbook -Inspection $Inspection -P7 $P7 `
        -Bank $Bank -RowCount ($YearCount + 1)
    $spill = @()
    for ($offset = 0; $offset -lt $offsets.Width; $offset++) {
        $value = Get-P7BlockCell -Block $tail -Year ($YearCount + 1) -Offset $offset
        if (-not (Test-SimBlank -Value $value)) {
            $spill += ('offset ' + [string]$offset + ' = ' + (Format-SimValue $value))
        }
    }
    $null = Add-P7Check ($Label + ': the row after the last project year is blank') `
        ($spill.Count -eq 0) ($spill -join '; ')

    return $block
}

function Test-P7AnnualStampBinding {
    param($Workbook, $Inspection, $P7, [string]$Bank, $State, [int]$YearCount,
          [string]$PxLabel, [string]$Label)
    $stamp = Get-P7AnnualStamp -Workbook $Workbook -Inspection $Inspection -P7 $P7 -Bank $Bank
    $published = [string]$P7.annual_records.stamp.published_marker
    $null = Add-P7Check ($Label + ': the annual block is marked published') `
        (([string]$stamp['published']) -ceq $published) (Format-SimValue $stamp['published'])
    # EXACT IDENTITY BINDING: the stamp names the run, not merely a run.
    foreach ($pair in @(@('run_id', 'run_id'), @('effective_seed', 'effective_seed'),
                        @('request_fingerprint', 'request_fingerprint'),
                        @('result_digest', 'result_digest'),
                        @('iterations', 'iterations_run'))) {
        $mine = $stamp[$pair[0]]
        $theirs = $State[('bank_' + $Bank)][$pair[1]]
        $null = Add-P7Check ($Label + ': the annual stamp binds ' + $pair[0] + ' to the published run') `
            (Test-SimSameValue -A $mine -B $theirs) `
            ((Format-SimValue $mine) + ' vs ' + (Format-SimValue $theirs))
    }
    $null = Add-P7Check ($Label + ': the stamped year count is the applied project year count') `
        (Test-SimExactDouble -Value $stamp['year_count'] -Expected ([double]$YearCount)) `
        (Format-SimValue $stamp['year_count'])
    $null = Add-P7Check ($Label + ': the stamped selected Px is the resolved selector') `
        (Test-SimExactText -Value $stamp['selected_px_label'] -Expected $PxLabel) `
        (Format-SimValue $stamp['selected_px_label'])
    $null = Add-P7Check ($Label + ': the stamped probability agrees with the stamped label') `
        (Test-SimExactDouble -Value $stamp['selected_px_probability'] `
            -Expected (Get-P7LabelProbability -Label ([string]$stamp['selected_px_label']))) `
        (Format-SimValue $stamp['selected_px_probability'])
    return $stamp
}

function Invoke-P7W5 {
    param($Excel, $Workbook, $Inspection, $SimInspection, $P7, $Case)

    Write-P7Line ''
    Write-P7Line 'W5 - THE ANNUAL ANSWER, PUBLISHED AND BOUND TO THE CURRENT RUN'
    Write-P7Line '--------------------------------------------------------------'

    $beforeState = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bank = Get-Phase6ActiveBank -State $beforeState
    $iterations = [int]$beforeState[('bank_' + $bank)]['iterations_run']
    $before = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection
    $iterationsBefore = Get-P7IterationTotals -Workbook $Workbook -Inspection $SimInspection `
        -Bank $bank -Count $iterations

    $announcement = Invoke-P7Annual -Excel $Excel -P7 $P7
    $null = Add-P7Check 'PCCM_RunAnnualStochastic succeeded' ($announcement -like 'OK|*') $announcement
    if ($announcement -notlike 'OK|*') { return $false }

    $afterState = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $after = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection
    $null = Add-P7InvariantChecks 'the successful annual run' $before $after
    $null = Add-P7StatusChecks -Workbook $Workbook -Inspection $SimInspection `
        -Label 'the successful annual run' -BeforeState $beforeState -AfterState $afterState

    $iterationsAfter = Get-P7IterationTotals -Workbook $Workbook -Inspection $SimInspection `
        -Bank $bank -Count $iterations
    $null = Add-P7Check 'no published iteration total changed' `
        (Test-P7SameGrid -A $iterationsBefore -B $iterationsAfter)

    $yearCount = [int]$Case.model.timeline.duration
    $px = [string]$Case.selected_confidence_level
    $null = Test-P7AnnualStampBinding -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -State $afterState -YearCount $yearCount -PxLabel $px -Label 'W5'
    $null = Test-P7AnnualAnswer -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -YearCount $yearCount -Iterations $iterations -PxLabel $px -Label 'W5'

    # THE HANDOFF, which is what Phase 8 will actually read.
    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $current = [string]$P7.handoff.distribution_states[1]
    $null = Add-P7Check 'the annual distribution state is CURRENT' `
        (([string]$handoff['PCCM_AnnualDistributionState']) -ceq $current) `
        (Format-SimValue $handoff['PCCM_AnnualDistributionState'])
    $null = Add-P7Check 'the annual profile state is CURRENT' `
        (([string]$handoff['PCCM_AnnualProfileState']) -ceq $current) `
        (Format-SimValue $handoff['PCCM_AnnualProfileState'])
    $null = Add-P7Check 'the handoff names the Px the profile belongs to' `
        (([string]$handoff['PCCM_AnnualProfilePx']) -ceq $px) `
        (Format-SimValue $handoff['PCCM_AnnualProfilePx'])
    $null = Add-P7Check 'the handoff reports the applied year count' `
        ($yearCount -eq [int]$handoff['PCCM_AnnualYearCount']) `
        (Format-SimValue $handoff['PCCM_AnnualYearCount'])
    return $true
}

# ===========================================================================
# W6 - THE SELECTOR MOVES, AND THE SIMULATION DOES NOT RUN AGAIN
# ===========================================================================
# THE MANDATORY PHASE-7 BEHAVIOURAL CLAIM. Selected Confidence Level is a
# REPORTING selector: moving it does not invalidate the simulation, does not
# consume a run and does not require a re-simulation. The annual LADDERS are a
# property of the run alone, so they must survive it byte for byte; the PROFILE
# is the blend at ONE resolved Px, so it must stop being current - and must
# never be relabelled as the newly selected level.
#
# NOTHING HERE RE-RUNS THE SIMULATION. The only production endpoint invoked
# between the two observations is the annual one.
function Invoke-P7W6 {
    param($Excel, $Workbook, $Inspection, $SimInspection, $P7, $Case)

    Write-P7Line ''
    Write-P7Line 'W6 - THE SELECTOR MOVES, AND THE SIMULATION DOES NOT RUN AGAIN'
    Write-P7Line '--------------------------------------------------------------'

    $beforeState = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bank = Get-Phase6ActiveBank -State $beforeState
    $yearCount = [int]$Case.model.timeline.duration
    $iterations = [int]$beforeState[('bank_' + $bank)]['iterations_run']
    $first = [string]$Case.selected_confidence_level
    $second = [string]$Case.second_confidence_level
    Write-P7Line ('  bank                 : ' + $bank)
    Write-P7Line ('  selector before      : ' + $first)
    Write-P7Line ('  selector after       : ' + $second)
    Write-P7Line ''

    # --- 1. RECORD WHAT MUST SURVIVE ----------------------------------------
    $ladderBefore = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -RowCount $yearCount
    $stampBefore = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bank
    $invariantsBefore = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection

    # --- 2. MOVE THE SELECTOR, THROUGH ITS OWN INPUT CELL --------------------
    Set-P7NamedText -Workbook $Workbook `
        -DefinedName ([string]$Inspection.inputs.selected_confidence_level.defined_name) `
        -Value $second
    $readBack = Get-NamedValue -Workbook $Workbook `
        -DefinedName ([string]$Inspection.inputs.selected_confidence_level.defined_name)
    $null = Add-P7Check 'the reporting selector was moved through its own input cell' `
        (([string]$readBack) -ceq $second) (Format-SimValue $readBack) 'PREREQUISITE'

    # --- 3. THE SIMULATION IS UNAFFECTED ------------------------------------
    $status = [string]$Excel.Run('PCCM_SimulationStatus')
    $null = Add-P7Check 'the simulation is still CURRENT after the selector moved' `
        ($status -ceq 'CURRENT') $status
    $movedState = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $null = Add-P7InvariantChecks 'the moved selector' $invariantsBefore `
        (Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection)

    # --- 4. THE LADDERS SURVIVE, THE PROFILE DOES NOT ------------------------
    $ladderAfter = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -RowCount $yearCount
    $null = Add-P7Check 'the persisted annual records are value-identical after the selector moved' `
        (Test-P7SameGrid -A $ladderBefore -B $ladderAfter)

    $stampMoved = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bank
    $null = Add-P7Check 'the stamped Px is still the level the profile was computed at' `
        (Test-SimExactText -Value $stampMoved['selected_px_label'] -Expected $first) `
        (Format-SimValue $stampMoved['selected_px_label'])
    $null = Add-P7Check 'the stamp was not rewritten by the selector move' `
        ((Test-SimSameValue -A $stampBefore['selected_px_label'] -B $stampMoved['selected_px_label']) -and
         (Test-SimSameValue -A $stampBefore['selected_px_probability'] -B $stampMoved['selected_px_probability']) -and
         (Test-SimSameValue -A $stampBefore['published'] -B $stampMoved['published']))

    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $current = [string]$P7.handoff.distribution_states[1]
    $otherPx = [string]$P7.handoff.inconsistent_stamp_state
    $historical = [string]$P7.handoff.distribution_states[2]
    $null = Add-P7Check 'the annual DISTRIBUTION state is still CURRENT' `
        (([string]$handoff['PCCM_AnnualDistributionState']) -ceq $current) `
        (Format-SimValue $handoff['PCCM_AnnualDistributionState'])
    $null = Add-P7Check 'the annual PROFILE state is OTHER Px' `
        (([string]$handoff['PCCM_AnnualProfileState']) -ceq $otherPx) `
        (Format-SimValue $handoff['PCCM_AnnualProfileState'])
    $null = Add-P7Check 'the profile is NOT reported HISTORICAL - the run was not superseded' `
        (([string]$handoff['PCCM_AnnualProfileState']) -cne $historical) `
        (Format-SimValue $handoff['PCCM_AnnualProfileState'])
    $null = Add-P7Check 'the profile is NOT RELABELLED as the newly selected level' `
        (([string]$handoff['PCCM_AnnualProfilePx']) -ceq $first) `
        ('the handoff names ' + (Format-SimValue $handoff['PCCM_AnnualProfilePx']) +
         '; the selector now asks for ' + $second)

    # --- 5. RE-RUN THE ANNUAL STEP, STILL WITHOUT RE-SIMULATING --------------
    Write-P7Line ''
    $announcement = Invoke-P7Annual -Excel $Excel -P7 $P7
    $null = Add-P7Check 'the annual step re-runs against the same simulation' `
        ($announcement -like 'OK|*') $announcement
    if ($announcement -notlike 'OK|*') { return $false }

    $republished = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $null = Add-P7InvariantChecks 'the re-run annual step' $invariantsBefore `
        (Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection)
    $null = Add-P7Check 'the run identity is the same run' `
        ((Test-SimSameValue -A $beforeState[('bank_' + $bank)]['run_id'] `
                            -B $republished[('bank_' + $bank)]['run_id']) -and
         (Test-SimSameValue -A $beforeState[('bank_' + $bank)]['result_digest'] `
                            -B $republished[('bank_' + $bank)]['result_digest']))

    # THE LADDERS ARE THE SAME LADDERS. This is the half of the settlement that
    # says the distributions never depended on the selector: recomputed for a
    # different Px, they must come out identical.
    $ladderRepublished = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -RowCount $yearCount
    $offsets = Get-P7AnnualOffsets -P7 $P7 -Bank $bank
    $rungs = [int]$P7.annual_records.quantile_count
    $ladderProblems = @()
    $profileMoved = $false
    for ($year = 1; $year -le $yearCount; $year++) {
        foreach ($first_offset in @($offsets.NominalFirst, $offsets.PvFirst)) {
            for ($rung = 0; $rung -lt $rungs; $rung++) {
                $was = Get-P7BlockCell -Block $ladderBefore -Year $year -Offset ($first_offset + $rung)
                $now = Get-P7BlockCell -Block $ladderRepublished -Year $year -Offset ($first_offset + $rung)
                if (-not (Test-SimSameValue -A $was -B $now)) {
                    $ladderProblems += ('year ' + [string]$year + ' offset ' +
                                        [string]($first_offset + $rung) + ': ' +
                                        (Format-SimValue $was) + ' -> ' + (Format-SimValue $now))
                }
            }
        }
        foreach ($profileOffset in @($offsets.NominalProfile, $offsets.PvProfile)) {
            $was = Get-P7BlockCell -Block $ladderBefore -Year $year -Offset $profileOffset
            $now = Get-P7BlockCell -Block $ladderRepublished -Year $year -Offset $profileOffset
            if (-not (Test-SimSameValue -A $was -B $now)) { $profileMoved = $true }
        }
    }
    $null = Add-P7Check 'every annual ladder value is unchanged by republication at a different Px' `
        ($ladderProblems.Count -eq 0) ($ladderProblems -join '; ')
    # AND THE PROFILE DID MOVE. A profile identical at two different confidence
    # levels would mean the selector reached nothing, which would make the whole
    # scenario vacuous.
    $null = Add-P7Check 'the selected-Px profile DID change with the selected level' $profileMoved `
        ('a profile identical at ' + $first + ' and ' + $second +
         ' would make this scenario vacuous')

    $null = Test-P7AnnualStampBinding -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -State $republished -YearCount $yearCount -PxLabel $second -Label 'W6'
    $null = Test-P7AnnualAnswer -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -YearCount $yearCount -Iterations $iterations -PxLabel $second -Label 'W6'

    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $null = Add-P7Check 'the profile state is CURRENT for the newly selected level' `
        (([string]$handoff['PCCM_AnnualProfileState']) -ceq $current) `
        (Format-SimValue $handoff['PCCM_AnnualProfileState'])
    $null = Add-P7Check 'the handoff now names the newly selected level' `
        (([string]$handoff['PCCM_AnnualProfilePx']) -ceq $second) `
        (Format-SimValue $handoff['PCCM_AnnualProfilePx'])
    return $true
}

# ===========================================================================
# W7 - A -> B -> A, AND THE REUSED BANK PUBLISHES A SHORTER ANSWER
# ===========================================================================
# ONE SEQUENCE, THREE PROPERTIES. Bank isolation, bank REUSE, and duration
# shrink ON A BANK THAT WAS ALREADY POPULATED. A simple A->B test would prove
# isolation and nothing about what happens when a bank comes round again, and a
# shrink test on a fresh bank would prove nothing at all - the surplus rows it
# claims to clear would never have been written.
#
#   A1  20 project years, simulated and decomposed  -> bank A holds 20 years
#   B   another run                                 -> bank B becomes active
#   A2  4 project years                             -> bank A comes round again
#
# After A2 the old A1 years 5..20 must be gone, the stamped count must be 4, and
# bank B's answer must be exactly where B left it.
function Invoke-P7W7 {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $P7, $Case)

    Write-P7Line 'W7 - BANK ALTERNATION AND DURATION SHRINK'
    Write-P7Line '-----------------------------------------'
    $longYears = [int]$Case.model.timeline.duration
    $shortYears = [int]$Case.shrink_model.timeline.duration
    if ($shortYears -ge $longYears) {
        $null = Add-P7Check 'the shrink model is shorter than the model it replaces' $false `
            ([string]$shortYears + ' vs ' + [string]$longYears)
        return $false
    }
    Write-P7Line ('  run A1 : ' + [string]$longYears + ' project years')
    Write-P7Line ('  run B  : ' + [string]$longYears + ' project years, a second run')
    Write-P7Line ('  run A2 : ' + [string]$shortYears + ' project years, on the reused bank')
    Write-P7Line ''

    # --- A1 ------------------------------------------------------------------
    $applied = Set-P7BehaviouralModel -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
        -Inspection $Inspection -SimInspection $SimInspection -Case $Case
    $null = Add-P7Check 'A1 prerequisite: the 20-year model was applied' `
        ($applied -like 'OK|*') $applied 'PREREQUISITE'
    $calc = [string](Invoke-Phase5ProductionOperation -Excel $Excel -Operation 'PCCM_Calculate' -Stage 'W7 A1')
    $null = Add-P7Check 'A1 prerequisite: PCCM_Calculate succeeded' ($calc -like 'OK|*') $calc 'PREREQUISITE'
    $null = Invoke-P7RunSimulation -Excel $Excel -Label 'A1'
    $null = Add-P7Check 'A1 prerequisite: the annual step published' `
        ((Invoke-P7Annual -Excel $Excel -P7 $P7) -like 'OK|*') '' 'PREREQUISITE'

    $stateA1 = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bankA1 = Get-Phase6ActiveBank -State $stateA1
    $blockA1 = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bankA1 -RowCount $longYears
    $stampA1 = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bankA1
    $null = Add-P7Check ('A1 published ' + [string]$longYears + ' project years into bank ' + $bankA1) `
        (Test-SimExactDouble -Value $stampA1['year_count'] -Expected ([double]$longYears)) `
        (Format-SimValue $stampA1['year_count'])

    # --- B -------------------------------------------------------------------
    Write-P7Line ''
    $null = Invoke-P7RunSimulation -Excel $Excel -Label 'B'
    $null = Add-P7Check 'B prerequisite: the annual step published' `
        ((Invoke-P7Annual -Excel $Excel -P7 $P7) -like 'OK|*') '' 'PREREQUISITE'
    $stateB = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bankB = Get-Phase6ActiveBank -State $stateB
    $null = Add-P7Check 'the second run moved the active bank' ($bankB -cne $bankA1) `
        ($bankA1 + ' -> ' + $bankB)

    # BANK A IS EXACTLY WHERE A1 LEFT IT. Publishing into B may not read, clear
    # or write one cell of A.
    $blockAfterB = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bankA1 -RowCount $longYears
    $null = Add-P7Check ('publishing into bank ' + $bankB + ' did not disturb bank ' + $bankA1) `
        (Test-P7SameGrid -A $blockA1 -B $blockAfterB)
    $stampAfterB = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bankA1
    $moved = @()
    foreach ($key in $stampA1.Keys) {
        if (-not (Test-SimSameValue -A $stampA1[$key] -B $stampAfterB[$key])) {
            $moved += ($key + ': ' + (Format-SimValue $stampA1[$key]) + ' -> ' +
                       (Format-SimValue $stampAfterB[$key]))
        }
    }
    $null = Add-P7Check ('bank ' + $bankA1 + "'s annual stamp is untouched") ($moved.Count -eq 0) ($moved -join '; ')
    $blockB = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bankB -RowCount $longYears
    $stampB = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bankB

    # --- A2: THE BANK COMES ROUND AGAIN, SHORTER -----------------------------
    Write-P7Line ''
    $applied = Set-P7BehaviouralModel -Excel $Excel -Workbook $Workbook -Manifest $Manifest `
        -Inspection $Inspection -SimInspection $SimInspection `
        -Case ([pscustomobject]@{
            model = $Case.shrink_model; iterations = $Case.iterations
            supplied_seed = $Case.supplied_seed
            selected_confidence_level = $Case.selected_confidence_level })
    $null = Add-P7Check 'A2 prerequisite: the 4-year model was applied' `
        ($applied -like 'OK|*') $applied 'PREREQUISITE'
    $calc = [string](Invoke-Phase5ProductionOperation -Excel $Excel -Operation 'PCCM_Calculate' -Stage 'W7 A2')
    $null = Add-P7Check 'A2 prerequisite: PCCM_Calculate succeeded' ($calc -like 'OK|*') $calc 'PREREQUISITE'
    $null = Invoke-P7RunSimulation -Excel $Excel -Label 'A2'
    $stateA2 = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bankA2 = Get-Phase6ActiveBank -State $stateA2
    $null = Add-P7Check 'the third run returned to the first bank' ($bankA2 -ceq $bankA1) `
        ($bankA1 + ' -> ' + $bankB + ' -> ' + $bankA2)
    $null = Add-P7Check 'A2: the annual step published on the reused bank' `
        ((Invoke-P7Annual -Excel $Excel -P7 $P7) -like 'OK|*')

    # 1. THE STAMP BELONGS TO A2, NOT A1.
    $stampA2 = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bankA2
    $null = Add-P7Check 'the reused bank stamp belongs to the new run, not the old one' `
        ((-not (Test-SimSameValue -A $stampA1['run_id'] -B $stampA2['run_id'])) -and
         (Test-SimSameValue -A $stampA2['run_id'] -B $stateA2[('bank_' + $bankA2)]['run_id'])) `
        ('A1 run_id ' + (Format-SimValue $stampA1['run_id']) + ', now ' +
         (Format-SimValue $stampA2['run_id']))
    $null = Add-P7Check ('the authoritative year count is ' + [string]$shortYears) `
        (Test-SimExactDouble -Value $stampA2['year_count'] -Expected ([double]$shortYears)) `
        (Format-SimValue $stampA2['year_count'])
    $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
    $null = Add-P7Check 'the handoff reports the shorter year count' `
        ($shortYears -eq [int]$handoff['PCCM_AnnualYearCount']) `
        (Format-SimValue $handoff['PCCM_AnnualYearCount'])

    # 2. THE SURPLUS ROWS ARE GONE. Read PAST the new answer, all the way to the
    #    old one's last year, and require every cell of every surplus row blank.
    $offsets = Get-P7AnnualOffsets -P7 $P7 -Bank $bankA2
    $tail = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bankA2 -RowCount $longYears
    $residue = @()
    for ($year = $shortYears + 1; $year -le $longYears; $year++) {
        for ($offset = 0; $offset -lt $offsets.Width; $offset++) {
            $value = Get-P7BlockCell -Block $tail -Year $year -Offset $offset
            if (-not (Test-SimBlank -Value $value)) {
                $residue += ('year ' + [string]$year + ' offset ' + [string]$offset + ' = ' +
                             (Format-SimValue $value))
            }
        }
        if ($residue.Count -gt 12) { break }
    }
    $null = Add-P7Check ('years ' + [string]($shortYears + 1) + '..' + [string]$longYears +
                         ' of the previous answer were cleared from the reused bank') `
        ($residue.Count -eq 0) ($residue -join '; ')

    # 3. AND BANK B IS STILL BANK B'S.
    $blockBAfter = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bankB -RowCount $longYears
    $null = Add-P7Check ('republishing into bank ' + $bankA2 + ' did not disturb bank ' + $bankB) `
        (Test-P7SameGrid -A $blockB -B $blockBAfter)
    $stampBAfter = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bankB
    $moved = @()
    foreach ($key in $stampB.Keys) {
        if (-not (Test-SimSameValue -A $stampB[$key] -B $stampBAfter[$key])) {
            $moved += ($key + ': ' + (Format-SimValue $stampB[$key]) + ' -> ' +
                       (Format-SimValue $stampBAfter[$key]))
        }
    }
    $null = Add-P7Check ('bank ' + $bankB + "'s annual stamp is untouched by the bank-" +
                         $bankA2 + ' publication') ($moved.Count -eq 0) ($moved -join '; ')
    return $true
}

# ===========================================================================
# W8 - THE TWO LIVE STATE REFUSALS, THROUGH ACCEPTED INPUT PATHS
# ===========================================================================
# NO HIDDEN-SHEET CORRUPTION. Both states are reached the way a user reaches
# them, and they are reached by genuinely DIFFERENT routes, which is what makes
# them two states rather than one:
#
#   STALE    the SIMULATION request changes while Phase 5 stays CURRENT. The
#            Monte Carlo iteration control is a Phase-6 request input that enters
#            the request fingerprint and no Phase-5 calculation, so moving it
#            makes the published run stale and nothing else.
#   INVALID  the MODEL changes, so Phase 5 is no longer CURRENT, the bridge
#            refuses to hand over a resolved model, and no current request
#            fingerprint can be formed at all.
function Invoke-P7W8 {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection, $P7, $Case)

    Write-P7Line ''
    Write-P7Line 'W8 - THE TWO LIVE STATE REFUSALS'
    Write-P7Line '--------------------------------'
    $yearCount = [int]$Case.model.timeline.duration
    $state = Get-Phase6State -Workbook $Workbook -Inspection $SimInspection
    $bank = Get-Phase6ActiveBank -State $state
    $historicalBlock = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
        -Bank $bank -RowCount $yearCount
    $historicalStamp = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bank
    $iterationsControl = [string]$SimInspection.controls.monte_carlo_iterations.defined_name
    $originalIterations = Get-NamedValue -Workbook $Workbook -DefinedName $iterationsControl

    foreach ($part in @(
        [pscustomobject]@{ Name = 'STALE'; Expect = 'STALE' },
        [pscustomobject]@{ Name = 'INVALID'; Expect = 'INVALID' })) {

        Write-P7Line ''
        Write-P7Line ('  PART ' + $part.Name)
        $invariantsBefore = Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection

        if ($part.Name -eq 'STALE') {
            # The accepted request input, moved by one iteration. Phase 5 is
            # untouched, so the calculation stays CURRENT and only the
            # SIMULATION request has changed.
            Set-NamedValue -Workbook $Workbook -DefinedName $iterationsControl `
                -Value ([double]([int]$Case.iterations + 1))
            $calcStatus = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-P7Check 'STALE was reached without disturbing Phase 5' `
                ($calcStatus -ceq 'CURRENT') ('PCCM_CalculationStatus = ' + $calcStatus) 'PREREQUISITE'
        } else {
            # THE MODEL ITSELF, through the accepted register edit path: one
            # Cost Line's Max value. Phase 5 goes stale, so no current request
            # fingerprint can be formed and the simulation reads INVALID.
            $register = $null
            foreach ($candidate in @($Manifest.registers)) {
                if ([string]$candidate.key -eq 'cost_lines') { $register = $candidate }
            }
            $columns = Get-TableColumnNames -Workbook $Workbook -SheetName ([string]$register.sheet) `
                -TableName ([string]$register.table)
            $at = [array]::IndexOf($columns, 'Max') + 1
            $null = Add-P7Check 'the Cost Line register exposes a Max column to edit' ($at -ge 1) `
                ('columns: ' + ($columns -join ', ')) 'PREREQUISITE'
            if ($at -ge 1) {
                Set-TableCell -Workbook $Workbook -SheetName ([string]$register.sheet) `
                    -TableName ([string]$register.table) -RowIndex 1 -ColumnIndex $at `
                    -Value ([double]999999.0)
            }
            $calcStatus = [string]$Excel.Run('PCCM_CalculationStatus')
            $null = Add-P7Check 'the model edit made Phase 5 non-CURRENT' `
                ($calcStatus -cne 'CURRENT') ('PCCM_CalculationStatus = ' + $calcStatus) 'PREREQUISITE'
        }

        $simStatus = [string]$Excel.Run('PCCM_SimulationStatus')
        $null = Add-P7Check ('the simulation reports ' + $part.Expect) `
            ($simStatus -ceq [string]$part.Expect) ('PCCM_SimulationStatus = ' + $simStatus)

        $announcement = Invoke-P7Annual -Excel $Excel -P7 $P7
        $null = Add-P7Check ('the annual endpoint REFUSES a ' + $part.Expect + ' simulation') `
            ($announcement -like 'FAIL|*') $announcement

        $null = Add-P7InvariantChecks ('the ' + $part.Expect + ' refusal') $invariantsBefore `
            (Get-P7RunInvariants -Workbook $Workbook -Inspection $SimInspection)

        # THE HISTORICAL ANSWER SURVIVES, AND IS NOT PRESENTED AS CURRENT.
        $block = Get-P7AnnualBlock -Workbook $Workbook -Inspection $SimInspection -P7 $P7 `
            -Bank $bank -RowCount $yearCount
        $null = Add-P7Check ('the ' + $part.Expect + ' refusal preserved the historical annual answer') `
            (Test-P7SameGrid -A $historicalBlock -B $block)
        $stamp = Get-P7AnnualStamp -Workbook $Workbook -Inspection $SimInspection -P7 $P7 -Bank $bank
        $moved = @()
        foreach ($key in $historicalStamp.Keys) {
            if (-not (Test-SimSameValue -A $historicalStamp[$key] -B $stamp[$key])) {
                $moved += ($key + ': ' + (Format-SimValue $historicalStamp[$key]) + ' -> ' +
                           (Format-SimValue $stamp[$key]))
            }
        }
        $null = Add-P7Check ('the ' + $part.Expect + ' refusal wrote no annual stamp') `
            ($moved.Count -eq 0) ($moved -join '; ')

        $handoff = Get-P7Handoff -Excel $Excel -P7 $P7
        $historical = [string]$P7.handoff.distribution_states[2]
        $null = Add-P7Check ('the annual answer is reported HISTORICAL, not CURRENT') `
            (([string]$handoff['PCCM_AnnualDistributionState']) -ceq $historical) `
            (Format-SimValue $handoff['PCCM_AnnualDistributionState'])
        $null = Add-P7Check 'the profile is not reported current either' `
            (([string]$handoff['PCCM_AnnualProfileState']) -ceq $historical) `
            (Format-SimValue $handoff['PCCM_AnnualProfileState'])

        if ($part.Name -eq 'STALE') {
            # RESTORED SO PART B STARTS FROM THE SAME PLACE PART A DID, and the
            # restoration is verified rather than assumed.
            Set-NamedValue -Workbook $Workbook -DefinedName $iterationsControl -Value $originalIterations
            $restored = [string]$Excel.Run('PCCM_SimulationStatus')
            $null = Add-P7Check 'restoring the iteration control returns the simulation to CURRENT' `
                ($restored -ceq 'CURRENT') $restored 'PREREQUISITE'
        }
    }
    return $true
}

# ===========================================================================
# PREFLIGHT: THE AUTHORITIES, AND THE SOURCE REVISION
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 WINDOWS ACCEPTANCE' -ForegroundColor Cyan
Write-Host '=================================' -ForegroundColor Cyan
Write-Host ''

$declared = @(Get-Phase7AcceptanceScenarios)
$case = $declared | Where-Object { [string]$_.Id -ceq $Scenario }
if ($null -eq $case) { Write-Host "no scenario '$Scenario'" -ForegroundColor Red; exit 1 }

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$p7InspectPath  = Join-Path $BuildDir 'phase7_acceptance_inspection.json'
$p7CasesPath    = Join-Path $BuildDir 'phase7_acceptance_cases.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $p7InspectPath, $p7CasesPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ("$required not found. Run the Stage-A build first: " +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest      = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json
$p7            = Get-Content -LiteralPath $p7InspectPath  -Raw | ConvertFrom-Json
$p7Cases       = Get-Content -LiteralPath $p7CasesPath    -Raw | ConvertFrom-Json

# THE WORKBOOK MUST BE ATTRIBUTABLE TO A SOURCE REVISION, proved before Excel is
# started. An acceptance result taken from a dirty tree names nothing.
$revision = $null
try { $revision = Get-Phase7SourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so an acceptance result ' +
                'could not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
# The real build output is never opened, never mutated and never saved over.
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-acceptance-' + $Scenario + '-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
foreach ($artefact in @($manifestPath, $inspectPath, $simInspectPath, $p7InspectPath, $p7CasesPath)) {
    Copy-Item -LiteralPath $artefact -Destination $tempRoot
}
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:Phase7AcceptancePath = Join-Path $tempRoot ('phase7_acceptance_' + $Scenario + '.txt')
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
if (($bootstrapExit -ne 0) -or (-not (Test-Path -LiteralPath $stageBPath))) {
    Write-Host ''
    Write-Host ('The Stage-B bootstrap did not complete (exit ' + [string]$bootstrapExit +
                '). Nothing was accepted.') -ForegroundColor Red
    exit 1
}

# The last moment the executed .xlsm is both built and unlocked.
$artefacts = Get-Phase6RuntimeArtefactIdentity -TempRoot $tempRoot -Manifest $manifest
$modules = @(Get-Phase7ModuleIdentities -Manifest $manifest -PccmRoot $pccmRoot -TempRoot $tempRoot)

Write-P7Line 'PCCM - PHASE 7 WINDOWS ACCEPTANCE'
Write-P7Line '================================='
Write-P7Line ''
Write-P7Line 'THIS IS THE PHASE-7 ACCEPTANCE SESSION. It is not Gate B: no Gate-B'
Write-P7Line 'scenario ran and no Gate-B result is recorded or implied. The historical'
Write-P7Line 'Phase-6 runtime authority remains Run 6 on its own closure commit until'
Write-P7Line 'this matrix has been executed and reviewed in full.'
Write-P7Line ''
Write-P7Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-P7Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-P7Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-P7Line ('scenario               : ' + $Scenario + ' - ' + [string]$case.Title)
Write-P7Line ('purpose                : ' + [string]$case.Purpose)
if (@($case.Prerequisites).Count -gt 0) {
    Write-P7Line ('prerequisites          : ' + (@($case.Prerequisites) -join ', ') +
                  ' (established by this invocation and reported as PREREQUISITE)')
} else {
    Write-P7Line  'prerequisites          : none; this scenario is self-contained'
}
Write-P7Line ''
Write-P7Line 'SOURCE REVISION AND BUILD IDENTITY'
Write-P7Line '----------------------------------'
Write-P7Line ('git HEAD               : ' + [string]$revision.Head)
Write-P7Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-P7Line ('model version          : ' + [string]$manifest.model_version)
Write-P7Line ('calc contract version  : ' + [string]$p7Cases.provenance.calc_contract_version)
Write-P7Line ('sim contract version   : ' + [string]$p7.provenance.sim_contract_version)
Write-P7Line ('build directory        : ' + $BuildDir)
Write-P7Line ('working copy           : ' + $tempRoot)
foreach ($item in $artefacts) {
    $shown = $item.Hash
    if ([string]::IsNullOrWhiteSpace($shown)) { $shown = '(' + $item.Problem + ')' }
    Write-P7Line ('  ' + $item.Label.PadRight(30) + ' SHA-256 ' + $shown)
}
Write-P7Line ''
Write-P7Line 'THE VBA THIS RUN EXECUTES (canonicalised SHA-256)'
Write-P7Line '------------------------------------------------'
Write-P7Line 'The generated projection identity is PHASE-7 and is NOT the Phase-6 Run-6'
Write-P7Line 'identity. Contracting the Phase-7 blocks regenerated modSimContract, which'
Write-P7Line 'is correct; the historical controls continue to prove what Run 6 executed'
Write-P7Line 'at its own closure commit, and these lines record what THIS run executes.'
foreach ($module in $modules) {
    $shown = $module.Hash
    if ([string]::IsNullOrWhiteSpace($shown)) { $shown = 'UNREADABLE: ' + $module.Problem }
    Write-P7Line ('  ' + $module.Name.PadRight(20) + ' ' + $module.Origin.PadRight(24) + ' ' + $shown)
}
Write-P7Line ''

# ===========================================================================
# THE ACCEPTANCE SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$rel = $null
$fatal = ''

try {
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    Write-P7Line ('EXCEL PROCESS OWNERSHIP: this run created PID ' +
                  [string]$excelIdentity.ProcessId + '. No process it did not create is ever')
    Write-P7Line 'terminated, and the workbook is never saved.'
    Write-P7Line ''
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)

    # Automation on for the whole session, no failure stage armed: every
    # operation runs its real path and confirmations are answered by the harness
    # rather than by a human.
    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # THE COMPILE CHECK RUNS FIRST IN EVERY SCENARIO. Behavioural evidence taken
    # from a project that does not compile is not evidence.
    if ($Scenario -ne 'W1') {
        $compileFailure = Invoke-P7Compile -Excel $excel -P7 $p7
        $null = Add-P7Check 'the VBAProject compiles (the W1 precondition, re-checked here)' `
            ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
        if (-not [string]::IsNullOrWhiteSpace($compileFailure)) {
            Write-P7Line ''
            Write-P7Line 'STOP. Run W1 and fix the compile failure before any behavioural scenario.'
            throw ('the VBAProject does not compile: ' + $compileFailure)
        }
        Write-P7Line ''
    }

    $behavioural = $p7Cases.scenarios | Where-Object { [string]$_.id -ceq 'W4' }
    $bankCase = $p7Cases.scenarios | Where-Object { [string]$_.id -ceq 'W7' }

    switch ($Scenario) {
        'W1' { Invoke-P7W1 -Excel $excel -Workbook $wb -Manifest $manifest -P7 $p7 }
        'W2' {
            Invoke-P7InheritedScenario -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection `
                -Case ($p7Cases.scenarios | Where-Object { [string]$_.id -ceq 'W2' }) -Cases $p7Cases -Label 'W2'
        }
        'W3' {
            Invoke-P7InheritedScenario -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection `
                -Case ($p7Cases.scenarios | Where-Object { [string]$_.id -ceq 'W3' }) -Cases $p7Cases -Label 'W3'
        }
        'W4' {
            $null = Invoke-P7W4 -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $behavioural
        }
        'W5' {
            if (Invoke-P7W4 -Excel $excel -Workbook $wb -Manifest $manifest `
                    -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $behavioural) {
                $null = Invoke-P7W5 -Excel $excel -Workbook $wb -Inspection $inspection `
                    -SimInspection $simInspection -P7 $p7 -Case $behavioural
            }
        }
        'W6' {
            if (Invoke-P7W4 -Excel $excel -Workbook $wb -Manifest $manifest `
                    -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $behavioural) {
                if (Invoke-P7W5 -Excel $excel -Workbook $wb -Inspection $inspection `
                        -SimInspection $simInspection -P7 $p7 -Case $behavioural) {
                    $null = Invoke-P7W6 -Excel $excel -Workbook $wb -Inspection $inspection `
                        -SimInspection $simInspection -P7 $p7 -Case $behavioural
                }
            }
        }
        'W7' {
            $null = Invoke-P7W7 -Excel $excel -Workbook $wb -Manifest $manifest `
                -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $bankCase
        }
        'W8' {
            if (Invoke-P7W4 -Excel $excel -Workbook $wb -Manifest $manifest `
                    -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $behavioural) {
                if (Invoke-P7W5 -Excel $excel -Workbook $wb -Inspection $inspection `
                        -SimInspection $simInspection -P7 $p7 -Case $behavioural) {
                    $null = Invoke-P7W8 -Excel $excel -Workbook $wb -Manifest $manifest `
                        -Inspection $inspection -SimInspection $simInspection -P7 $p7 -Case $behavioural
                }
            }
        }
    }
} catch {
    $fatal = (Format-Err $_)
    Write-P7Line ''
    Write-P7Line ('THE SCENARIO DID NOT COMPLETE: ' + $fatal)
} finally {
    # THE ACCEPTED SHUTDOWN PATH, unchanged. Workbook.Close without saving,
    # Application.Quit, named releases leaf-before-parent, Wait-ExcelExit, and
    # emergency cleanup ONLY for a process whose identity is still verified.
    # THE LEDGER IS CREATED FIRST AND SEPARATELY, so that a failure while
    # building it cannot skip the wait and the emergency path below.
    $rel = $null
    try { $rel = New-ReleaseLedger 'phase 7 acceptance session' } catch { $rel = $null }
    try {
        if ($null -ne $wb) {
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-NamedRelease $rel $wb        'Workbook';   $wb        = $null
        Invoke-NamedRelease $rel $workbooks 'Workbooks';  $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.Quit = $true } catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-NamedRelease $rel $excel 'Excel.Application'; $excel = $null
    } finally {
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        # THE OUTCOME IS REPORTED ON EVERY PATH, INCLUDING THIS ONE.
        #
        # The W1 run that died on an undefined helper DID reach this block -
        # the exception was raised inside the session try - so the owned
        # process was closed, quit and waited for. What the report could not
        # show is that it happened, because a line was written only when
        # emergency cleanup was needed. Silence read exactly like an orphan.
        # A harness failure must leave evidence of the process's fate.
        $exited = $false
        if ($null -ne $excelIdentity) { $exited = Wait-ExcelExit -Identity $excelIdentity }
        if ($exited) {
            Write-P7Line ('EXCEL SHUTDOWN: the owned process (PID ' +
                          [string]$excelIdentity.ProcessId + ') exited naturally.')
        } else {
            $cleaned = Invoke-EmergencyExcelCleanup -Identity $excelIdentity
            Write-P7Line ('EXCEL SHUTDOWN: emergency cleanup was required (' + [string]$cleaned + ')')
        }
    }
}

# ===========================================================================
# THE VERDICT
# ===========================================================================
$results = @($script:Phase7Checks | Where-Object { [string]$_.Kind -eq 'RESULT' })
$prereqs = @($script:Phase7Checks | Where-Object { [string]$_.Kind -eq 'PREREQUISITE' })
$failedResults = @($results | Where-Object { -not $_.Ok })
$failedPrereqs = @($prereqs | Where-Object { -not $_.Ok })

Write-P7Line ''
Write-P7Line 'VERDICT'
Write-P7Line '-------'
Write-P7Line ('prerequisites          : ' + [string]$prereqs.Count + ' checked, ' +
              [string]$failedPrereqs.Count + ' failed')
Write-P7Line ('scenario results       : ' + [string]$results.Count + ' checked, ' +
              [string]$failedResults.Count + ' failed')
foreach ($failure in @($failedPrereqs + $failedResults)) {
    Write-P7Line ('  FAILED [' + $failure.Kind + '] ' + $failure.Label +
                  $(if ([string]::IsNullOrWhiteSpace($failure.Detail)) { '' } else { ' -- ' + $failure.Detail }))
}
$ok = (($failedResults.Count -eq 0) -and ($failedPrereqs.Count -eq 0) -and
       [string]::IsNullOrWhiteSpace($fatal) -and ($results.Count -gt 0))
Write-P7Line ''
if ($ok) {
    Write-P7Line ($Scenario + ': PASS')
} else {
    Write-P7Line ($Scenario + ': FAIL')
    Write-P7Line ''
    Write-P7Line 'STOP AND REVIEW. Do not run the next scenario until this is understood.'
}
Write-P7Line ''
Write-P7Line ('report                 : ' + $script:Phase7AcceptancePath)

if (-not $KeepArtifacts) {
    Write-Host ''
    Write-Host ('The report is at ' + $script:Phase7AcceptancePath) -ForegroundColor Cyan
    Write-Host 'The working copy is left in place so the report survives; delete it when done.'
}
if ($ok) { exit 0 } else { exit 1 }
