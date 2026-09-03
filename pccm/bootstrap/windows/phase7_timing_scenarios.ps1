<#
.SYNOPSIS
    PCCM Phase 7 - the sensitivity PERFORMANCE MEASUREMENT harness.

.DESCRIPTION
    THIS IS NOT AN ACCEPTANCE HARNESS AND IT PROVES NOTHING ANALYTICAL.

    It exists to answer ONE question that no Linux test can answer: how long
    does the REAL public endpoint `PCCM_RunSensitivity` take, on real Excel,
    against a real successful CURRENT simulation, at three model sizes.

    WHAT IT IS NOT
    --------------
      * It is NOT the Phase-6 Gate-B matrix. It records no Gate-B scenario
        result, touches no Gate-B ledger, and its outcome is never a Gate-B
        pass or fail. `phase4_functional_test.ps1` does not dot-source it.
      * It is NOT a surrogate. Nothing here re-implements mid-ranks, Spearman,
        replay or ranking. It calls the production endpoint and times it.
      * It is NOT instrumentation inside production VBA. No production module
        is modified, and no timing code is added to one. Everything measured is
        measured from PowerShell, around one synchronous Application.Run.
      * It does NOT decide anything. It produces evidence. Whether a measured
        time is acceptable, and whether the architecture needs a cap or
        subsampling, is a decision taken after the evidence exists, not here.

    WHAT IT EXERCISES
    -----------------
    The whole pipeline, through the real public path:

        model built through PCCM_AddCostLine / PCCM_AddRisk
        -> PCCM_Calculate
        -> PCCM_RunSimulation           (persists TotalNom for the run)
        -> PCCM_RunSensitivity          <- THE MEASURED CALL
             replay -> mid-ranks -> per-driver Spearman -> ranking
             -> _SimData persistence -> Sensitivity sheet materialisation

    THE FIXTURE IS THE ACCEPTED ONE. The model is established through
    `Set-Phase5Fixture` from `phase5_gate_b_scenarios.ps1` - the same helper the
    Phase-5 and Phase-6 Gate-B scenarios use - so the drivers are added by
    production's own Add endpoints, into rows production keyed, and every
    identifier is checked against the model before any data is written. This
    file writes no register row, no grid cell and no timeline of its own.

    THE MODEL SHAPE IS CONSTANT ACROSS THE THREE SCENARIOS. Timeline, discount
    rate, currencies, inflation profiles, distribution families, the cost/risk
    split and the profiling weights are identical; ONLY THE DRIVER COUNT
    CHANGES. Every driver has Min < Most Likely < Max, so no driver is
    zero-variance by construction and the ranked table is not mostly refusals.

.PARAMETER BuildDir
    The Stage-A build directory to copy from. Defaults to <repo>/pccm/build.

.PARAMETER Scenario
    A (20 drivers), B (100), C (300), or All. Defaults to A - the smallest
    reliable measurement, and the one the presentation question actually turns
    on.

.PARAMETER SensitivityBudgetSeconds
    THE BOUND, AND WHAT IT HONESTLY IS.

    Application.Run is SYNCHRONOUS and single-threaded. While the endpoint is
    executing, this script is not running, and there is no supported way to
    interrupt a running VBA procedure from the calling thread. A "hard timeout"
    could only be implemented by killing Excel from a second process in the
    middle of a workbook mutation, which is precisely the orphaned/corrupted
    Excel the accepted COM lifecycle policy exists to prevent. This harness
    therefore does NOT claim to interrupt a measurement.

    What it does instead is ENTRY-GATE the next, larger scenario. Scenarios run
    in ascending size, and each one's result is written to the report the
    moment it completes, so the cheap evidence always survives. If a scenario's
    measured sensitivity time exceeds this budget, the remaining scenarios are
    NOT ENTERED and are recorded as NOT ENTERED with the measured reason. That
    refusal is itself performance evidence and is reported as such.

.PARAMETER TotalBudgetSeconds
    The same gate over the whole run, checked before each scenario is entered.

.PARAMETER KeepArtifacts
    Leave the temporary workbook and the report on disk.

.NOTES
    SAFETY. No security setting is altered, no registry key is touched, no
    Trusted Location is added, and no Excel process this script did not create
    is ever terminated. Shutdown is the accepted `com_lifecycle.ps1` path:
    Workbook.Close, Application.Quit, named releases leaf-before-parent,
    Wait-ExcelExit, and Invoke-EmergencyExcelCleanup ONLY for a process whose
    identity is still positively verified. The workbook is never saved, so no
    save cost is included in any measurement and no build output is mutated.
#>

[CmdletBinding()]
param(
    [string]$BuildDir,
    [ValidateSet('A', 'B', 'C', 'All')]
    [string]$Scenario = 'A',
    [int]$SensitivityBudgetSeconds = 600,
    [int]$TotalBudgetSeconds = 3600,
    [switch]$KeepArtifacts
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# THE ACCEPTED FILES ARE REUSED, NOT REIMPLEMENTED.
#
#   com_lifecycle.ps1            the COM ownership and shutdown policy
#   phase5_gate_b_scenarios.ps1  Set-Phase5Fixture and the fixture choreography
#   phase6_gate_b_scenarios.ps1  the _SimData readers and Invoke-Phase6Simulation
#
# All three are definition-only at top level: dot-sourcing them defines
# functions and two script variables and runs no scenario. NOTHING in this file
# calls Invoke-Phase5GateBScenarios or Invoke-Phase6GateBScenarios, so no Gate-B
# result is produced, recorded or implied by a timing run.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

# ===========================================================================
# THE WORKBOOK PRIMITIVES, COPIED VERBATIM FROM THE ACCEPTED DRIVER
# ===========================================================================
# `phase5_gate_b_scenarios.ps1` calls these by name and they are defined at the
# top level of `phase4_functional_test.ps1`, which cannot be dot-sourced here:
# dot-sourcing the driver would run the entire 103-case Gate-B matrix.
#
# THEY ARE COPIES, AND THE COPY IS PINNED. Every function below is byte-identical
# to the driver's, and `tests/test_phase7_timing_harness_source.py` compares them
# line for line. A divergence in either file fails on Linux, so the two cannot
# drift and this file cannot become a second, subtly different reader.
#
# The alternative - extracting them into a shared file and editing
# `phase4_functional_test.ps1` to dot-source it - would change the bytes of the
# accepted Gate-B harness for the sake of a measurement. That is not a trade
# this task is allowed to make.
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
# THE PHASE-7 ADDRESSES
# ===========================================================================
# `phase6_gate_b_inspection.json` is a PINNED artefact whose SHA-256 the
# Phase-6 Gate-B fixture-integrity checks depend on. Extending it to carry the
# Phase-7 sensitivity block would move that hash and break the historical
# Phase-6 identity controls, which this task must not do. So the addresses are
# declared here instead - and they are a CHECKED COPY, never a second
# declaration: `tests/test_phase7_timing_harness_source.py` asserts every value
# below against `pccm/spec/sim_contract.yaml`, `pccm/spec/workbook.yaml` and
# `pccm/src/vba/modSimPostReport.bas`. The contract stays the authority.
#
# THE COLUMNS MOVED, AND THE HARNESS WAS NOT WHAT WAS WRONG. These read J and S
# before the Windows run of 0734a38, because that is what the contract said. The
# stamp came back as 'Mean', 'Sample Standard Deviation', 'Minimum', 'P10',
# 'P50', 'P55', 'P60' - the summary-statistics labels of rows 8 to 14, which is
# what J8:J14 has held since Phase 6. The harness read the contracted cells
# faithfully; the contract had allocated the block on top of two accepted
# blocks. The sensitivity block now lives at CC-CJ and CL-CS, and the values
# below follow it because the tests below read them out of the contract.
$script:Phase7SensitivityGeometry = [pscustomobject]@{
    HeaderRow       = 33
    FirstRecordRow  = 34
    StampColumns    = @{ 'A' = 'CC'; 'B' = 'CL' }
    StatusColumns   = @{ 'A' = 'CJ'; 'B' = 'CS' }
    StampRows       = @{
        run_id              = 8
        effective_seed      = 9
        request_fingerprint = 10
        result_digest       = 11
        iterations          = 12
        record_count        = 13
        published           = 14
    }
    PublishedMarker = 'PUBLISHED'
    RankedLabel     = 'ranked'
    NoVarianceLabel = 'n/a - no variance'
}

$script:Phase7SensitivitySheet = [pscustomobject]@{
    Sheet              = 'Sensitivity'
    AvailabilityColumn = 'B'
    AvailabilityRow    = 10
    HeaderRow          = 12
    FirstRow           = 13
    RowWindow          = 200
    FirstColumn        = 'B'
    LastColumn         = 'I'
}

# ===========================================================================
# THE SCENARIOS
# ===========================================================================
# THREE SIZES OF ONE MODEL. The shape is fixed; the driver count is the only
# variable. 10,000 iterations everywhere - there is deliberately NO 100,000
# scenario in this first measurement.
function Get-Phase7TimingScenarios {
    return @(
        [pscustomobject]@{ Id = 'A'; Title = 'DEMO-LIKE';          DriverCount = 20;  Iterations = 10000 },
        [pscustomobject]@{ Id = 'B'; Title = 'MEDIUM';             DriverCount = 100; Iterations = 10000 },
        [pscustomobject]@{ Id = 'C'; Title = 'DESIGN-SCALE PROBE'; DriverCount = 300; Iterations = 10000 }
    )
}

# THE COST/RISK SPLIT, stated once and used by both the model builder and the
# report. 60% Cost Lines, the remainder Risks: 20 -> 12/8, 100 -> 60/40,
# 300 -> 180/120.
function Get-Phase7CostLineCount {
    param([int]$DriverCount)
    return [int][Math]::Ceiling([double]$DriverCount * 0.6)
}

# ONE DRIVER. Every field varies with the index, and Min < Most Likely < Max
# holds for every single driver, so NO driver is degenerate and the zero-variance
# arm of the analysis is not what is being timed.
function New-Phase7TimingDriver {
    param([int]$Index, [bool]$IsRisk)
    $families = @('Triangular', 'Beta-PERT', 'Uniform')
    $profiles = @('Standard', 'Escalated')
    $currencies = @('SAR', 'USD')

    $base = 100.0 + (7.0 * [double]$Index)
    $driver = [pscustomobject]@{
        permanent_id      = $(if ($IsRisk) { 'R-{0:D3}' -f $Index } else { 'CL-{0:D3}' -f $Index })
        distribution      = [string]$families[($Index - 1) % $families.Count]
        currency          = [string]$currencies[($Index - 1) % $currencies.Count]
        inflation_profile = [string]$profiles[($Index - 1) % $profiles.Count]
        min_value         = $base
        most_likely       = [double]($base * 1.35)
        max_value         = [double]($base * 2.10)
        profile_weights   = @(0.25, 0.50, 0.25)
    }
    if ($IsRisk) {
        # 0.1 .. 0.7, never 0 and never 1: a Risk that always occurs or never
        # occurs would carry no occurrence variance, and this measurement is
        # about the populated path.
        Add-Member -InputObject $driver -MemberType NoteProperty -Name 'probability' `
            -Value ([double](((($Index - 1) % 7) + 1) / 10.0))
    } else {
        Add-Member -InputObject $driver -MemberType NoteProperty -Name 'quantity' `
            -Value ([double](1 + (($Index - 1) % 5)))
    }
    return $driver
}

function New-Phase7TimingModel {
    param([int]$DriverCount)
    if ($DriverCount -lt 2) {
        throw ('a timing model needs at least two drivers; asked for ' + [string]$DriverCount)
    }
    $costCount = Get-Phase7CostLineCount -DriverCount $DriverCount
    $riskCount = $DriverCount - $costCount
    if ($riskCount -lt 1) {
        throw ('the timing model must contain Risks as well as Cost Lines; the split ' +
               'for ' + [string]$DriverCount + ' drivers produced ' + [string]$riskCount + ' Risk(s)')
    }

    $costLines = @()
    for ($index = 1; $index -le $costCount; $index++) {
        $costLines += (New-Phase7TimingDriver -Index $index -IsRisk $false)
    }
    $risks = @()
    for ($index = 1; $index -le $riskCount; $index++) {
        $risks += (New-Phase7TimingDriver -Index $index -IsRisk $true)
    }

    # base_year 2026 with start_year 2027 is the accepted shape: the generated
    # inflation columns begin at BaseYear + 1, so 2027..2029 are exactly the
    # three project years.
    return [pscustomobject]@{
        timeline      = [pscustomobject]@{ base_year = 2026; start_year = 2027; duration = 3 }
        discount_rate = 0.05
        fx            = @(
            [pscustomobject]@{ currency = 'SAR'; rate = 1.0 },
            [pscustomobject]@{ currency = 'USD'; rate = 3.75 }
        )
        inflation     = [pscustomobject]@{
            'Standard'  = [pscustomobject]@{ '2027' = 0.03;  '2028' = 0.03;   '2029' = 0.03 }
            'Escalated' = [pscustomobject]@{ '2027' = 0.06;  '2028' = 0.055;  '2029' = 0.05 }
        }
        cost_lines    = $costLines
        risks         = $risks
    }
}

# ===========================================================================
# THE SENSITIVITY READERS
# ===========================================================================
function Get-Phase7SensitivityStamp {
    param($Workbook, $Inspection, [string]$Bank)
    $geometry = $script:Phase7SensitivityGeometry
    $column = [string]$geometry.StampColumns[$Bank]
    if ([string]::IsNullOrEmpty($column)) {
        throw ('there is no sensitivity stamp column for bank ' + [char]39 + $Bank + [char]39)
    }
    $out = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($key in @('run_id', 'effective_seed', 'request_fingerprint',
                       'result_digest', 'iterations', 'record_count', 'published')) {
        $address = $column + [string]([int]$geometry.StampRows[$key])
        $out.Add($key, (Get-SimRawCell -Workbook $Workbook -Inspection $Inspection -Address $address))
    }
    return $out
}

# The per-record status column, counted. ONE sheet acquisition for the whole
# block; each cell Range is still its own named object with its own release.
function Get-Phase7SensitivityStatusCounts {
    param($Workbook, $Inspection, [string]$Bank, [int]$RecordCount)
    $geometry = $script:Phase7SensitivityGeometry
    $column = [string]$geometry.StatusColumns[$Bank]
    if ([string]::IsNullOrEmpty($column)) {
        throw ('there is no sensitivity status column for bank ' + [char]39 + $Bank + [char]39)
    }
    $counts = [pscustomobject]@{ Ranked = 0; NoVariance = 0; Other = 0; OtherLabels = @() }
    if ($RecordCount -le 0) { return $counts }

    $sheets = $null; $sheet = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$Inspection.sim_data.sheet)
        for ($offset = 0; $offset -lt $RecordCount; $offset++) {
            $range = $null
            try {
                $range = $sheet.Range($column + [string]([int]$geometry.FirstRecordRow + $offset))
                $value = $range.Value2
                $text = ''
                if ($null -ne $value) { $text = [string]$value }
                if ($text -ceq [string]$geometry.RankedLabel) {
                    $counts.Ranked = $counts.Ranked + 1
                } elseif ($text -ceq [string]$geometry.NoVarianceLabel) {
                    $counts.NoVariance = $counts.NoVariance + 1
                } else {
                    $counts.Other = $counts.Other + 1
                    if ($counts.OtherLabels -notcontains $text) { $counts.OtherLabels += $text }
                }
            } finally {
                if ($null -ne $range) { Release-Transient $range 'Range(_SimData)'; $range = $null }
            }
        }
    } finally {
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(_SimData)'; $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';          $sheets = $null }
    }
    return $counts
}

# THE MATERIALISATION, read from the user-facing sheet rather than inferred from
# the block. The sheet is pure lookups over the block, so a populated first row
# and a CURRENT availability line are what "the Sensitivity sheet materialised"
# actually means.
function Get-Phase7SensitivitySheetState {
    param($Workbook)
    $sheetSpec = $script:Phase7SensitivitySheet
    $sheets = $null; $sheet = $null
    $availability = ''
    $populated = 0
    $firstRow = ''
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item([string]$sheetSpec.Sheet)

        $range = $null
        try {
            $range = $sheet.Range([string]$sheetSpec.AvailabilityColumn + [string][int]$sheetSpec.AvailabilityRow)
            $value = $range.Value2
            if ($null -ne $value) { $availability = [string]$value }
        } finally {
            if ($null -ne $range) { Release-Transient $range 'Range(Sensitivity)'; $range = $null }
        }

        for ($offset = 0; $offset -lt [int]$sheetSpec.RowWindow; $offset++) {
            $range = $null
            try {
                $range = $sheet.Range([string]$sheetSpec.FirstColumn + [string]([int]$sheetSpec.FirstRow + $offset))
                $value = $range.Value2
                $text = ''
                if ($null -ne $value) { $text = [string]$value }
                if (-not [string]::IsNullOrWhiteSpace($text)) { $populated = $populated + 1 }
            } finally {
                if ($null -ne $range) { Release-Transient $range 'Range(Sensitivity)'; $range = $null }
            }
        }

        $range = $null
        try {
            $range = $sheet.Range([string]$sheetSpec.FirstColumn + [string][int]$sheetSpec.FirstRow + ':' +
                                  [string]$sheetSpec.LastColumn + [string][int]$sheetSpec.FirstRow)
            $cells = @()
            for ($column = 1; $column -le 8; $column++) {
                $cell = $null
                try {
                    $cell = $range.Cells(1, $column)
                    $value = $cell.Value2
                    if ($null -eq $value) { $cells += '' } else { $cells += [string]$value }
                } finally {
                    if ($null -ne $cell) { Release-Transient $cell 'Range(cell)'; $cell = $null }
                }
            }
            $firstRow = ($cells -join ' | ')
        } finally {
            if ($null -ne $range) { Release-Transient $range 'Range(Sensitivity)'; $range = $null }
        }
    } finally {
        if ($null -ne $sheet)  { Release-Transient $sheet  'Worksheet(Sensitivity)'; $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';             $sheets = $null }
    }
    return [pscustomobject]@{
        Availability  = $availability
        PopulatedRows = $populated
        FirstRow      = $firstRow
    }
}

# ===========================================================================
# THE PROVENANCE OF THE WORKBOOK BEING TIMED
# ===========================================================================
# "The workbook contains the current Phase-7 executable projection" is a claim
# about BYTES, so it is proved with bytes.
#
#   * the hand-written modules are imported by `build_stage_b.ps1` from the
#     REPOSITORY `pccm/src/vba`, so a clean `pccm/src` plus the commit hash
#     identifies them exactly;
#   * the generated modules are imported from the DISPOSABLE BuildDir copy, so
#     those are hashed from the copy that was actually imported;
#   * and the executed .xlsm is hashed at the only moment it is both built and
#     unlocked, through the accepted Phase-6 capture.
#
# FAIL CLOSED. If `pccm/src`, `pccm/spec` or `pccm/builder` is dirty, the
# workbook cannot be attributed to a revision and the run stops before Excel is
# started. A measurement nobody can attribute is not evidence.
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

# ===========================================================================
# THE REPORT
# ===========================================================================
$script:Phase7ReportLines = New-Object System.Collections.ArrayList
$script:Phase7ReportPath = ''

function Write-Phase7Line {
    param([string]$Text = '')
    $null = $script:Phase7ReportLines.Add($Text)
    Write-Host $Text
    # WRITTEN THROUGH ON EVERY LINE. A run that is stopped, killed or times out
    # must still leave every measurement it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:Phase7ReportPath)) {
        try {
            Set-Content -LiteralPath $script:Phase7ReportPath `
                -Value ($script:Phase7ReportLines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Format-Phase7Seconds {
    param([double]$Milliseconds)
    return ('{0:N3} s ({1:N0} ms)' -f ($Milliseconds / 1000.0), $Milliseconds)
}

# ONE INVARIANT, BEFORE AND AFTER, COMPARED WITH THE ACCEPTED COMPARATOR.
# Test-SimSameValue compares Value2 INCLUDING TYPE, so a Double that became the
# text of itself is reported as CHANGED rather than as unchanged.
function Compare-Phase7Invariant {
    param([string]$Label, $Before, $After)
    $same = Test-SimSameValue -A $Before -B $After
    return [pscustomobject]@{
        Label = $Label
        Same  = $same
        Text  = ('    ' + $Label + ': ' + $(if ($same) { 'UNCHANGED' } else { 'CHANGED' }) +
                 ' (before ' + (Format-SimValue $Before) + ', after ' + (Format-SimValue $After) + ')')
    }
}

# ===========================================================================
# PREFLIGHT: THE AUTHORITIES, AND THE SOURCE REVISION
# ===========================================================================
Write-Host ''
Write-Host 'PCCM - Phase 7 sensitivity performance measurement' -ForegroundColor Cyan
Write-Host '==================================================' -ForegroundColor Cyan
Write-Host ''
Write-Host 'This is a MEASUREMENT run. It is not Gate B, it records no Gate-B' -ForegroundColor Yellow
Write-Host 'result, and it proves nothing analytical.' -ForegroundColor Yellow
Write-Host ''

$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ("$required not found. Run the Stage-A build first: " +
                    'python3 pccm/builder/build_stage_a.py') -ForegroundColor Red
        exit 1
    }
}
$manifest      = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json

$revision = $null
try {
    $revision = Get-Phase7SourceRevision -RepoRoot $repoRoot
} catch {
    Write-Host (Format-Err $_) -ForegroundColor Red
    exit 1
}
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so the workbook this ' +
                'run would build cannot be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    Write-Host ''
    Write-Host 'Commit or stash those changes and run again.' -ForegroundColor Red
    exit 1
}

# THE DECLARED SET IS NEVER FILTERED AWAY.
#
# The first real run of this harness reported scenario A and then simply
# stopped. There was no NOT ENTERED line for B or C and no gate reason, because
# `-Scenario A` had removed them from the collection BEFORE the loop ever ran.
# The loop was not defective; the REPORT was. A measurement harness whose whole
# product is a report must never let a scenario disappear without saying why,
# and "out of scope for this run" is a reason exactly as a budget refusal is.
#
# So selection stops being a filter and becomes a REPORTED OUTCOME: the loop
# iterates every DECLARED scenario, always, and an unselected one is skipped
# with its own line and its own summary row. The summary can then account for
# every scenario the harness declares, whatever was asked of it.
$declared = @(Get-Phase7TimingScenarios)
$selectedIds = @($declared | ForEach-Object { [string]$_.Id })
if ($Scenario -ne 'All') { $selectedIds = @([string]$Scenario) }
$selectedCount = @($declared | Where-Object { $selectedIds -contains [string]$_.Id }).Count
if ($selectedCount -eq 0) { Write-Host "no scenario '$Scenario'" -ForegroundColor Red; exit 1 }

# ===========================================================================
# A DISPOSABLE COPY OF THE BUILD, AND THE STAGE-B BOOTSTRAP
# ===========================================================================
# Identical in policy to the accepted driver: the real build output is never
# opened, never mutated and never saved over.
$tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('pccm-phase7-timing-' + (Get-Date).ToString('yyyyMMdd-HHmmss'))
$null = New-Item -ItemType Directory -Path $tempRoot -Force
Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$manifest.stage_a_filename)) -Destination $tempRoot
Copy-Item -LiteralPath $manifestPath   -Destination $tempRoot
Copy-Item -LiteralPath $inspectPath    -Destination $tempRoot
Copy-Item -LiteralPath $simInspectPath -Destination $tempRoot
Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

$script:Phase7ReportPath = Join-Path $tempRoot 'phase7_timing_report.txt'
$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)

$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'
& $bootstrap -BuildDir $tempRoot -Force
$bootstrapExit = $LASTEXITCODE
if (($bootstrapExit -ne 0) -or (-not (Test-Path -LiteralPath $stageBPath))) {
    Write-Host ''
    Write-Host ('The Stage-B bootstrap did not complete (exit ' + [string]$bootstrapExit +
                '). Nothing was measured.') -ForegroundColor Red
    exit 1
}

# The last moment the executed .xlsm is both built and unlocked.
$artefacts = Get-Phase6RuntimeArtefactIdentity -TempRoot $tempRoot -Manifest $manifest
$modules = @(Get-Phase7ModuleIdentities -Manifest $manifest -PccmRoot $pccmRoot -TempRoot $tempRoot)

Write-Phase7Line 'PCCM - PHASE 7 SENSITIVITY PERFORMANCE MEASUREMENT'
Write-Phase7Line '================================================='
Write-Phase7Line ''
Write-Phase7Line 'THIS IS PREPARATION EVIDENCE, NOT PHASE-7 WINDOWS ACCEPTANCE.'
Write-Phase7Line 'No Gate-B scenario ran. No Gate-B result is recorded or implied.'
Write-Phase7Line ''
Write-Phase7Line ('run started            : ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
Write-Phase7Line ('host                   : ' + [string]$env:COMPUTERNAME)
Write-Phase7Line ('PowerShell             : ' + [string]$PSVersionTable.PSVersion)
Write-Phase7Line ''
Write-Phase7Line 'SOURCE REVISION AND BUILD IDENTITY'
Write-Phase7Line '----------------------------------'
Write-Phase7Line ('git HEAD               : ' + [string]$revision.Head)
Write-Phase7Line  'pccm/src, pccm/spec, pccm/builder : clean (proved before Excel was started)'
Write-Phase7Line ('model version          : ' + [string]$manifest.model_version)
Write-Phase7Line ('sim contract version   : ' + [string]$simInspection.provenance.sim_contract_version)
Write-Phase7Line ('build directory        : ' + $BuildDir)
Write-Phase7Line ('working copy           : ' + $tempRoot)
foreach ($item in $artefacts) {
    $shown = $item.Hash
    if ([string]::IsNullOrWhiteSpace($shown)) { $shown = '(' + $item.Problem + ')' }
    Write-Phase7Line ('  ' + $item.Label.PadRight(30) + ' SHA-256 ' + $shown)
}
Write-Phase7Line ''
Write-Phase7Line 'THE VBA THE TIMED WORKBOOK CONTAINS (canonicalised SHA-256)'
foreach ($module in $modules) {
    $shown = $module.Hash
    if ([string]::IsNullOrWhiteSpace($shown)) { $shown = 'UNREADABLE: ' + $module.Problem }
    Write-Phase7Line ('  ' + $module.Name.PadRight(20) + ' ' + $module.Origin.PadRight(24) + ' ' + $shown)
}
Write-Phase7Line ''
Write-Phase7Line 'SCENARIOS DECLARED, AND THE SCOPE OF THIS RUN'
Write-Phase7Line '---------------------------------------------'
Write-Phase7Line ('-Scenario ' + $Scenario)
foreach ($scopeCase in $declared) {
    $scope = 'NOT SELECTED for this run'
    if ($selectedIds -contains [string]$scopeCase.Id) { $scope = 'selected' }
    Write-Phase7Line ('  ' + $scopeCase.Id + '  ' + ([string]$scopeCase.DriverCount).PadLeft(4) +
                      ' drivers, ' + [string]$scopeCase.Iterations + ' iterations  ' + $scope)
}
Write-Phase7Line ''
Write-Phase7Line 'BOUNDS AND SAFETY'
Write-Phase7Line '-----------------'
Write-Phase7Line ('sensitivity budget     : ' + [string]$SensitivityBudgetSeconds + ' s per scenario')
Write-Phase7Line ('total budget           : ' + [string]$TotalBudgetSeconds + ' s for the run')
Write-Phase7Line  'bound mechanism        : ENTRY-GATE, not interruption. Application.Run is'
Write-Phase7Line  '                         synchronous and a running VBA procedure cannot be'
Write-Phase7Line  '                         interrupted from the calling thread; killing Excel'
Write-Phase7Line  '                         mid-mutation is the orphan the COM lifecycle policy'
Write-Phase7Line  '                         exists to prevent. A scenario that exceeds the'
Write-Phase7Line  '                         budget therefore runs to completion and STOPS THE'
Write-Phase7Line  '                         LARGER SCENARIOS FROM BEING ENTERED. That refusal is'
Write-Phase7Line  '                         itself performance evidence.'
Write-Phase7Line  'workbook save          : NEVER. No measurement includes a save cost.'
Write-Phase7Line ''

# ===========================================================================
# THE MEASUREMENT SESSION
# ===========================================================================
$preExisting = @(Get-PreExistingExcelPids)
$excel = $null; $workbooks = $null; $wb = $null
$excelIdentity = $null
$rel = $null
$runStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$measurements = New-Object System.Collections.ArrayList
$gateReason = ''

try {
    $excel = New-Object -ComObject Excel.Application
    $excelIdentity = Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $wb = $workbooks.Open($stageBPath)

    $costRegister = $null; $riskRegister = $null
    foreach ($register in @($manifest.registers)) {
        if ([string]$register.key -eq 'cost_lines')    { $costRegister = $register }
        if ([string]$register.key -eq 'risk_register') { $riskRegister = $register }
    }

    # Automation on for the whole session, no failure stage armed: every
    # operation runs its real path and confirmations are answered by the
    # harness rather than by a human.
    $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null

    # The locked FX seed, captured ONCE on the untouched workbook, exactly as
    # the accepted fixture requires.
    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    foreach ($case in $declared) {
        # SCOPE FIRST, AND IT IS NOT A REFUSAL. An unselected scenario neither
        # spends the budget nor arms the gate - it is reported and skipped, so a
        # scoped run and a gated run can never be mistaken for one another.
        if ($selectedIds -notcontains [string]$case.Id) {
            Write-Phase7Line ('SCENARIO ' + $case.Id + ' - ' + $case.Title + ' (' +
                              [string]$case.DriverCount + ' drivers)')
            Write-Phase7Line ('  NOT SELECTED: this run was scoped to -Scenario ' + $Scenario +
                              '. That is a scope, not a budget refusal and not a failure.')
            Write-Phase7Line ''
            $null = $measurements.Add([pscustomobject]@{
                Id = $case.Id; Selected = $false; Entered = $false
                Reason = ('not selected; this run was scoped to -Scenario ' + $Scenario)
                SensitivityMs = [double]0; DriverCount = [int]$case.DriverCount
            })
            continue
        }
        if ([string]::IsNullOrEmpty($gateReason)) {
            $elapsedTotal = $runStopwatch.Elapsed.TotalSeconds
            if ($elapsedTotal -gt [double]$TotalBudgetSeconds) {
                $gateReason = ('the total budget of ' + [string]$TotalBudgetSeconds +
                               ' s was already spent (' + ('{0:N0}' -f $elapsedTotal) + ' s)')
            }
        }
        if (-not [string]::IsNullOrEmpty($gateReason)) {
            Write-Phase7Line ('SCENARIO ' + $case.Id + ' - ' + $case.Title + ' (' +
                              [string]$case.DriverCount + ' drivers)')
            Write-Phase7Line ('  NOT ENTERED: ' + $gateReason)
            Write-Phase7Line ''
            $null = $measurements.Add([pscustomobject]@{
                Id = $case.Id; Selected = $true; Entered = $false; Reason = $gateReason
                SensitivityMs = [double]0; DriverCount = [int]$case.DriverCount
            })
            continue
        }

        Write-Phase7Line ('SCENARIO ' + $case.Id + ' - ' + $case.Title)
        Write-Phase7Line ('-' * (11 + $case.Title.Length))
        Write-Phase7Line ('  requested drivers    : ' + [string]$case.DriverCount +
                          ' (' + [string](Get-Phase7CostLineCount -DriverCount $case.DriverCount) +
                          ' Cost Lines + ' +
                          [string]([int]$case.DriverCount - (Get-Phase7CostLineCount -DriverCount $case.DriverCount)) +
                          ' Risks)')
        Write-Phase7Line ('  requested iterations : ' + [string]$case.Iterations)

        # --- the model, through the accepted fixture ------------------------
        $model = New-Phase7TimingModel -DriverCount ([int]$case.DriverCount)
        $fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $null = Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
            -Inspection $inspection -Model $model
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name) `
            -Value ([double]$case.Iterations)
        # A BLANK Random Seed IS the AUTO request. AUTO is used deliberately: it
        # is the only mode in which "the AUTO nonce did not move during
        # sensitivity" is a check with anything to catch.
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.random_seed.defined_name) -Value $null
        $calculateWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $null = Invoke-Phase5ProductionOperation -Excel $excel -Operation 'PCCM_Calculate' `
            -Stage ('establishing the Phase-7 timing fixture for ' + [string]$case.DriverCount + ' drivers')
        $calculateWatch.Stop()
        $fixtureWatch.Stop()

        # --- the counts, READ BACK from the workbook -------------------------
        # Not taken from the model: what is timed is what the workbook holds.
        $costIds = @(Get-IdColumnValues -Workbook $wb -Info $costRegister)
        $riskIds = @(Get-IdColumnValues -Workbook $wb -Info $riskRegister)
        $iterationsSet = [string](Get-NamedValue -Workbook $wb `
            -DefinedName ([string]$simInspection.controls.monte_carlo_iterations.defined_name))
        Write-Phase7Line ('  Cost Lines in book   : ' + [string]$costIds.Count)
        Write-Phase7Line ('  Risks in book        : ' + [string]$riskIds.Count)
        Write-Phase7Line ('  drivers in book      : ' + [string]($costIds.Count + $riskIds.Count))
        Write-Phase7Line ('  iterations control   : ' + $iterationsSet)
        Write-Phase7Line ('  fixture build time   : ' + (Format-Phase7Seconds $fixtureWatch.Elapsed.TotalMilliseconds) +
                          '  [setup, NOT part of any measurement]')
        Write-Phase7Line ('  PCCM_Calculate time  : ' + (Format-Phase7Seconds $calculateWatch.Elapsed.TotalMilliseconds) +
                          '  [setup, NOT part of any measurement]')

        # --- the simulation the sensitivity will explain ---------------------
        $simulationWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $simulationResult = Invoke-Phase6Simulation -Excel $excel
        $simulationWatch.Stop()
        $statusBefore = [string]$excel.Run('PCCM_SimulationStatus')
        Write-Phase7Line ('  PCCM_RunSimulation   : ' + $simulationResult)
        Write-Phase7Line ('  simulation time      : ' + (Format-Phase7Seconds $simulationWatch.Elapsed.TotalMilliseconds) +
                          '  [reported separately, NOT the measurement]')
        Write-Phase7Line ('  status before        : ' + $statusBefore)

        if ($simulationResult -notlike 'OK|*') {
            Write-Phase7Line '  ABANDONED: the simulation did not succeed, so there is nothing current to explain.'
            Write-Phase7Line ''
            $null = $measurements.Add([pscustomobject]@{
                Id = $case.Id; Selected = $true; Entered = $true
                Reason = 'the simulation did not succeed'
                SensitivityMs = [double]0; DriverCount = [int]$case.DriverCount
            })
            $gateReason = 'an earlier scenario could not produce a successful simulation'
            continue
        }

        # --- the identity that must not move --------------------------------
        $before = Get-Phase6State -Workbook $wb -Inspection $simInspection
        $activeBank = Get-Phase6ActiveBank -State $before

        # ===================================================================
        # THE MEASUREMENT
        # ===================================================================
        # NOTHING between the two Stopwatch statements except the ONE
        # Application.Run of the REAL public endpoint. The automation envelope
        # is opened before the clock starts and the announcement is read after
        # it stops, so neither is inside the measured interval.
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $sensitivityWatch = [System.Diagnostics.Stopwatch]::StartNew()
        $excel.Run('PCCM_RunSensitivity') | Out-Null
        $sensitivityWatch.Stop()
        $sensitivityResult = [string]$excel.Run('PCCM_AutomationResult')
        # ===================================================================

        $after = Get-Phase6State -Workbook $wb -Inspection $simInspection
        $statusAfter = [string]$excel.Run('PCCM_SimulationStatus')

        Write-Phase7Line ''
        Write-Phase7Line ('  >>> PCCM_RunSensitivity ELAPSED : ' +
                          (Format-Phase7Seconds $sensitivityWatch.Elapsed.TotalMilliseconds))
        Write-Phase7Line ('  sensitivity result   : ' + $sensitivityResult)
        Write-Phase7Line ('  status after         : ' + $statusAfter)

        # --- what the run produced ------------------------------------------
        $recordCount = 0
        if (-not [string]::IsNullOrEmpty($activeBank)) {
            $stamp = Get-Phase7SensitivityStamp -Workbook $wb -Inspection $simInspection -Bank $activeBank
            Write-Phase7Line ('  active bank          : ' + $activeBank)
            foreach ($key in $stamp.Keys) {
                Write-Phase7Line ('    stamp.' + ([string]$key).PadRight(20) + ' = ' + (Format-SimValue $stamp[$key]))
            }
            if ($stamp['record_count'] -is [double]) { $recordCount = [int][double]$stamp['record_count'] }
            $counts = Get-Phase7SensitivityStatusCounts -Workbook $wb -Inspection $simInspection `
                -Bank $activeBank -RecordCount $recordCount
            Write-Phase7Line ('  sensitivity records  : ' + [string]$recordCount)
            Write-Phase7Line ('  eligible ranked      : ' + [string]$counts.Ranked)
            Write-Phase7Line ('  zero variance        : ' + [string]$counts.NoVariance)
            Write-Phase7Line ('  other status labels  : ' + [string]$counts.Other +
                              $(if ($counts.Other -gt 0) { ' (' + (($counts.OtherLabels) -join ', ') + ')' } else { '' }))
        } else {
            Write-Phase7Line '  active bank          : (none - nothing has ever been published)'
        }

        $sheetState = Get-Phase7SensitivitySheetState -Workbook $wb
        Write-Phase7Line ('  Sensitivity sheet    : ' + $sheetState.Availability)
        Write-Phase7Line ('  sheet rows populated : ' + [string]$sheetState.PopulatedRows +
                          ' of a ' + [string]$script:Phase7SensitivitySheet.RowWindow + '-row window')
        Write-Phase7Line ('  sheet first row      : ' + $sheetState.FirstRow)

        # --- the invariants --------------------------------------------------
        Write-Phase7Line '  identity invariants across PCCM_RunSensitivity:'
        $invariants = @()
        if (-not [string]::IsNullOrEmpty($activeBank)) {
            $invariants += (Compare-Phase7Invariant -Label ('bank ' + $activeBank + ' run_id') `
                -Before $before[('bank_' + $activeBank)]['run_id'] `
                -After  $after[('bank_' + $activeBank)]['run_id'])
            $invariants += (Compare-Phase7Invariant -Label ('bank ' + $activeBank + ' result_digest') `
                -Before $before[('bank_' + $activeBank)]['result_digest'] `
                -After  $after[('bank_' + $activeBank)]['result_digest'])
            $invariants += (Compare-Phase7Invariant -Label ('bank ' + $activeBank + ' consumed_auto_nonce') `
                -Before $before[('bank_' + $activeBank)]['consumed_auto_nonce'] `
                -After  $after[('bank_' + $activeBank)]['consumed_auto_nonce'])
        }
        $invariants += (Compare-Phase7Invariant -Label 'shared next_auto_nonce' `
            -Before $before['shared']['next_auto_nonce'] -After $after['shared']['next_auto_nonce'])
        $invariants += (Compare-Phase7Invariant -Label 'pending AUTO nonce cell' `
            -Before $before['pending_auto_nonce'] -After $after['pending_auto_nonce'])
        $invariants += (Compare-Phase7Invariant -Label 'shared last_run_id' `
            -Before $before['shared']['last_run_id'] -After $after['shared']['last_run_id'])
        foreach ($invariant in $invariants) { Write-Phase7Line $invariant.Text }

        # AND EVERY OTHER CAPTURED FIELD, not just the named ones. A sensitivity
        # run that moved something nobody thought to name would otherwise pass a
        # check that listed only what was expected to hold.
        #
        # THE `derived` ROWS ARE EXCLUDED FROM THE SWEEP AND REPORTED SEPARATELY,
        # because moving them is the endpoint doing its job: RequireCurrentRun
        # asks modSimReport.PCCM_SimulationStatus whether the run is CURRENT
        # before anything else happens, and that evaluation is what
        # simulation_status and status_evaluated_at record. Counting a re-read of
        # the status as a violated invariant would report the gate working as a
        # defect. The GROUP comes from the projection - this file does not decide
        # which rows are derived.
        $groups = $simInspection.sim_data.run_identity.groups
        $moved = @()
        $derived = @()
        foreach ($blockKey in $before.Keys) {
            if ($blockKey -eq 'pending_auto_nonce') {
                if (-not (Test-SimSameValue -A $before[$blockKey] -B $after[$blockKey])) { $moved += $blockKey }
                continue
            }
            foreach ($fieldKey in $before[$blockKey].Keys) {
                $group = ''
                if ($null -ne $groups.PSObject.Properties[$fieldKey]) { $group = [string]$groups.$fieldKey }
                if ($group -eq 'derived') {
                    $derived += ('    ' + $blockKey + '.' + $fieldKey + ': before ' +
                                 (Format-SimValue $before[$blockKey][$fieldKey]) + ', after ' +
                                 (Format-SimValue $after[$blockKey][$fieldKey]))
                    continue
                }
                if (-not (Test-SimSameValue -A $before[$blockKey][$fieldKey] -B $after[$blockKey][$fieldKey])) {
                    $moved += ($blockKey + '.' + $fieldKey)
                }
            }
        }
        if ($moved.Count -eq 0) {
            Write-Phase7Line '    every other captured _SimData run-identity field is unchanged'
        } else {
            Write-Phase7Line ('    FIELDS THAT MOVED: ' + ($moved -join ', '))
        }
        Write-Phase7Line '  the derived status rows, which the CURRENT gate re-evaluates by design:'
        foreach ($line in $derived) { Write-Phase7Line $line }
        Write-Phase7Line ''

        $null = $measurements.Add([pscustomobject]@{
            Id = $case.Id; Selected = $true; Entered = $true; Reason = ''
            SensitivityMs = [double]$sensitivityWatch.Elapsed.TotalMilliseconds
            DriverCount = [int]$case.DriverCount
        })

        if ($sensitivityWatch.Elapsed.TotalSeconds -gt [double]$SensitivityBudgetSeconds) {
            $gateReason = ('scenario ' + $case.Id + ' took ' +
                           ('{0:N1}' -f $sensitivityWatch.Elapsed.TotalSeconds) +
                           ' s, over the ' + [string]$SensitivityBudgetSeconds +
                           ' s budget; the larger scenarios were not entered')
        }
    }

    $excel.Run('PCCM_AutomationEnd') | Out-Null
} catch {
    Write-Phase7Line ''
    Write-Phase7Line ('THE MEASUREMENT SESSION RAISED: ' + (Format-Err $_))
    Write-Phase7Line 'Whatever was measured before this point is above and is still valid.'
} finally {
    # --- shutdown, the accepted path, leaf before parent --------------------
    $rel = New-ReleaseLedger 'phase-7 timing instance'
    try {
        if ($null -ne $wb) {
            # NEVER SAVED. The disposable copy is discarded, and no measurement
            # carries a save cost.
            try { $wb.Close($false); $rel.WorkbookClosed = $true }
            catch { $null = $rel.Failed.Add('Workbook.Close') }
        }
        Invoke-NamedRelease $rel $wb        'Workbook';  $wb        = $null
        Invoke-NamedRelease $rel $workbooks 'Workbooks'; $workbooks = $null
        if ($null -ne $excel) {
            try { $excel.Quit(); $rel.QuitCalled = $true }
            catch { $null = $rel.Failed.Add('Application.Quit') }
        }
        Invoke-NamedRelease $rel $excel 'Application'; $excel = $null

        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()

        $rel.NaturalExit = Wait-ExcelExit -Identity $excelIdentity
        if (-not $rel.NaturalExit) {
            $rel.EmergencyRequired = $true
            Write-Phase7Line (Invoke-EmergencyExcelCleanup -Identity $excelIdentity `
                -Label 'phase-7 timing instance')
        }
    } catch {
        Write-Phase7Line ('Shutdown raised: ' + (Format-Err $_))
    }
    Write-Phase7Line 'SHUTDOWN'
    Write-Phase7Line '--------'
    Write-Phase7Line (Format-ReleaseLedger $rel)
    $transient = @(Get-TransientFailures)
    if ($transient.Count -gt 0) {
        Write-Phase7Line ('transient COM release failures: ' + ($transient -join '; '))
    } else {
        Write-Phase7Line 'every transient COM object released cleanly'
    }
}

# ===========================================================================
# SUMMARY
# ===========================================================================
$runStopwatch.Stop()
Write-Phase7Line ''
Write-Phase7Line 'SUMMARY - PCCM_RunSensitivity WALL-CLOCK TIME'
Write-Phase7Line '--------------------------------------------'
foreach ($row in $measurements) {
    if (-not $row.Selected) {
        Write-Phase7Line ('  ' + $row.Id + '  ' + ([string]$row.DriverCount).PadLeft(4) +
                          ' drivers   NOT SELECTED - ' + $row.Reason)
    } elseif (-not $row.Entered) {
        Write-Phase7Line ('  ' + $row.Id + '  ' + ([string]$row.DriverCount).PadLeft(4) +
                          ' drivers   NOT ENTERED - ' + $row.Reason)
    } elseif ($row.SensitivityMs -le 0) {
        Write-Phase7Line ('  ' + $row.Id + '  ' + ([string]$row.DriverCount).PadLeft(4) +
                          ' drivers   NOT MEASURED - ' + $row.Reason)
    } else {
        Write-Phase7Line ('  ' + $row.Id + '  ' + ([string]$row.DriverCount).PadLeft(4) +
                          ' drivers   ' + (Format-Phase7Seconds $row.SensitivityMs))
    }
}
Write-Phase7Line ('  ' + [string]@($measurements).Count + ' of ' + [string]$declared.Count +
                  ' declared scenario(s) accounted for above')
Write-Phase7Line ''
Write-Phase7Line ('whole run wall clock   : ' + (Format-Phase7Seconds $runStopwatch.Elapsed.TotalMilliseconds))
Write-Phase7Line ''
Write-Phase7Line 'NO CONCLUSION IS DRAWN HERE. These numbers are the input to an'
Write-Phase7Line 'architecture decision that has not been taken; nothing was capped and'
Write-Phase7Line 'nothing was subsampled to make a number look better.'
Write-Phase7Line ''
Write-Phase7Line ('report                 : ' + $script:Phase7ReportPath)

if ($KeepArtifacts) {
    Write-Host ''
    Write-Host ('Artefacts kept in ' + $tempRoot) -ForegroundColor Yellow
} else {
    $kept = $script:Phase7ReportPath
    try {
        $kept = Join-Path ([System.IO.Path]::GetTempPath()) `
            ('pccm-phase7-timing-report-' + (Get-Date).ToString('yyyyMMdd-HHmmss') + '.txt')
        Copy-Item -LiteralPath $script:Phase7ReportPath -Destination $kept -Force
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    } catch { }
    Write-Host ''
    Write-Host ('Report kept at ' + $kept) -ForegroundColor Yellow
}
