<#
.SYNOPSIS
    PCCM Phase 10 - THE FIXTURE SEMANTIC-EQUIVALENCE GATE. Windows only.

.DESCRIPTION
    WHAT THIS DECIDES. PERF-LARGE was operator-aborted after more than four hours
    in fixture construction, because the accepted fixture invokes a production
    Add per driver and each of those is O(register rows x project years):
    snapshot both tables for rollback, re-sync the profiling grid, validate the
    structure, re-protect fourteen sheets, recalculate four sheets. Three hundred
    of those is O(N^2 x years) and is impractical.

    `Set-BenchmarkBulkFixture` reaches the same state by writing the INPUTS a
    user types as rectangular blocks and letting ONE real `PCCM_ApplyTimeline`
    produce every structural thing - the year columns, the profiling rows and
    their permanent-ID keying, the applied-timeline names, and
    ValidateStructure.

    THAT IS ONLY LEGITIMATE IF THE TWO PATHS END IN THE SAME WORKBOOK. This
    builds PERF-SMALL BOTH WAYS, in two separate Excel sessions over two
    disposable copies of the same Stage-A build, captures a full state snapshot
    from each, and compares them field for field - then runs the REAL
    `PCCM_Calculate` on both and compares production's own status and
    fingerprint.

    IT TESTS THE SHIPPING BUILDER. Every function it needs is lifted out of
    `bootstrap/windows/phase10_benchmark.ps1` BY AST - its real bytes, not a
    copy - so a builder that drifted from what the benchmark runs would fail
    here rather than pass against a restatement of itself.

    IT ASSERTS NOTHING ITSELF. It prints one tagged line per field family and
    exits 0; `tests/test_phase10_fixture_equivalence.py` decides. A gate that
    both produced and judged its own evidence would be one authority, not two.

.PARAMETER Runner
    The benchmark runner to lift the builder from. Defaults to the shipping one.

.NOTES
    Two Excel sessions, two Stage-B bootstraps, two disposable workbooks. Neither
    workbook is ever saved. Shutdown is the accepted com_lifecycle path.

    Prints EQUIV|<field family>|<match|differ>|<detail> and
    CALC|<mode>|<status>|<fingerprint>. Exit 0 always.
#>
param(
    [string]$Runner,
    [string]$BuildDir,
    [string]$WorkDir
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pccmRoot = Split-Path -Parent $here
$windows = Join-Path $pccmRoot 'bootstrap/windows'
if ([string]::IsNullOrWhiteSpace($Runner)) { $Runner = Join-Path $windows 'phase10_benchmark.ps1' }
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }
if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = [System.IO.Path]::GetTempPath() }
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

# THE ACCEPTED FILES ARE REUSED, NOT REIMPLEMENTED. Both are definition-only at
# top level.
. (Join-Path $windows 'com_lifecycle.ps1')
. (Join-Path $windows 'phase5_gate_b_scenarios.ps1')

# --- THE SHIPPING BUILDER, BY AST -------------------------------------------
$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}
# Ordered so the overrides land after the dot-sourced definitions they replace,
# exactly as they do in the runner itself.
$wanted = @(
    'Write-RowObject', 'Get-NamedValue', 'Set-NamedValue', 'Get-TableColumnNames',
    'Get-TableBody', 'Get-TableRowCount', 'Add-BlankTableRow', 'Remove-TableRow',
    'Reset-Phase5FxTable', 'Set-TableCell', 'Get-IdColumnValues',
    'Set-BenchmarkRangeBlock', 'Set-BenchmarkRegisterRowCount',
    'Get-BenchmarkPermanentId', 'New-BenchmarkRegisterBlock', 'New-BenchmarkWeightBlock',
    'Set-BenchmarkBulkFixture',
    'Import-BenchmarkFixtureWindow', 'Get-BenchmarkProtectionState',
    'Assert-BenchmarkProtectionApplied', 'Open-BenchmarkFixtureWindow',
    'Invoke-BenchmarkWindowRollback', 'Close-BenchmarkFixtureWindow',
    'New-BenchmarkWeights', 'New-BenchmarkDriver', 'New-BenchmarkModel',
    'Get-BenchmarkSeed')
$found = @{}
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($wanted -contains $fn.Name) { $found[$fn.Name] = $fn.Extent.Text }
}
foreach ($name in $wanted) {
    if (-not $found.ContainsKey($name)) {
        Write-Output ('MISSING|' + $name + ' is not defined in ' + (Split-Path -Leaf $runnerPath))
        exit 0
    }
    Invoke-Expression $found[$name]
}
# The runner keeps these in script scope; the lifted copies need them too.
$script:FixtureWindowModule = 'modPhase10FixtureWindow'
$script:FixtureWindowSource = 'phase10_fixture_window.bas'

# ===========================================================================
# THE STATE SNAPSHOT
# ===========================================================================
# EVERY FIELD FAMILY THE AUTHORISATION NAMES, read out of the live workbook and
# rendered as text so two of them can be compared exactly. Nothing here is
# computed: each entry is what the workbook says.
function Get-EquivalenceSnapshot {
    param($Excel, $Workbook, $Manifest, $Inspection, $SimInspection)
    $state = New-Object System.Collections.Specialized.OrderedDictionary

    $registerByKey = @{}
    foreach ($register in @($Manifest.registers)) { $registerByKey[[string]$register.key] = $register }
    $gridByKey = @{}
    foreach ($grid in @($Manifest.grids)) { $gridByKey[[string]$grid.key] = $grid }

    # --- identifiers, in physical order --------------------------------------
    foreach ($key in @('cost_lines', 'risk_register')) {
        $register = $registerByKey[$key]
        $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $register)
        $state.Add(($key + '.ids'), ($ids -join ','))
    }
    # --- the counters --------------------------------------------------------
    foreach ($counter in @($Manifest.counters)) {
        $state.Add(('counter.' + [string]$counter.key),
                   [string](Get-NamedValue -Workbook $Workbook `
                        -DefinedName ([string]$counter.defined_name)))
    }
    # --- every business input in both registers ------------------------------
    # The whole body, column by column, so distributions, currencies, profiles,
    # quantities, probabilities and the three-point values are all covered by one
    # comparison that cannot omit a column it was not told about.
    foreach ($key in @('cost_lines', 'risk_register')) {
        $register = $registerByKey[$key]
        $state.Add(($key + '.columns'),
                   ((@(Get-TableColumnNames -Workbook $Workbook -SheetName $register.sheet `
                        -TableName $register.table_name)) -join '|'))
        $lines = @()
        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $register.sheet `
                -TableName $register.table_name)) {
            $lines += ((@($row) -join '|'))
        }
        $state.Add(($key + '.body'), ($lines -join ' ;; '))
    }
    # --- FX ------------------------------------------------------------------
    $fx = $Inspection.input_tables.fx_rates
    $fxLines = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $fx.sheet `
            -TableName $fx.table_name)) {
        $fxLines += ((@($row) -join '|'))
    }
    $state.Add('fx.body', ($fxLines -join ' ;; '))
    # --- the Config inflation-profile master ---------------------------------
    $master = $Inspection.input_tables.inflation_profiles
    $masterLines = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $master.sheet `
            -TableName $master.table_name)) {
        $masterLines += ((@($row) -join '|'))
    }
    $state.Add('inflation_profiles.body', ($masterLines -join ' ;; '))
    # --- the applied timeline ------------------------------------------------
    foreach ($name in @('nmBaseYear_Applied', 'nmStartYear_Applied', 'nmDuration_Applied',
                        'nmLastYear_Applied', 'nmYearCount_Applied',
                        'nmInflFirstYear', 'nmInflLastYear')) {
        $state.Add(('applied.' + $name),
                   [string](Get-NamedValue -Workbook $Workbook -DefinedName $name))
    }
    # --- the generated year headers and the grid bodies ---------------------
    foreach ($key in @('cost_profiling', 'risk_profiling', 'inflation')) {
        $grid = $gridByKey[$key]
        $state.Add(($key + '.headers'),
                   ((@(Get-TableColumnNames -Workbook $Workbook -SheetName $grid.sheet `
                        -TableName $grid.table_name)) -join '|'))
        $lines = @()
        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $grid.sheet `
                -TableName $grid.table_name)) {
            $lines += ((@($row) -join '|'))
        }
        $state.Add(($key + '.body'), ($lines -join ' ;; '))
    }
    # --- production's own reports on itself ---------------------------------
    $state.Add('structural.state',
               [string](Get-NamedValue -Workbook $Workbook -DefinedName 'nmStructuralState'))
    $state.Add('structural.report', [string]$Excel.Run('PCCM_StructuralReport'))
    $state.Add('fingerprint.calculation_inputs', [string]$Excel.Run('PCCM_CurrentInputFingerprint'))
    $state.Add('fingerprint.simulation_request',
               [string]$Excel.Run('PCCM_CurrentSimulationRequestFingerprint'))
    $state.Add('modelcheck.calculation_state', [string]$Excel.Run('PCCM_ModelCheckCalculationState'))
    return $state
}

# ===========================================================================
# ONE PASS: BOOTSTRAP, BUILD ONE WAY, SNAPSHOT, CALCULATE, SHUT DOWN
# ===========================================================================
function Invoke-EquivalencePass {
    param([string]$Mode, $Manifest, $Inspection, $SimInspection, $Plan)
    $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss-fff')
    $tempRoot = Join-Path $WorkDir ('pccm-equivalence-' + $Mode.ToLower() + '-' + $stamp)
    $null = New-Item -ItemType Directory -Path $tempRoot -Force
    Copy-Item -LiteralPath (Join-Path $BuildDir ([string]$Manifest.stage_a_filename)) -Destination $tempRoot
    Copy-Item -LiteralPath (Join-Path $BuildDir 'vba') -Destination $tempRoot -Recurse

    $bootstrap = Join-Path $windows 'build_stage_b.ps1'
    & $bootstrap -BuildDir $tempRoot -Force | Out-Null
    $stageB = Join-Path $tempRoot ([string]$Manifest.stage_b_filename)
    if (-not (Test-Path -LiteralPath $stageB)) {
        throw ('the Stage-B bootstrap produced no workbook for the ' + $Mode + ' pass')
    }

    $excel = $null; $workbooks = $null; $wb = $null
    $rel = New-ReleaseLedger ('equivalence ' + $Mode)
    $snapshot = $null; $calcStatus = ''; $calcFingerprint = ''; $calcResult = ''
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $workbooks = $excel.Workbooks
        $wb = $workbooks.Open($stageB)

        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $Inspection
        Import-BenchmarkFixtureWindow -Excel $excel -Workbook $wb -Manifest $Manifest `
            -ScriptDir $windows

        $scenarioSpec = $null
        foreach ($entry in @($Plan.scenarios)) {
            if ([string]$entry.id -eq 'PERF-SMALL') { $scenarioSpec = $entry }
        }
        if ($null -eq $scenarioSpec) { throw 'the plan declares no PERF-SMALL scenario' }
        $model = New-BenchmarkModel -ScenarioSpec $scenarioSpec

        # THE SAME WINDOW BOTH PASSES USE. Neither builder gets a privilege the
        # other does not.
        $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $Manifest
        try {
            if ($Mode -eq 'Bulk') {
                $null = Set-BenchmarkBulkFixture -Excel $excel -Workbook $wb -Manifest $Manifest `
                    -Inspection $Inspection -Model $model -ScenarioSpec $scenarioSpec
            } else {
                $null = Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $Manifest `
                    -Inspection $Inspection -Model $model
            }
        } finally {
            $null = Close-BenchmarkFixtureWindow -Excel $excel -Manifest $Manifest
        }
        Set-NamedValue -Workbook $wb `
            -DefinedName ([string]$SimInspection.controls.random_seed.defined_name) `
            -Value ([double](Get-BenchmarkSeed))

        $snapshot = Get-EquivalenceSnapshot -Excel $excel -Workbook $wb -Manifest $Manifest `
            -Inspection $Inspection -SimInspection $SimInspection

        # AND THEN THE REAL ENDPOINT. Nothing in either builder produced a result;
        # this is where every number comes from, in both passes.
        $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
        $excel.Run('PCCM_Calculate') | Out-Null
        $calcResult = [string]$excel.Run('PCCM_AutomationResult')
        $calcStatus = [string]$excel.Run('PCCM_CalculationStatus')
        $calcFingerprint = [string]$excel.Run('PCCM_CalculationFingerprint')
        $excel.Run('PCCM_AutomationEnd') | Out-Null
    } finally {
        try {
            if ($null -ne $wb) { $wb.Close($false) | Out-Null }
        } catch { Write-Output ('SHUTDOWN|' + $Mode + '|close raised: ' + $_.Exception.Message) }
        Invoke-NamedRelease $rel $wb         'Workbook';   $wb         = $null
        Invoke-NamedRelease $rel $workbooks  'Workbooks';  $workbooks  = $null
        try {
            if ($null -ne $excel) { $excel.Quit() }
        } catch { Write-Output ('SHUTDOWN|' + $Mode + '|quit raised: ' + $_.Exception.Message) }
        Invoke-NamedRelease $rel $excel      'Application'; $excel     = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    return [pscustomobject]@{
        Mode = $Mode
        State = $snapshot
        CalcResult = $calcResult
        CalcStatus = $calcStatus
        CalcFingerprint = $calcFingerprint
    }
}

# ===========================================================================
# THE TWO PASSES, AND THE COMPARISON
# ===========================================================================
$manifest      = Get-Content -LiteralPath (Join-Path $BuildDir 'stage_b_manifest.json') -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath (Join-Path $BuildDir 'phase5_gate_b_inspection.json') -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath (Join-Path $BuildDir 'phase6_gate_b_inspection.json') -Raw | ConvertFrom-Json
$plan          = Get-Content -LiteralPath (Join-Path $BuildDir 'phase10_benchmark_plan.json') -Raw | ConvertFrom-Json

$passes = @{}
foreach ($mode in @('Endpoints', 'Bulk')) {
    try {
        $passes[$mode] = Invoke-EquivalencePass -Mode $mode -Manifest $manifest `
            -Inspection $inspection -SimInspection $simInspection -Plan $plan
        Write-Output ('PASS|' + $mode + '|BUILT')
    } catch {
        Write-Output ('PASS|' + $mode + '|RAISED|' + (($_.Exception.Message) -replace '\s+', ' '))
    }
}
if (-not ($passes.ContainsKey('Endpoints') -and $passes.ContainsKey('Bulk'))) {
    Write-Output 'EQUIV|<no comparison>|differ|one of the two passes did not build'
    exit 0
}

$reference = $passes['Endpoints']
$optimised = $passes['Bulk']
foreach ($field in @($reference.State.Keys)) {
    $left = [string]$reference.State[$field]
    $right = ''
    if ($optimised.State.Contains($field)) { $right = [string]$optimised.State[$field] }
    else { $right = '<absent>' }
    if ($left -ceq $right) {
        Write-Output ('EQUIV|' + $field + '|match|' + [string]$left.Length + ' chars')
    } else {
        Write-Output ('EQUIV|' + $field + '|differ|endpoints=' + $left + ' :: bulk=' + $right)
    }
}
foreach ($field in @($optimised.State.Keys)) {
    if (-not $reference.State.Contains($field)) {
        Write-Output ('EQUIV|' + $field + '|differ|only the bulk pass reported this field')
    }
}
foreach ($pass in @($reference, $optimised)) {
    Write-Output ('CALC|' + $pass.Mode + '|' + $pass.CalcResult + '|' + $pass.CalcStatus +
                  '|' + $pass.CalcFingerprint)
}
if (($reference.CalcFingerprint -ceq $optimised.CalcFingerprint) -and
    (-not [string]::IsNullOrWhiteSpace($reference.CalcFingerprint))) {
    Write-Output 'CALCEQUIV|match|the production calculation fingerprint is identical'
} else {
    Write-Output ('CALCEQUIV|differ|endpoints=' + $reference.CalcFingerprint +
                  ' :: bulk=' + $optimised.CalcFingerprint)
}
exit 0
