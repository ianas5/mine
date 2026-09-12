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

    THE VOCABULARY IS UNAMBIGUOUS, because run 1 printed `PASS|Endpoints|RAISED`
    for a pass that never produced a workbook, and a reader could take that for a
    pass. PASS is now reserved for a pass that COMPLETED:

      BUNDLE|<mode>|<relative path>|<sha256>   one required artifact, as copied
      BUNDLE|identical                         the two starting bundles agree
      BUNDLE|differ|<detail>                   they do not - nothing is built
      PASS|<mode>|COMPLETED|<detail>           this pass built and calculated
      FAIL|<mode>|SETUP|<detail>               the bundle could not be prepared
      FAIL|<mode>|BOOTSTRAP|<detail>           Stage-B did not produce a workbook
      FAIL|<mode>|RAISED|<detail>              anything after Excel started
      EQUIV|<family>|match|<detail>            a real comparison, equal
      EQUIV|<family>|differ|<detail>           a real comparison, unequal
      EQUIV|<not evaluated>|invalid|<detail>   NO comparison happened
      CALC|<mode>|<result>|<status>|<print>    only from a COMPLETED pass
      CALCEQUIV|match / differ                only when both passes COMPLETED

    `differ` is reserved for a comparison that actually ran. A setup failure is
    `invalid`, never `differ`, and emits no CALC or CALCEQUIV line at all.

    Exit 0 always.
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
    'Get-BulkOpVocabulary', 'Set-BulkOp', 'Get-BulkOp', 'Get-BulkComHResult',
    'Format-BulkFailureLine', 'Save-BulkFailure', 'New-BulkFailureLine', 'Reset-BulkOp',
    'Set-BenchmarkBulkFixture', 'Invoke-BenchmarkEndpointsResync',
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
# THE STARTING BUNDLE
# ===========================================================================
# WHAT RUN 1 GOT WRONG. This gate copied the Stage-A workbook and the generated
# `vba` directory into each disposable workdir and then invoked
# `build_stage_b.ps1 -BuildDir <that dir>`. It did NOT copy
# `stage_b_manifest.json`, which is the first thing the bootstrap reads:
#
#   build_stage_b.ps1:86   $manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'
#   build_stage_b.ps1:93   if (-not (Test-Path -LiteralPath $manifestPath)) {
#   build_stage_b.ps1:94       throw "stage_b_manifest.json not found at $manifestPath..."
#
# Both passes therefore failed before Excel was started, and the message told the
# operator to run Stage A - which had already been run, in the repository root,
# where the artifact still sat.
#
# THE LIST IS DERIVED, NOT COPIED FROM THE BENCHMARK. `build_stage_b.ps1`
# resolves exactly three things against the SUPPLIED BuildDir:
#
#   stage_b_manifest.json          :86  - the authority for everything else
#   $manifest.stage_a_filename     :98  - the Stage-A workbook
#   <BuildDir>/<leaf of manifest.vba.generated_dir>
#                                  :125 - the GENERATED modules only
#
# and two things against the REPOSITORY, deliberately:
#
#   $pccmRoot/$manifest.vba.source_dir  :124 - the version-controlled modules
#   $srcDir/$manifest.vba.document_module.file  :274 - ThisWorkbook
#
# Those two are shared input and are the same files for any build, which is the
# distinction build_stage_b.ps1's own comment draws. Nothing else is read from the
# BuildDir: `grep -c inspection build_stage_b.ps1` is 0, so the two Gate-B
# inspection projections the benchmark also copies are not part of this contract
# and are read by this gate from the repository build directory.
#
# `$manifest.stage_b_filename` is the OUTPUT, and it must not be carried in: a
# stale repository .xlsm copied into the workdir would be opened instead of the
# one this bundle builds.
function Get-BundleArtifacts {
    param($Manifest)
    $generatedLeaf = Split-Path -Leaf ([string]$Manifest.vba.generated_dir)
    if ([string]::IsNullOrWhiteSpace($generatedLeaf)) {
        throw 'the manifest declares no generated VBA directory'
    }
    return @(
        [pscustomobject]@{ Name = 'stage_b_manifest.json'; Kind = 'file' }
        [pscustomobject]@{ Name = [string]$Manifest.stage_a_filename; Kind = 'file' }
        [pscustomobject]@{ Name = $generatedLeaf; Kind = 'directory' }
    )
}

# ONE PRISTINE BUNDLE PER PASS, AND ITS DIGEST.
#
# Each pass gets its OWN directory under the work root, so neither can see the
# other's Stage-B workbook, its mutated copy of the manifest, or anything else.
# The repository build directory is READ and never written.
function New-EquivalenceBundle {
    param([string]$Mode, $Manifest, [string]$Stamp)
    $root = Join-Path $WorkDir ('pccm-equivalence-' + $Mode.ToLower() + '-' + $Stamp)
    if (Test-Path -LiteralPath $root) {
        throw ('the disposable bundle directory ' + $root + ' already exists')
    }
    $null = New-Item -ItemType Directory -Path $root -Force

    $digests = New-Object System.Collections.Specialized.OrderedDictionary
    foreach ($artifact in @(Get-BundleArtifacts -Manifest $Manifest)) {
        $source = Join-Path $BuildDir ([string]$artifact.Name)
        if (-not (Test-Path -LiteralPath $source)) {
            throw ([string]$artifact.Name + ' is not in the repository build directory ' +
                   $BuildDir + '. Run the Stage-A build first: python pccm\builder\build_stage_a.py')
        }
        if ([string]$artifact.Kind -eq 'directory') {
            Copy-Item -LiteralPath $source -Destination $root -Recurse
        } else {
            Copy-Item -LiteralPath $source -Destination $root
        }
        # EVERY FILE THAT ARRIVED, HASHED. A directory contributes one entry per
        # file so a missing generated module is a difference rather than a silence.
        $landed = Join-Path $root ([string]$artifact.Name)
        if (-not (Test-Path -LiteralPath $landed)) {
            throw ([string]$artifact.Name + ' did not arrive in ' + $root)
        }
        foreach ($file in @(Get-ChildItem -LiteralPath $landed -Recurse -File -ErrorAction Stop)) {
            # THE KEY IS SEPARATOR-INDEPENDENT. Trimming only a backslash left a
            # leading separator on any host whose separator is not one, and the
            # two bundles would then be compared by keys that no longer named the
            # same artifact. Both separators are trimmed and the survivor is
            # normalised, so `vba/modConstants.bas` is the key either way.
            $relative = $file.FullName.Substring($root.Length)
            $relative = $relative.TrimStart([char]92, [char]47).Replace([char]92, [char]47)
            $digests.Add($relative, [string](Get-FileHash -LiteralPath $file.FullName `
                -Algorithm SHA256).Hash)
        }
    }
    # AND THE OUTPUT MUST NOT BE PRESENT. A stale Stage-B workbook carried in would
    # be opened instead of the one this bundle is about to build.
    $stageB = Join-Path $root ([string]$Manifest.stage_b_filename)
    if (Test-Path -LiteralPath $stageB) {
        throw ('a Stage-B workbook is already present at ' + $stageB +
               ' before the bootstrap ran, so the bundle carried a stale build')
    }
    # NOTHING IS WRITTEN HERE. A function that emits to the output stream AND
    # returns a value has its return polluted by everything it wrote: `$bundle`
    # would be an array of report lines with the object at the end, and the very
    # first `$bundle.Root` would fail. The caller prints from `Digests`.
    return [pscustomobject]@{
        Mode = $Mode
        Root = $root
        Digests = $digests
        StageB = $stageB
    }
}

# THE TWO STARTING STATES ARE THE SAME STARTING STATE, or nothing is built.
#
# Both bundles come from one repository Stage-A build, so they SHOULD be
# identical - which is exactly why it is worth proving rather than assuming. A
# difference here would mean the two passes were never comparable and every
# EQUIV line afterwards would be measuring the wrong thing.
function Test-BundleIdentity {
    param($Left, $Right)
    $problems = @()
    $leftKeys = @($Left.Digests.Keys)
    $rightKeys = @($Right.Digests.Keys)
    foreach ($key in $leftKeys) {
        if (-not $Right.Digests.Contains($key)) {
            $problems += ($key + ' is only in the ' + $Left.Mode + ' bundle')
            continue
        }
        if ([string]$Left.Digests[$key] -cne [string]$Right.Digests[$key]) {
            $problems += ($key + ' differs: ' + $Left.Mode + '=' +
                          [string]$Left.Digests[$key] + ' ' + $Right.Mode + '=' +
                          [string]$Right.Digests[$key])
        }
    }
    foreach ($key in $rightKeys) {
        if (-not $Left.Digests.Contains($key)) {
            $problems += ($key + ' is only in the ' + $Right.Mode + ' bundle')
        }
    }
    if ($leftKeys.Count -lt 1) { $problems += 'the bundles are empty' }
    return $problems
}

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
    param($Bundle, $Manifest, $Inspection, $SimInspection, $Plan)
    $Mode = [string]$Bundle.Mode
    $tempRoot = [string]$Bundle.Root

    # THE BOOTSTRAP RUNS AGAINST THIS PASS'S OWN BUNDLE and nothing shared.
    $bootstrap = Join-Path $windows 'build_stage_b.ps1'
    & $bootstrap -BuildDir $tempRoot -Force | Out-Null
    # THE EXIT CODE, NOT ONLY THE FILE. A build that saved the .xlsm and was then
    # refused at a later operation leaves the workbook on disk WITHOUT its modules,
    # its buttons or its protection - and Test-Path alone would let this pass open
    # it and report a fixture result against a half-built workbook. The bootstrap
    # names the failing operation in its own transcript above.
    $bootstrapExit = $LASTEXITCODE
    $stageB = [string]$Bundle.StageB
    if ($bootstrapExit -ne 0) {
        throw ('BOOTSTRAP: the Stage-B bootstrap for the ' + $Mode + ' pass exited ' +
               [string]$bootstrapExit + '; its transcript above names the failing operation. ' +
               'A workbook may exist and be half-built, so this pass is not run.')
    }
    if (-not (Test-Path -LiteralPath $stageB)) {
        throw ('BOOTSTRAP: the Stage-B bootstrap produced no workbook for the ' + $Mode +
               ' pass at ' + $stageB)
    }

    $excel = $null; $workbooks = $null; $wb = $null
    $rel = New-ReleaseLedger ('equivalence ' + $Mode)
    $snapshot = $null; $calcStatus = ''; $calcFingerprint = ''; $calcResult = ''
    try {
        # THE PREFIX IS LABELLED TOO. Run 11 raised RPC_E_CALL_REJECTED in the Bulk
        # pass and the diagnostic read '<before the first bulk operation>': the
        # first label was set at the window open, and the nine Excel calls between
        # the bootstrap's return and that point - starting Excel, three property
        # sets, Workbooks, Open, PCCM_AutomationBegin, the FX seed read and the
        # shim import - ran unlabelled. Each is now named immediately before it
        # runs, and the catch below saves the label at the throw, exactly as the
        # builder's catch does. Nothing is retried and nothing waits.
        if ($Mode -eq 'Bulk') { Reset-BulkOp }
        try {
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.excel.create' }
            $excel = New-Object -ComObject Excel.Application
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.excel.visible' }
            $excel.Visible = $false
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.excel.displayalerts' }
            $excel.DisplayAlerts = $false
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.excel.asktoupdatelinks' }
            $excel.AskToUpdateLinks = $false
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.workbooks.acquire' }
            $workbooks = $excel.Workbooks
            # A REFUSED PROPERTY GET CAN ANSWER WITH NOTHING - Stage-B saw exactly
            # that on Windows - and the next statement would then fail under the
            # next label. Named here, under its own.
            if ($null -eq $workbooks) {
                throw 'RAISED: the Workbooks collection read answered with nothing'
            }
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.workbook.open' }
            $wb = $workbooks.Open($stageB)

            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.automation.begin' }
            $excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.fxseed.read' }
            $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $Inspection
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.preflight.window.import' }
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
            # THE WINDOW IS LABELLED FOR THE BULK BUILDER, so a refusal in P10FW_Begin
            # or P10FW_End is named as that and not as the first or last fixture step.
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.window.open' }
            $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $Manifest
        } catch {
            # SAVED AT THE THROW. The window's finally is not entered from here -
            # it was never opened - so nothing relabels; saved all the same, so the
            # outer catch never has to read a label after the fact.
            if ($Mode -eq 'Bulk') { Save-BulkFailure -ErrorRecord $_ }
            throw
        }
        try {
            if ($Mode -eq 'Bulk') {
                try {
                    $null = Set-BenchmarkBulkFixture -Excel $excel -Workbook $wb -Manifest $Manifest `
                        -Inspection $Inspection -Model $model -ScenarioSpec $scenarioSpec
                } catch {
                    # SAVED HERE, before the finally below relabels the window close
                    # and the outer catch asks which operation was in flight.
                    Save-BulkFailure -ErrorRecord $_
                    throw
                }
            } else {
                $null = Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $Manifest `
                    -Inspection $Inspection -Model $model
                # THE SAME RESYNCHRONISATION THE RUNNER MAKES, lifted from it: the
                # reference fixture ends with production's own SyncRows over the
                # complete registers, so the trace column of the last driver is
                # what production copies there, in both passes.
                $null = Invoke-BenchmarkEndpointsResync -Excel $excel
            }
        } finally {
            if ($Mode -eq 'Bulk') { Set-BulkOp 'bulk.window.close' }
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
    }
    # COMPLETED MEANS BOTH HALVES ARRIVED. A snapshot with no calculation, or a
    # calculation with no snapshot, is not a pass this gate may compare.
    if ($null -eq $snapshot) {
        throw ('RAISED: the ' + $Mode + ' pass produced no state snapshot')
    }
    if ([string]::IsNullOrWhiteSpace($calcFingerprint)) {
        throw ('RAISED: the ' + $Mode + ' pass produced no calculation fingerprint; ' +
               'result was ' + [char]39 + $calcResult + [char]39)
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

# --- BOTH STARTING BUNDLES, BEFORE EITHER IS BUILT --------------------------
$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss-fff')
$bundles = @{}
$setupFailed = $false
foreach ($mode in @('Endpoints', 'Bulk')) {
    try {
        $bundles[$mode] = New-EquivalenceBundle -Mode $mode -Manifest $manifest -Stamp $stamp
        foreach ($relative in @($bundles[$mode].Digests.Keys)) {
            Write-Output ('BUNDLE|' + $mode + '|' + $relative + '|' +
                          [string]$bundles[$mode].Digests[$relative])
        }
    } catch {
        $setupFailed = $true
        Write-Output ('FAIL|' + $mode + '|SETUP|' + (($_.Exception.Message) -replace '\s+', ' '))
    }
}

if (-not $setupFailed) {
    $problems = @(Test-BundleIdentity -Left $bundles['Endpoints'] -Right $bundles['Bulk'])
    if ($problems.Count -gt 0) {
        $setupFailed = $true
        Write-Output ('BUNDLE|differ|' + ($problems -join '; '))
    } else {
        Write-Output ('BUNDLE|identical|' +
                      [string]@($bundles['Endpoints'].Digests.Keys).Count + ' artifact(s)')
    }
}

# --- THE TWO PASSES ---------------------------------------------------------
# NOTHING IS BUILT IF THE STARTING STATES DID NOT AGREE. Refusing here is the
# whole point of proving identity before Excel is started.
$passes = @{}
if (-not $setupFailed) {
    foreach ($mode in @('Endpoints', 'Bulk')) {
        try {
            $passes[$mode] = Invoke-EquivalencePass -Bundle $bundles[$mode] -Manifest $manifest `
                -Inspection $inspection -SimInspection $simInspection -Plan $plan
            Write-Output ('PASS|' + $mode + '|COMPLETED|fixture built and PCCM_Calculate ran')
        } catch {
            # THE STAGE IS NAMED. A bootstrap that produced no workbook and a COM
            # call that raised inside Excel are different facts, and run 1 proved
            # that collapsing them into one word costs a round.
            $detail = ($_.Exception.Message) -replace '\s+', ' '
            $stage = 'RAISED'
            if ($detail -like 'BOOTSTRAP:*') { $stage = 'BOOTSTRAP'; $detail = $detail.Substring(10).Trim() }
            elseif ($detail -like 'RAISED:*') { $detail = $detail.Substring(7).Trim() }
            Write-Output ('FAIL|' + $mode + '|' + $stage + '|' + $detail)
            # THE EXACT BULK OPERATION, SEPARATELY. Two runs printed the line above
            # and nothing else, and the correction that follows depends on which
            # call it was. Printed AFTER the FAIL line and never in place of it; a
            # classifier that itself threw would still leave the failure recorded.
            if (($mode -eq 'Bulk') -and ($stage -eq 'RAISED')) {
                try   { Write-Output (New-BulkFailureLine -ErrorRecord $_) }
                catch { Write-Output ('BULKFAIL|' + (Get-BulkOp) + '|hresult=unclassified|the error could not be classified') }
            }
        }
    }
}

# --- THE COMPARISON, WHICH ONLY SPEAKS IF IT HAPPENED -----------------------
# `differ` IS RESERVED FOR A REAL COMPARISON. Run 1 printed it after a setup
# failure, which reads as "the two fixtures are not equivalent" and was not what
# happened at all.
$completed = @($passes.Keys)
if ($completed.Count -ne 2) {
    $why = 'no pass completed'
    if ($completed.Count -eq 1) { $why = 'only the ' + [string]$completed[0] + ' pass completed' }
    if ($setupFailed) { $why = 'the starting bundles were not prepared and proved identical' }
    Write-Output ('EQUIV|<not evaluated>|invalid|comparison was not executed: ' + $why)
} else {
    $reference = $passes['Endpoints']
    $optimised = $passes['Bulk']
    foreach ($field in @($reference.State.Keys)) {
        $left = [string]$reference.State[$field]
        $right = '<absent>'
        if ($optimised.State.Contains($field)) { $right = [string]$optimised.State[$field] }
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
    # CALC AND CALCEQUIV EXIST ONLY HERE, inside the branch where both passes
    # completed - so a setup failure cannot print a placeholder verdict on a
    # calculation that never ran.
    foreach ($pass in @($reference, $optimised)) {
        Write-Output ('CALC|' + $pass.Mode + '|' + $pass.CalcResult + '|' + $pass.CalcStatus +
                      '|' + $pass.CalcFingerprint)
    }
    if ($reference.CalcFingerprint -ceq $optimised.CalcFingerprint) {
        Write-Output 'CALCEQUIV|match|the production calculation fingerprint is identical'
    } else {
        Write-Output ('CALCEQUIV|differ|endpoints=' + $reference.CalcFingerprint +
                      ' :: bulk=' + $optimised.CalcFingerprint)
    }
}

# --- THE DISPOSABLE BUNDLES -------------------------------------------------
# Removed last, and only the ones that were created. A bundle whose pass failed
# is removed too: it is disposable by construction and the diagnosis is in the
# FAIL line, not in the directory.
foreach ($mode in @($bundles.Keys)) {
    Remove-Item -LiteralPath ([string]$bundles[$mode].Root) -Recurse -Force -ErrorAction SilentlyContinue
}
exit 0
