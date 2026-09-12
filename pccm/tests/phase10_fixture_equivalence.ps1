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
    runs the Stage-B bootstrap ONCE - build, SaveAs, modules, buttons, protection
    and the reopen verification, exactly as it ships - takes that verified .xlsm
    as the canonical baseline, makes two plain filesystem copies of it and proves
    all three SHA-256 digests equal, then builds PERF-SMALL BOTH WAYS in two
    separate Excel sessions over those copies, captures a full state snapshot
    from each, and compares them field for field - then runs the REAL
    `PCCM_Calculate` on both and compares production's own status and
    fingerprint. The only COM settlement in either pass is one bounded, read-only
    readiness barrier after the open and before anything is asked to change.

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
    Two Excel sessions, ONE Stage-B bootstrap, two disposable copies of its
    verified workbook. Neither copy is ever saved. Shutdown is the accepted
    com_lifecycle path.

    THE VOCABULARY IS UNAMBIGUOUS, because run 1 printed `PASS|Endpoints|RAISED`
    for a pass that never produced a workbook, and a reader could take that for a
    pass. PASS is now reserved for a pass that COMPLETED:

      BUNDLE|Baseline|<relative path>|<sha256> one required Stage-A artifact, as copied
      BASELINE|StageB|verified|sha256=<hash>   the one Stage-B build, bootstrap-verified
      COPY|<mode>|sha256=<hash>                one filesystem copy of it
      COPIES|identical                         canonical == Endpoints copy == Bulk copy
      COPIES|differ|<detail>                   they do not - nothing is opened
      READY|<mode>|attempt=N|waited=X          the read-only readiness barrier passed
      PASS|<mode>|COMPLETED|<detail>           this pass built and calculated
      FAIL|Baseline|SETUP|<detail>             the bundle or the copies could not be prepared
      FAIL|Baseline|BOOTSTRAP|<detail>         Stage-B did not produce a verified workbook
      FAIL|<mode>|RAISED|<detail>              anything after that pass's Excel started
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
# ===========================================================================
# ONE BASELINE, TWO COPIES
# ===========================================================================
# Runs 3 to 11 spent their failures inside a second Stage-B bootstrap, a second
# SaveAs, a second module import and a second reopen verification - none of which
# is the question. The question is whether two FIXTURE paths reach the same state
# from the same start, so the start is now one verified workbook, copied twice,
# with the three digests proved equal before either copy is opened.
function Get-WorkbookDigest {
    param([string]$Path)
    return [string](Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Copy-EquivalenceWorkbook {
    param([string]$Source, [string]$Root, [string]$Mode)
    $target = Join-Path $Root ('PCCM_equiv_' + $Mode.ToLower() + '.xlsm')
    if (Test-Path -LiteralPath $target) {
        throw ('the ' + $Mode + ' copy ' + $target + ' already exists')
    }
    Copy-Item -LiteralPath $Source -Destination $target
    if (-not (Test-Path -LiteralPath $target)) {
        throw ('the ' + $Mode + ' copy did not arrive at ' + $target)
    }
    return [pscustomobject]@{
        Mode   = $Mode
        Path   = $target
        Digest = (Get-WorkbookDigest -Path $target)
    }
}

function Test-CopyIdentity {
    param([string]$Canonical, $Copies)
    $problems = @()
    if ([string]::IsNullOrWhiteSpace($Canonical)) { $problems += 'the canonical workbook has no digest' }
    if (@($Copies).Count -ne 2) { $problems += ('expected two copies, found ' + [string]@($Copies).Count) }
    $paths = @()
    foreach ($copy in @($Copies)) {
        if ([string]$copy.Digest -cne $Canonical) {
            $problems += ('the ' + [string]$copy.Mode + ' copy differs: canonical=' + $Canonical +
                          ' copy=' + [string]$copy.Digest)
        }
        if ($paths -contains [string]$copy.Path) { $problems += ('two copies share the path ' + [string]$copy.Path) }
        $paths += [string]$copy.Path
    }
    return $problems
}

# ===========================================================================
# THE READINESS BARRIER - READ-ONLY, BOUNDED, ONCE PER PASS
# ===========================================================================
# The Stage-B bootstrap proved on Windows that a freshly opened workbook can
# answer a read with nothing for a moment. This is the one settlement each pass
# is allowed: after Workbooks.Open and before PCCM_AutomationBegin, the shim
# import, the window and any write, it reads FullName (which must be THIS copy),
# Worksheets, and one known worksheet. Every read goes through the accepted
# Invoke-ComRetryRead envelope for refusals; an answer of nothing waits 250 ms,
# doubling to 2000 ms, at most 12 attempts and 15000 ms in all; a real answer that
# is the wrong workbook aborts at once, because waiting cannot change which
# workbook this is. Nothing here writes, runs a macro or retries a mutation.
function Get-EquivalenceComparablePath {
    param([string]$Path)
    $full = $Path
    try { $full = [System.IO.Path]::GetFullPath($Path) } catch { $full = $Path }
    return $full.TrimEnd([char]92, [char]47).ToLowerInvariant()
}

function Wait-EquivalenceWorkbookReady {
    param($Workbook, [string]$ExpectedPath, [string]$KnownSheet,
          [int]$MaxAttempts = 12, [int]$FirstDelayMs = 250,
          [int]$MaxDelayMs = 2000, [int]$TotalBudgetMs = 15000)
    if ($null -eq $Workbook) { throw 'READY: no workbook to wait for.' }
    if ([string]::IsNullOrWhiteSpace($ExpectedPath)) { throw 'READY: no expected path.' }
    if ([string]::IsNullOrWhiteSpace($KnownSheet)) { throw 'READY: no known worksheet to acquire.' }
    $attempt = 0
    $waitedMs = 0
    $delay = $FirstDelayMs
    while ($true) {
        $attempt = $attempt + 1
        $missing = ''
        $nameRead = Invoke-ComRetryRead -Target $Workbook -Member 'FullName' `
            -Description 'the opened workbook FullName'
        $name = $nameRead.Value
        if (($null -eq $name) -or [string]::IsNullOrWhiteSpace([string]$name)) {
            $missing = 'FullName answered with nothing'
        } elseif ((Get-EquivalenceComparablePath ([string]$name)) -ne
                  (Get-EquivalenceComparablePath $ExpectedPath)) {
            throw ('READY: the opened workbook is bound to ' + [string]$name + ' and not to ' +
                   $ExpectedPath + '. Waiting cannot change which workbook this is.')
        }
        if ($missing -eq '') {
            $sheetsRead = Invoke-ComRetryRead -Target $Workbook -Member 'Worksheets' `
                -Description 'the opened workbook Worksheets'
            if ($null -eq $sheetsRead.Value) {
                $missing = 'Worksheets answered with nothing'
            } else {
                $sheetRead = Invoke-ComRetryRead -Target $sheetsRead.Value -Member 'Item' -Key $KnownSheet `
                    -Description ('the worksheet ' + $KnownSheet)
                if ($null -eq $sheetRead.Value) {
                    $missing = ('Worksheets.Item(' + $KnownSheet + ') answered with nothing')
                }
            }
        }
        if ($missing -eq '') {
            return [pscustomobject]@{ Attempt = $attempt; WaitedMs = $waitedMs; FullName = [string]$name }
        }
        if (($attempt -ge $MaxAttempts) -or (($waitedMs + $delay) -gt $TotalBudgetMs)) {
            throw ('READY: the opened workbook was not ready after ' + [string]$attempt +
                   ' attempt(s) and ' + [string]$waitedMs + ' ms: ' + $missing)
        }
        # ONLY AFTER AN OBSERVED NO-ANSWER, never unconditionally.
        Start-Sleep -Milliseconds $delay
        $waitedMs = $waitedMs + $delay
        $delay = [Math]::Min($delay * 2, $MaxDelayMs)
    }
}

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
    param([string]$Mode, [string]$WorkbookPath, $Manifest, $Inspection, $SimInspection, $Plan)
    if (-not (Test-Path -LiteralPath $WorkbookPath)) {
        throw ('RAISED: the ' + $Mode + ' copy is not at ' + $WorkbookPath)
    }
    $excel = $null; $workbooks = $null; $wb = $null
    $rel = New-ReleaseLedger ('equivalence ' + $Mode)
    $snapshot = $null; $calcStatus = ''; $calcFingerprint = ''; $calcResult = ''
    # ASSIGN, NEVER EMIT. This function returns exactly ONE object. Run 12 crashed
    # at the comparison because a READY line was written to the output stream
    # from inside here, and the caller then held an array where it expected the
    # result record. Readiness and any shutdown note travel IN the record.
    $ready = $null
    $shutdown = New-Object System.Collections.ArrayList
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AskToUpdateLinks = $false
        $workbooks = $excel.Workbooks
        $wb = $workbooks.Open($WorkbookPath)

        # THE ONE SETTLEMENT, before anything is asked to change.
        $ready = Wait-EquivalenceWorkbookReady -Workbook $wb -ExpectedPath $WorkbookPath `
            -KnownSheet ([string](@($Manifest.registers)[0].sheet))

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
        # THE WINDOW IS LABELLED FOR THE BULK BUILDER, so a refusal in P10FW_Begin
        # or P10FW_End is named as that and not as the first or last fixture step.
        if ($Mode -eq 'Bulk') { Reset-BulkOp; Set-BulkOp 'bulk.window.open' }
        $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $Manifest
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
        } catch { $null = $shutdown.Add('SHUTDOWN|' + $Mode + '|close raised: ' + $_.Exception.Message) }
        Invoke-NamedRelease $rel $wb         'Workbook';   $wb         = $null
        Invoke-NamedRelease $rel $workbooks  'Workbooks';  $workbooks  = $null
        try {
            if ($null -ne $excel) { $excel.Quit() }
        } catch { $null = $shutdown.Add('SHUTDOWN|' + $Mode + '|quit raised: ' + $_.Exception.Message) }
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
        Ready = $ready
        Shutdown = @($shutdown)
        CalcResult = $calcResult
        CalcStatus = $calcStatus
        CalcFingerprint = $calcFingerprint
    }
}

# ===========================================================================
# ONE STAGE-B, TWO COPIES, TWO PASSES, AND THE COMPARISON
# ===========================================================================
$manifest      = Get-Content -LiteralPath (Join-Path $BuildDir 'stage_b_manifest.json') -Raw | ConvertFrom-Json
$inspection    = Get-Content -LiteralPath (Join-Path $BuildDir 'phase5_gate_b_inspection.json') -Raw | ConvertFrom-Json
$simInspection = Get-Content -LiteralPath (Join-Path $BuildDir 'phase6_gate_b_inspection.json') -Raw | ConvertFrom-Json
$plan          = Get-Content -LiteralPath (Join-Path $BuildDir 'phase10_benchmark_plan.json') -Raw | ConvertFrom-Json

$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss-fff')
$bundle = $null
$copies = @{}
$setupFailed = $false
try {
    # --- THE ONE STARTING BUNDLE, AND THE ONE STAGE-B BUILD ------------------
    $bundle = New-EquivalenceBundle -Mode 'Baseline' -Manifest $manifest -Stamp $stamp
    foreach ($relative in @($bundle.Digests.Keys)) {
        Write-Output ('BUNDLE|Baseline|' + $relative + '|' + [string]$bundle.Digests[$relative])
    }
    # THE BOOTSTRAP RUNS ONCE, exactly as it ships: build, SaveAs, modules, buttons,
    # protection, the reopen verification, and its own Excel shutdown.
    $bootstrap = Join-Path $windows 'build_stage_b.ps1'
    & $bootstrap -BuildDir ([string]$bundle.Root) -Force | Out-Null
    # THE EXIT CODE, NOT ONLY THE FILE. A build that saved the .xlsm and was then
    # refused at a later operation leaves the workbook on disk WITHOUT its modules,
    # its buttons or its protection - and Test-Path alone would let the passes open
    # it and report fixture results against a half-built workbook.
    $bootstrapExit = $LASTEXITCODE
    $stageB = [string]$bundle.StageB
    if ($bootstrapExit -ne 0) {
        throw ('BOOTSTRAP: the Stage-B bootstrap exited ' + [string]$bootstrapExit +
               '; its transcript above names the failing operation. A workbook may exist ' +
               'and be half-built, so nothing is copied and no pass is run.')
    }
    if (-not (Test-Path -LiteralPath $stageB)) {
        throw ('BOOTSTRAP: the Stage-B bootstrap produced no workbook at ' + $stageB)
    }
    $canonical = Get-WorkbookDigest -Path $stageB
    Write-Output ('BASELINE|StageB|verified|sha256=' + $canonical)

    # --- TWO PLAIN COPIES, PROVED IDENTICAL BEFORE EITHER IS OPENED ----------
    foreach ($mode in @('Endpoints', 'Bulk')) {
        $copies[$mode] = Copy-EquivalenceWorkbook -Source $stageB -Root ([string]$bundle.Root) -Mode $mode
        Write-Output ('COPY|' + $mode + '|sha256=' + [string]$copies[$mode].Digest)
    }
    $problems = @(Test-CopyIdentity -Canonical $canonical -Copies @($copies['Endpoints'], $copies['Bulk']))
    if ($problems.Count -gt 0) {
        $setupFailed = $true
        Write-Output ('COPIES|differ|' + ($problems -join '; '))
    } else {
        Write-Output 'COPIES|identical'
    }
} catch {
    $setupFailed = $true
    $detail = ($_.Exception.Message) -replace '\s+', ' '
    $stage = 'SETUP'
    if ($detail -like 'BOOTSTRAP:*') { $stage = 'BOOTSTRAP'; $detail = $detail.Substring(10).Trim() }
    Write-Output ('FAIL|Baseline|' + $stage + '|' + $detail)
}

$passes = @{}
if (-not $setupFailed) {
    foreach ($mode in @('Endpoints', 'Bulk')) {
        try {
            $result = Invoke-EquivalencePass -Mode $mode -WorkbookPath ([string]$copies[$mode].Path) `
                -Manifest $manifest -Inspection $inspection -SimInspection $simInspection -Plan $plan
            # ONE RECORD, WITH ITS STATE, or this pass is not comparable.
            if (($result -isnot [pscustomobject]) -or ($null -eq $result.PSObject.Properties['State'])) {
                throw ('RAISED: the ' + $mode + ' pass returned ' + [string]@($result).Count +
                       ' object(s) instead of one result record')
            }
            foreach ($note in @($result.Shutdown)) { Write-Output $note }
            Write-Output ('READY|' + $mode + '|attempt=' + [string]$result.Ready.Attempt +
                          '|waited=' + [string]$result.Ready.WaitedMs)
            $passes[$mode] = $result
            Write-Output ('PASS|' + $mode + '|COMPLETED|fixture built and PCCM_Calculate ran')
        } catch {
            # THE STAGE IS NAMED. A copy that could not be opened and a COM call that
            # raised inside Excel are different facts.
            $detail = ($_.Exception.Message) -replace '\s+', ' '
            $stage = 'RAISED'
            if ($detail -like 'BOOTSTRAP:*') { $stage = 'BOOTSTRAP'; $detail = $detail.Substring(10).Trim() }
            elseif ($detail -like 'RAISED:*') { $detail = $detail.Substring(7).Trim() }
            Write-Output ('FAIL|' + $mode + '|' + $stage + '|' + $detail)
            # THE EXACT BULK OPERATION, SEPARATELY. Printed AFTER the FAIL line and
            # never in place of it; a classifier that itself threw would still
            # leave the failure recorded.
            if (($mode -eq 'Bulk') -and ($stage -eq 'RAISED')) {
                try   { Write-Output (New-BulkFailureLine -ErrorRecord $_) }
                catch { Write-Output ('BULKFAIL|' + (Get-BulkOp) + '|hresult=unclassified|the error could not be classified') }
            }
        }
    }
}

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
if ($null -ne $bundle) {
    Remove-Item -LiteralPath ([string]$bundle.Root) -Recurse -Force -ErrorAction SilentlyContinue
}
exit 0
