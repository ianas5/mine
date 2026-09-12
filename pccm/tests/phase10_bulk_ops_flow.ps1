<#
.SYNOPSIS
    PCCM test-only harness: WHICH BULK FIXTURE CALL WAS REFUSED, executed.

.DESCRIPTION
    WHY THIS EXISTS. Two consecutive Windows runs got through readiness, SaveAs, the
    Stage-B build and the reopen verification, and then Bulk fixture construction
    raised RPC_E_CALL_REJECTED with nothing saying WHICH of its several dozen COM and
    VBA calls it was. The fixture now labels each call from a closed vocabulary, and
    the equivalence gate prints the label beside the failure.

    "Which label was in flight at each helper call", "does a failure name THAT
    operation", "can a LATER operation be reported once an earlier one failed", and
    "is anything called twice" are BEHAVIOUR. So this lifts the real orchestrator and
    the real block and model builders out of `bootstrap/windows/phase10_benchmark.ps1`
    BY AST, replaces every helper that touches Excel with a recording stand-in, feeds
    it the REAL manifest, inspection and plan from the build directory, and drives
    it - completing, growing, and failing at chosen operations.

    Excel is never started and no workbook is opened.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      BULKOPS|<n>|<comma-separated vocabulary>
      LABELCHECK|<case>|<outcome>
      ORDER|<scenario>|<labels in order of first helper touch, comma-separated>
      CALLS|<scenario>|<label>|<helper calls under that label>
      GROW|<scenario>|<table>|<ListRows.Add calls>
      RANK|<scenario>|<table>|<rank>|<rows>|<cols>
      RESULT|<scenario>|<completed or RAISED>|<detail>
      FAILAT|<label>|<kind>|<BULKFAIL line>|<labels touched after the failure>|<window closed>
      PREFIX|<gate or runner>|<labels in order of first touch, through the lifted real prefix region>
      PREFIXFAIL|<gate or runner>|<label>|<BULKFAIL line>|after=<labels touched after>|touches=<n>|opened=<window opened>|closed=<window closed>|message=<the exception message>
    Exit 0 always.
#>
param(
    [string]$Runner,
    [string]$Gate,
    [string]$BuildDir
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pccmRoot = Split-Path -Parent $here
if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path $pccmRoot 'bootstrap/windows/phase10_benchmark.ps1'
}
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }
if ([string]::IsNullOrWhiteSpace($Gate)) { $Gate = Join-Path $pccmRoot 'tests/phase10_fixture_equivalence.ps1' }
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

# The accepted classifier comes from the real file; nothing about it is restated.
. (Join-Path $pccmRoot 'bootstrap/windows/com_lifecycle.ps1')

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}
foreach ($name in @('Get-BulkOpVocabulary', 'Set-BulkOp', 'Get-BulkOp', 'Get-BulkComHResult',
                    'Format-BulkFailureLine', 'Save-BulkFailure', 'New-BulkFailureLine',
                    'Reset-BulkOp', 'Get-BenchmarkPermanentId', 'New-BenchmarkRegisterBlock',
                    'New-BenchmarkWeightBlock', 'Get-IdColumnValues', 'New-BenchmarkWeights',
                    'New-BenchmarkDriver', 'New-BenchmarkModel', 'Set-BenchmarkBulkFixture')) {
    $body = $null
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($fn.Name -eq $name) { $body = $fn.Extent.Text }
    }
    if ($null -eq $body) {
        Write-Output ('MISSING|' + $name + ' is not defined in the runner')
        exit 0
    }
    Invoke-Expression $body
}

# THE REAL CONTRACTS, not restatements of them.
$manifest   = Get-Content -LiteralPath (Join-Path $BuildDir 'stage_b_manifest.json') -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath (Join-Path $BuildDir 'phase5_gate_b_inspection.json') -Raw | ConvertFrom-Json
$plan       = Get-Content -LiteralPath (Join-Path $BuildDir 'phase10_benchmark_plan.json') -Raw | ConvertFrom-Json

# --- THE RECORDING STAND-INS -------------------------------------------------
# Every helper that would touch Excel records the label in flight at the moment it
# is called, and can be told to raise there. The workbook is a hashtable of tables.
$script:Trace    = New-Object System.Collections.ArrayList
$script:Tables   = @{}
$script:Adds     = @{}
$script:Ranks    = @{}
$script:FailAt   = ''
$script:FailKind = 'refused'
$script:WindowClosed = $false

function Note-Op {
    param([string]$Helper)
    $label = Get-BulkOp
    $null = $script:Trace.Add([pscustomobject]@{ Label = $label; Helper = $Helper })
    if (($script:FailAt -ne '') -and ($label -eq $script:FailAt)) {
        if ($script:FailKind -eq 'accepted') {
            throw (New-Object System.Runtime.InteropServices.COMException 'Exception from HRESULT: 0x800A03EC', -2146827284)
        }
        throw (New-Object System.Runtime.InteropServices.COMException 'Call was rejected by callee.', -2147418111)
    }
}

function Reset-FakeWorkbook {
    param([int]$Years)
    $script:Trace.Clear()
    $script:Tables = @{}
    $script:Adds = @{}
    $script:Ranks = @{}
    $script:WindowClosed = $false
    foreach ($register in @($manifest.registers)) {
        $rows = New-Object System.Collections.ArrayList
        for ($r = 0; $r -lt [int]$register.reserved_rows; $r++) {
            $row = @(); for ($c = 0; $c -lt @($register.columns).Count; $c++) { $row += '' }
            $null = $rows.Add($row)
        }
        $script:Tables[[string]$register.table_name] = $rows
        $script:Adds[[string]$register.table_name] = 0
    }
    foreach ($grid in @($manifest.grids)) {
        $rows = New-Object System.Collections.ArrayList
        $width = @($grid.fixed_columns).Count + $Years
        for ($r = 0; $r -lt [int]$grid.reserved_rows; $r++) {
            $row = @(); for ($c = 0; $c -lt $width; $c++) { $row += '' }
            $null = $rows.Add($row)
        }
        $script:Tables[[string]$grid.table_name] = $rows
    }
}

function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    Note-Op -Helper 'Get-TableBody'
    $out = @()
    foreach ($row in $script:Tables[$TableName]) { $out += ,@($row) }
    # PLAIN, NOT COMMA-WRAPPED. Every caller collects this with @(...), and
    # `return ,$out` + @(f) double-wraps: Get-IdColumnValues then iterated ONE row
    # that was the whole table and counted 11 keyed rows in a 12-driver fixture.
    # That is the trap the Phase-10 shape defects were made of, in the fake.
    return $out
}
function Set-NamedValue { param($Workbook, [string]$DefinedName, $Value) Note-Op -Helper 'Set-NamedValue' }
function Get-NamedValue { param($Workbook, [string]$DefinedName) Note-Op -Helper 'Get-NamedValue'; return 'SAR' }
function Set-TableCell { param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value) Note-Op -Helper 'Set-TableCell' }
function Add-BlankTableRow { param($Workbook, [string]$SheetName, [string]$TableName) Note-Op -Helper 'Add-BlankTableRow'; return 2 }
function Reset-Phase5FxTable { param($Workbook, $Inspection, $Seed) Note-Op -Helper 'Reset-Phase5FxTable' }
function Get-Phase5LockedFxSeed { return @() }
function Set-Phase5InflationProfileMaster { param($Workbook, $Inspection, $Profiles) Note-Op -Helper 'Set-Phase5InflationProfileMaster' }
function Write-Phase5InflationRates { param($Workbook, $Manifest, $Model) Note-Op -Helper 'Write-Phase5InflationRates' }
function Assert-Phase5StructurallyCoherent { param($Excel, [string]$Stage) Note-Op -Helper 'Assert-Phase5StructurallyCoherent'; return $true }

function Set-BenchmarkRegisterRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$MinimumRows)
    Note-Op -Helper 'Set-BenchmarkRegisterRowCount'
    $rows = $script:Tables[$TableName]
    $width = @($rows[0]).Count
    while ($rows.Count -lt $MinimumRows) {
        $row = @(); for ($c = 0; $c -lt $width; $c++) { $row += '' }
        $null = $rows.Add($row)
        $script:Adds[$TableName] = $script:Adds[$TableName] + 1
    }
    return [int]$rows.Count
}

function Set-BenchmarkRangeBlock {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$FirstRow, [int]$FirstColumn, $Block, [string]$Description)
    Note-Op -Helper 'Set-BenchmarkRangeBlock'
    $rank = -1; $r = -1; $c = -1
    if ($Block -is [System.Array]) {
        $rank = [int]$Block.Rank
        if ($rank -eq 2) { $r = [int]$Block.GetLength(0); $c = [int]$Block.GetLength(1) }
    }
    $script:Ranks[$TableName] = ([string]$rank + '|' + [string]$r + '|' + [string]$c)
    if ($rank -ne 2) { throw ('the bulk write for ' + $Description + ' was handed a rank-' + [string]$rank + ' array') }
    $rows = $script:Tables[$TableName]
    for ($i = 0; $i -lt $r; $i++) {
        $row = @($rows[$FirstRow - 1 + $i])
        for ($j = 0; $j -lt $c; $j++) {
            $v = $Block[$i, $j]
            if ($null -eq $v) { $v = '' }
            $row[$FirstColumn - 1 + $j] = $v
        }
        $rows[$FirstRow - 1 + $i] = $row
    }
}

# PRODUCTION'S OWN SYNCHRONISATION, MODELLED: ApplyTimeline keys each profiling grid
# from its register, one row per driver in register order, and grows the grid to fit.
function Invoke-Phase5ProductionOperation {
    param($Excel, [string]$Operation, [string]$Stage, [string]$Argument = '', [switch]$WithArgument)
    Note-Op -Helper ('Invoke-Phase5ProductionOperation:' + $Operation)
    if ($Operation -eq 'PCCM_ApplyTimeline') {
        foreach ($grid in @($manifest.grids)) {
            if ($null -eq $grid.driver_register) { continue }
            $register = @($manifest.registers | Where-Object { [string]$_.key -eq [string]$grid.driver_register })[0]
            $ids = @()
            foreach ($row in $script:Tables[[string]$register.table_name]) { if ([string]$row[0] -ne '') { $ids += [string]$row[0] } }
            $rows = $script:Tables[[string]$grid.table_name]
            $width = @($rows[0]).Count
            while ($rows.Count -lt $ids.Count) {
                $row = @(); for ($c = 0; $c -lt $width; $c++) { $row += '' }
                $null = $rows.Add($row)
            }
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $row = @($rows[$i])
                if ($i -lt $ids.Count) { $row[0] = $ids[$i] } else { $row[0] = '' }
                $rows[$i] = $row
            }
        }
    }
    return 'OK'
}

# --- THE VOCABULARY ------------------------------------------------------------
$vocab = @(Get-BulkOpVocabulary)
Write-Output ('BULKOPS|' + [string]$vocab.Count + '|' + ($vocab -join ','))
Reset-BulkOp
$outcome = 'accepted'
try { Set-BulkOp 'bulk.costprofiling.wrte' } catch { $outcome = 'REFUSED' }
Write-Output ('LABELCHECK|typo|' + $outcome + '|' + (Get-BulkOp))

function Get-Scenario { param([string]$Id) foreach ($s in @($plan.scenarios)) { if ([string]$s.id -eq $Id) { return $s } }; throw ('no scenario ' + $Id) }

# --- ONE PASS, IN THE GATE'S OWN SHAPE ------------------------------------------
# Window open, fixture inside a try, window close in a finally that RELABELS, and
# the failure line read only afterwards - exactly as Invoke-EquivalencePass does it.
function Invoke-FakeBulkPass {
    param([string]$ScenarioId)
    $spec = Get-Scenario -Id $ScenarioId
    Reset-FakeWorkbook -Years ([int]$spec.years)
    $model = New-BenchmarkModel -ScenarioSpec $spec
    Reset-BulkOp
    Set-BulkOp 'bulk.window.open'
    Note-Op -Helper 'Open-BenchmarkFixtureWindow'
    $caught = $null
    try {
        try {
            $null = Set-BenchmarkBulkFixture -Excel $null -Workbook $null -Manifest $manifest `
                -Inspection $inspection -Model $model -ScenarioSpec $spec
        } catch {
            Save-BulkFailure -ErrorRecord $_
            throw
        }
    } catch {
        $caught = $_
    } finally {
        Set-BulkOp 'bulk.window.close'
        $script:WindowClosed = $true
        $null = $script:Trace.Add([pscustomobject]@{ Label = (Get-BulkOp); Helper = 'Close-BenchmarkFixtureWindow' })
    }
    return $caught
}

# --- COMPLETE RUNS: SMALL AND THE LARGE GROWTH CASE ------------------------------
foreach ($id in @('PERF-SMALL', 'PERF-LARGE')) {
    $script:FailAt = ''
    $caught = Invoke-FakeBulkPass -ScenarioId $id
    $order = @(); $counts = @{}
    foreach ($entry in $script:Trace) {
        if ($order -notcontains $entry.Label) { $order += $entry.Label }
        if (-not $counts.ContainsKey($entry.Label)) { $counts[$entry.Label] = 0 }
        $counts[$entry.Label] = $counts[$entry.Label] + 1
    }
    Write-Output ('ORDER|' + $id + '|' + ($order -join ','))
    foreach ($label in $order) { Write-Output ('CALLS|' + $id + '|' + $label + '|' + [string]$counts[$label]) }
    foreach ($table in @($script:Adds.Keys | Sort-Object)) { Write-Output ('GROW|' + $id + '|' + $table + '|' + [string]$script:Adds[$table]) }
    foreach ($table in @($script:Ranks.Keys | Sort-Object)) { Write-Output ('RANK|' + $id + '|' + $table + '|' + [string]$script:Ranks[$table]) }
    if ($null -eq $caught) { Write-Output ('RESULT|' + $id + '|completed|' + [string]$script:Trace.Count + ' helper calls') }
    else { Write-Output ('RESULT|' + $id + '|RAISED|' + [string]$caught.Exception.Message) }
}

# --- FAILING AT CHOSEN OPERATIONS -----------------------------------------------
foreach ($case in @(
    @{ at = 'bulk.inputs.write';        kind = 'refused' },
    @{ at = 'bulk.cost.register.grow';  kind = 'refused' },
    @{ at = 'bulk.cost.register.write'; kind = 'refused' },
    @{ at = 'bulk.risk.counter.write';  kind = 'refused' },
    @{ at = 'bulk.timeline.apply';      kind = 'refused' },
    @{ at = 'bulk.costprofiling.write'; kind = 'refused' },
    @{ at = 'bulk.riskprofiling.acquire'; kind = 'refused' },
    @{ at = 'bulk.final.coherence';     kind = 'refused' },
    @{ at = 'bulk.timeline.apply';      kind = 'accepted' })) {
    $script:FailAt = [string]$case.at
    $script:FailKind = [string]$case.kind
    $caught = Invoke-FakeBulkPass -ScenarioId 'PERF-SMALL'
    $line = '<no failure>'
    if ($null -ne $caught) { $line = New-BulkFailureLine -ErrorRecord $caught }
    $seenAfter = @()
    $hit = $false
    foreach ($entry in $script:Trace) {
        if ($hit -and ($entry.Label -ne $script:FailAt) -and ($entry.Label -ne 'bulk.window.close')) { $seenAfter += $entry.Label }
        if ($entry.Label -eq $script:FailAt) { $hit = $true }
    }
    $touches = @($script:Trace | Where-Object { $_.Label -eq $script:FailAt }).Count
    Write-Output ('FAILAT|' + $script:FailAt + '|' + $script:FailKind + '|' + $line + '|after=' +
                  ($seenAfter -join ',') + '|touches=' + [string]$touches + '|closed=' + [string]$script:WindowClosed)
}

# =============================================================================
# THE PREFIX, LIFTED VERBATIM. Run 11 was refused before the first label. The
# gate's Bulk pass from its Reset-BulkOp through the window close, and the
# runner's Bulk prefix from its Reset-BulkOp through its catch, are cut out of
# the real files by their first and last statements and executed here over a
# fake Excel whose every property set, property get and Run records the label
# in flight and can be told to refuse there. New-Object itself is stood in for
# so that `New-Object -ComObject Excel.Application` hands back the fake.
# =============================================================================
function Get-LiftedRegion {
    param([string]$Path, [string]$StartMark, [string]$EndMark, [string]$Label)
    $text = Get-Content -LiteralPath $Path -Raw
    $from = $text.IndexOf($StartMark)
    $to = $text.IndexOf($EndMark)
    if (($from -lt 0) -or ($to -lt 0) -or ($to -lt $from)) {
        Write-Output ('MISSING|the ' + $Label + ' prefix region anchors are not in the file')
        exit 0
    }
    return $text.Substring($from, $to + $EndMark.Length - $from)
}
$gateRegion = Get-LiftedRegion -Path $Gate -Label 'gate' `
    -StartMark "        if (`$Mode -eq 'Bulk') { Reset-BulkOp }" `
    -EndMark ("            `$null = Close-BenchmarkFixtureWindow -Excel `$excel -Manifest `$Manifest`r`n        }")
$runnerRegion = Get-LiftedRegion -Path $runnerPath -Label 'runner' `
    -StartMark "    if (`$FixtureMode -eq 'Bulk') { Reset-BulkOp }" `
    -EndMark ("    } catch {`r`n        if (`$FixtureMode -eq 'Bulk') {`r`n            Save-BulkFailure -ErrorRecord `$_`r`n" +
              "            Write-BenchmarkLine ('  ' + (New-BulkFailureLine -ErrorRecord `$_))`r`n        }`r`n        throw`r`n    }")

$script:WindowOpened = $false
$script:Printed = New-Object System.Collections.ArrayList
function New-FakeExcel {
    $fake = Microsoft.PowerShell.Utility\New-Object PSObject
    foreach ($name in @('Visible', 'DisplayAlerts', 'AskToUpdateLinks')) {
        $setter = [scriptblock]::Create("Note-Op -Helper 'Excel.$name='; `$script:FakeExcelState['$name'] = `$args[0]")
        $getter = [scriptblock]::Create("return `$script:FakeExcelState['$name']")
        $fake | Add-Member -MemberType ScriptProperty -Name $name -Value $getter -SecondValue $setter
    }
    $fake | Add-Member -MemberType ScriptProperty -Name Workbooks -Value {
        Note-Op -Helper 'Excel.Workbooks'
        $books = Microsoft.PowerShell.Utility\New-Object PSObject
        $books | Add-Member -MemberType ScriptMethod -Name Open -Value { param($Path) Note-Op -Helper 'Workbooks.Open'; return 'workbook' }
        return $books
    }
    $fake | Add-Member -MemberType ScriptMethod -Name Run -Value {
        Note-Op -Helper ('Excel.Run:' + [string]$args[0]); return ('OK|' + [string]$args[0])
    }
    $fake | Add-Member -MemberType ScriptProperty -Name Hwnd -Value { return 1 }
    return $fake
}
$script:FakeExcelState = @{}
# STOOD IN LAST, so it wins: the COM request returns the fake, everything else
# goes to the real cmdlet.
function New-Object {
    param([Parameter(Position = 0)][string]$TypeName, [Parameter(Position = 1)]$ArgumentList,
          [string]$ComObject, $Property)
    if (-not [string]::IsNullOrEmpty($ComObject)) { Note-Op -Helper ('New-Object:' + $ComObject); return (New-FakeExcel) }
    # COMMA-WRAPPED: a collection or a rank-2 array returned bare is enumerated
    # by the pipeline - an empty ArrayList would come back as nothing at all.
    if ($null -ne $ArgumentList) { return ,(Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName -ArgumentList $ArgumentList) }
    if ($null -ne $Property) { return ,(Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName -Property $Property) }
    return ,(Microsoft.PowerShell.Utility\New-Object -TypeName $TypeName)
}
function Save-Phase5LockedFxSeed { param($Workbook, $Inspection) Note-Op -Helper 'Save-Phase5LockedFxSeed' }
function Import-BenchmarkFixtureWindow { param($Excel, $Workbook, $Manifest, [string]$ScriptDir) Note-Op -Helper 'Import-BenchmarkFixtureWindow' }
function Open-BenchmarkFixtureWindow { param($Excel, $Manifest) Note-Op -Helper 'Open-BenchmarkFixtureWindow'; $script:WindowOpened = $true; return 'opened' }
function Close-BenchmarkFixtureWindow { param($Excel, $Manifest) Note-Op -Helper 'Close-BenchmarkFixtureWindow'; $script:WindowClosed = $true; return 'closed' }
function Get-ExcelIdentity { param($ExcelApp, $PreExistingPids) Note-Op -Helper 'Get-ExcelIdentity'; return 'identity' }
function Get-BenchmarkEnvironment {
    param($Excel, $Identity, $WorkbookPath, $RepositoryPath, $ReleaseIdentity, $HarnessVersion, $SchemaVersion, $Revision)
    Note-Op -Helper 'Get-BenchmarkEnvironment'
    $out = @{}; foreach ($field in @($plan.environment_fields)) { $out[[string]$field] = 'x' }; return $out
}
function Assert-BenchmarkProtectionApplied { param($Excel, $Manifest, [string]$Stage) Note-Op -Helper 'Assert-BenchmarkProtectionApplied'; return @{ Raw = 'OK' } }
function Set-BenchmarkStage { param([string]$Stage, [string]$Action) }
function Write-BenchmarkLine { param([string]$Text) $null = $script:Printed.Add([string]$Text) }

function Invoke-LiftedPrefix {
    param([string]$Which)
    Reset-FakeWorkbook -Years 10
    $script:WindowOpened = $false
    $script:FakeExcelState = @{}
    $script:Printed.Clear()
    Reset-BulkOp
    # The lifted region's own variables, in both spellings.
    $Mode = 'Bulk'; $FixtureMode = 'Bulk'
    $Manifest = $manifest; $Inspection = $inspection; $Plan = $plan
    $stageB = 'fake.xlsm'; $stageBPath = 'fake.xlsm'
    $windows = $here; $scriptDir = $here; $repoRoot = $here
    $preExisting = @(); $releaseIdentity = 'release'; $revision = 'revision'
    $setupTimings = Microsoft.PowerShell.Utility\New-Object System.Collections.Specialized.OrderedDictionary
    $excel = $null; $workbooks = $null; $wb = $null
    $region = $gateRegion
    if ($Which -eq 'runner') { $region = $runnerRegion }
    $caught = $null
    try { Invoke-Expression $region } catch { $caught = $_ }
    return $caught
}

foreach ($which in @('gate', 'runner')) {
    $script:FailAt = ''
    $caught = Invoke-LiftedPrefix -Which $which
    $order = @()
    foreach ($entry in $script:Trace) { if ($order -notcontains $entry.Label) { $order += $entry.Label } }
    $suffix = 'completed'
    if ($null -ne $caught) { $suffix = 'RAISED|' + (([string]$caught.Exception.Message) -replace '\s+', ' ') }
    Write-Output ('PREFIX|' + $which + '|' + ($order -join ',') + '|' + $suffix)
}

$prefixCases = @{
    gate   = @('bulk.preflight.excel.create', 'bulk.preflight.excel.visible', 'bulk.preflight.excel.displayalerts',
               'bulk.preflight.excel.asktoupdatelinks', 'bulk.preflight.workbooks.acquire', 'bulk.preflight.workbook.open',
               'bulk.preflight.automation.begin', 'bulk.preflight.fxseed.read', 'bulk.preflight.window.import',
               'bulk.window.open')
    runner = @('bulk.preflight.excel.create', 'bulk.preflight.excel.identity', 'bulk.preflight.excel.visible',
               'bulk.preflight.excel.displayalerts', 'bulk.preflight.excel.asktoupdatelinks', 'bulk.preflight.workbooks.acquire',
               'bulk.preflight.workbook.open', 'bulk.preflight.environment.read', 'bulk.preflight.automation.begin',
               'bulk.preflight.fxseed.read', 'bulk.preflight.window.import', 'bulk.preflight.protection.read')
}
foreach ($which in @('gate', 'runner')) {
    foreach ($at in $prefixCases[$which]) {
        $script:FailAt = $at
        $script:FailKind = 'refused'
        $caught = Invoke-LiftedPrefix -Which $which
        $line = '<no failure>'
        if ($null -ne $caught) { $line = New-BulkFailureLine -ErrorRecord $caught }
        if ($which -eq 'runner') {
            # THE RUNNER PRINTS ITS OWN LINE in the catch; that is the one reported.
            # NOT `$printed`: variable names are case-insensitive and that would be
            # `$script:Printed` itself, reassigned to a fixed-size array.
            $runnerLines = @($script:Printed | Where-Object { $_ -like '*BULKFAIL|*' })
            if ($runnerLines.Count -eq 1) { $line = ([string]$runnerLines[0]).Trim() } else { $line = '<runner printed ' + [string]$runnerLines.Count + ' BULKFAIL lines>' }
        }
        $seenAfter = @(); $hit = $false
        foreach ($entry in $script:Trace) {
            if ($hit -and ($entry.Label -ne $at)) { $seenAfter += $entry.Label }
            if ($entry.Label -eq $at) { $hit = $true }
        }
        $touches = @($script:Trace | Where-Object { $_.Label -eq $at }).Count
        $message = '<none>'
        if ($null -ne $caught) { $message = ([string]$caught.Exception.Message) -replace '\s+', ' ' }
        Write-Output ('PREFIXFAIL|' + $which + '|' + $at + '|' + $line + '|after=' + ($seenAfter -join ',') +
                      '|touches=' + [string]$touches + '|opened=' + [string]$script:WindowOpened +
                      '|closed=' + [string]$script:WindowClosed + '|message=' + $message)
    }
}
$script:FailAt = ''

exit 0
