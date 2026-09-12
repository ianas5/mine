<#
.SYNOPSIS
    PCCM test-only harness: THE LAST DRIVER'S TRACE COLUMN, executed.

.DESCRIPTION
    WHY THIS EXISTS. Equivalence run 10 was the first run in which both passes
    completed, and it differed in exactly two fields: cost_profiling.body and
    risk_profiling.body - the profiling Description of the LAST Cost Line (CL-012)
    and of the LAST Risk (R-008) were blank after the Endpoints fixture and present
    after the Bulk fixture. Every other row, every weight and both calculation
    fingerprints matched.

    "Which fixture step writes the register Description, relative to the production
    Add that synchronises the grid", "does anything synchronise again after the
    last one", and "do both fixtures end with the same grids" are BEHAVIOUR. So this
    lifts the REAL accepted Gate-B fixture (`Set-Phase5Fixture` and its steps, the
    Add-and-prove helper, the driver and weight writers) out of
    `bootstrap/windows/phase5_gate_b_scenarios.ps1` by AST, the REAL bulk builder,
    block builders, model builder and the Endpoints resynchronisation out of
    `bootstrap/windows/phase10_benchmark.ps1` by AST, replaces every helper that
    touches Excel with a stand-in over an in-memory workbook, and replaces the
    production endpoints with an emulation of modDrivers.AddDriver and
    modProfiling.SyncRows that is written statement for statement from the VBA:
    the identifier is written, THEN the grid is rewritten in register order with the
    trace column copied from the register and every weight preserved by permanent
    ID. Excel is never started and no workbook is opened.

    Three scenarios, all on PERF-SMALL from the real plan:
      endpoints-nosync  the accepted fixture alone (the state Windows run 10 saw)
      endpoints         the accepted fixture followed by Invoke-BenchmarkEndpointsResync
      bulk              Set-BenchmarkBulkFixture

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      OPS|<scenario>|<production operations in order, comma-separated>
      DESCWRITE|<scenario>|<id>|write-after-add=<bool>|synced-after-write=<bool>
      DESC|<scenario>|<grid table>|<row>|<id>|<trace text>
      WEIGHTS|<scenario>|<grid table>|<id>|<year cells, comma-separated>
      MODELWEIGHTS|<grid table>|<id>|<the model's profile_weights, comma-separated>
      SUFFIX|<scenario>|<grid table>|<rows after the last keyed row>|<all blank>
      BODY|<scenario>|<grid table>|<every row joined with | and rows with ;; - the snapshot's own rendering>
      RESULT|<scenario>|<completed or RAISED>|<detail>
    Exit 0 always.
#>
param(
    [string]$Runner,
    [string]$GateB,
    [string]$BuildDir
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$pccmRoot = Split-Path -Parent $here
if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path $pccmRoot 'bootstrap/windows/phase10_benchmark.ps1'
}
if ([string]::IsNullOrWhiteSpace($GateB)) {
    $GateB = Join-Path $pccmRoot 'bootstrap/windows/phase5_gate_b_scenarios.ps1'
}
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

. (Join-Path $pccmRoot 'bootstrap/windows/com_lifecycle.ps1')

function Import-LiftedFunctions {
    param([string]$Path, [string[]]$Names, [string]$Label)
    $errors = $null; $tokens = $null
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($resolved, [ref]$tokens, [ref]$errors)
    if ($errors -and $errors.Count -gt 0) {
        Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the ' + $Label)
        exit 0
    }
    $found = @{}
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($Names -contains $fn.Name) { $found[$fn.Name] = $fn.Extent.Text }
    }
    $bodies = @()
    foreach ($name in $Names) {
        if (-not $found.ContainsKey($name)) {
            Write-Output ('MISSING|' + $name + ' is not defined in the ' + $Label)
            exit 0
        }
        $bodies += $found[$name]
    }
    return $bodies
}

# THE REAL FIXTURES. The accepted Gate-B fixture first, then the runner's own
# functions, then (below) the stand-ins - the same resolution order the runner has.
# Defined at SCRIPT scope: Invoke-Expression inside a function would define them
# in that function's scope and they would be gone when it returned.
foreach ($lifted in @(Import-LiftedFunctions -Path $GateB -Label 'accepted Gate-B harness' -Names @(
        'Set-Phase5Fixture', 'Invoke-Phase5FixtureSteps', 'Invoke-Phase5AddDriverAndRequireSuccess',
        'Write-Phase5Driver', 'Write-Phase5Weights'))) { Invoke-Expression $lifted }
foreach ($lifted in @(Import-LiftedFunctions -Path $Runner -Label 'runner' -Names @(
        'Get-IdColumnValues', 'Get-BenchmarkPermanentId', 'New-BenchmarkRegisterBlock',
        'New-BenchmarkWeightBlock', 'New-BenchmarkWeights', 'New-BenchmarkDriver', 'New-BenchmarkModel',
        'Get-BulkOpVocabulary', 'Set-BulkOp', 'Get-BulkOp', 'Reset-BulkOp',
        'Set-BenchmarkBulkFixture', 'Invoke-BenchmarkEndpointsResync'))) { Invoke-Expression $lifted }

# THE REAL CONTRACTS, not restatements of them.
$manifest   = Get-Content -LiteralPath (Join-Path $BuildDir 'stage_b_manifest.json') -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath (Join-Path $BuildDir 'phase5_gate_b_inspection.json') -Raw | ConvertFrom-Json
$plan       = Get-Content -LiteralPath (Join-Path $BuildDir 'phase10_benchmark_plan.json') -Raw | ConvertFrom-Json

$script:RegisterByKey = @{}
foreach ($register in @($manifest.registers)) { $script:RegisterByKey[[string]$register.key] = $register }
$script:CounterByRegister = @{}
foreach ($counter in @($manifest.counters)) { $script:CounterByRegister[[string]$counter.driver_register] = $counter }
$script:GridByRegister = @{}
foreach ($grid in @($manifest.grids)) {
    if ($null -ne $grid.driver_register) { $script:GridByRegister[[string]$grid.driver_register] = $grid }
}
$script:RegisterByTable = @{}
foreach ($register in @($manifest.registers)) { $script:RegisterByTable[[string]$register.table_name] = $register }

# --- THE IN-MEMORY WORKBOOK ------------------------------------------------------
$script:Tables = @{}
$script:Names  = @{}
$script:Seq    = @{}
$script:Ops    = New-Object System.Collections.ArrayList
$script:Writes = New-Object System.Collections.ArrayList

function New-BlankRow { param([int]$Width) $row = @(); for ($c = 0; $c -lt $Width; $c++) { $row += '' }; return ,$row }

function Reset-FakeWorkbook {
    param([int]$Years)
    $script:Tables = @{}
    $script:Names = @{}
    $script:Seq = @{}
    $script:Ops.Clear()
    $script:Writes.Clear()
    foreach ($register in @($manifest.registers)) {
        $rows = New-Object System.Collections.ArrayList
        for ($r = 0; $r -lt [int]$register.reserved_rows; $r++) { $null = $rows.Add((New-BlankRow -Width @($register.columns).Count)) }
        $script:Tables[[string]$register.table_name] = $rows
        $script:Seq[[string]$register.key] = 0
    }
    foreach ($grid in @($manifest.grids)) {
        $rows = New-Object System.Collections.ArrayList
        $width = @($grid.fixed_columns).Count + $Years
        for ($r = 0; $r -lt [int]$grid.reserved_rows; $r++) { $null = $rows.Add((New-BlankRow -Width $width)) }
        $script:Tables[[string]$grid.table_name] = $rows
    }
    $fx = $inspection.input_tables.fx_rates
    $rows = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt [int]$fx.locked_seed_rows; $r++) { $null = $rows.Add(@('SAR', [double]1)) }
    $script:Tables[[string]$fx.table_name] = $rows
    foreach ($counter in @($manifest.counters)) { $script:Names[[string]$counter.defined_name] = [double]$counter.initial }
    $script:Names[[string]$inspection.inputs.reporting_currency.defined_name] = 'SAR'
}

function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $out = @()
    if ($script:Tables.ContainsKey($TableName)) {
        foreach ($row in $script:Tables[$TableName]) { $out += ,@($row) }
    }
    return $out
}
function Set-NamedValue { param($Workbook, [string]$DefinedName, $Value) $script:Names[$DefinedName] = $Value }
function Get-NamedValue {
    param($Workbook, [string]$DefinedName)
    if ($script:Names.ContainsKey($DefinedName)) { return $script:Names[$DefinedName] }
    return $null
}
function Set-TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value)
    if (-not $script:Tables.ContainsKey($TableName)) { throw ('the fake workbook has no table ' + $TableName) }
    $rows = $script:Tables[$TableName]
    if (($RowIndex -lt 1) -or ($RowIndex -gt $rows.Count)) {
        throw ('row ' + [string]$RowIndex + ' of ' + $TableName + ' does not exist (' + [string]$rows.Count + ' rows)')
    }
    $row = @($rows[$RowIndex - 1])
    if (($ColumnIndex -lt 1) -or ($ColumnIndex -gt $row.Count)) {
        throw ('column ' + [string]$ColumnIndex + ' of ' + $TableName + ' does not exist')
    }
    if ($null -eq $Value) { $row[$ColumnIndex - 1] = '' } else { $row[$ColumnIndex - 1] = $Value }
    $rows[$RowIndex - 1] = $row
    $null = $script:Writes.Add([pscustomobject]@{ Table = $TableName; Row = $RowIndex; Column = $ColumnIndex; Width = 1; AfterOp = $script:Ops.Count })
}
function Add-BlankTableRow {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $rows = $script:Tables[$TableName]
    $width = 2
    if ($rows.Count -gt 0) { $width = @($rows[0]).Count }
    $null = $rows.Add((New-BlankRow -Width $width))
    return [int]$rows.Count
}
function Reset-Phase5FxTable {
    param($Workbook, $Inspection, $Seed)
    $fx = $Inspection.input_tables.fx_rates
    $rows = New-Object System.Collections.ArrayList
    for ($r = 0; $r -lt [int]$fx.locked_seed_rows; $r++) { $null = $rows.Add(@('SAR', [double]1)) }
    $script:Tables[[string]$fx.table_name] = $rows
}
function Get-Phase5LockedFxSeed { return @() }
function Set-Phase5InflationProfileMaster { param($Workbook, $Inspection, $Profiles) }
function Write-Phase5InflationRates { param($Workbook, $Manifest, $Model) }
function Assert-Phase5StructurallyCoherent { param($Excel, [string]$Stage) return '' }
function Clear-Phase5Registers {
    param($Excel, $Workbook, $Manifest)
    foreach ($register in @($Manifest.registers)) {
        $rows = $script:Tables[[string]$register.table_name]
        for ($r = 0; $r -lt $rows.Count; $r++) { $rows[$r] = New-BlankRow -Width @($register.columns).Count }
        $script:Seq[[string]$register.key] = 0
    }
    foreach ($counter in @($Manifest.counters)) { $script:Names[[string]$counter.defined_name] = [double]$counter.initial }
}
function Set-BenchmarkRegisterRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$MinimumRows)
    $rows = $script:Tables[$TableName]
    $width = @($rows[0]).Count
    while ($rows.Count -lt $MinimumRows) { $null = $rows.Add((New-BlankRow -Width $width)) }
    return [int]$rows.Count
}
function Set-BenchmarkRangeBlock {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$FirstRow, [int]$FirstColumn, $Block, [string]$Description)
    if (-not ($Block -is [System.Array]) -or ([int]$Block.Rank -ne 2)) {
        throw ('the bulk write for ' + $Description + ' was not handed a rank-2 array')
    }
    $r = [int]$Block.GetLength(0); $c = [int]$Block.GetLength(1)
    $rows = $script:Tables[$TableName]
    for ($i = 0; $i -lt $r; $i++) {
        $row = @($rows[$FirstRow - 1 + $i])
        for ($j = 0; $j -lt $c; $j++) {
            $v = $Block[$i, $j]
            if ($null -eq $v) { $v = '' }
            $row[$FirstColumn - 1 + $j] = $v
        }
        $rows[$FirstRow - 1 + $i] = $row
        $null = $script:Writes.Add([pscustomobject]@{ Table = $TableName; Row = ($FirstRow + $i); Column = $FirstColumn; Width = $c; AfterOp = $script:Ops.Count })
    }
}

# --- PRODUCTION, EMULATED FROM THE VBA --------------------------------------------
# modProfiling.SyncRows, statement for statement: snapshot the existing year cells
# by permanent ID; rewrite the grid in register order - column 1 the identifier,
# column 2 the register's trace text (COL_COST_LINES_DESCRIPTION / RISK_NAME),
# year cells preserved by ID including blanks, a new ID at the initial value; clear
# the tail.
function Sync-FakeRows {
    param([string]$RegisterKey)
    $register = $script:RegisterByKey[$RegisterKey]
    $grid = $script:GridByRegister[$RegisterKey]
    $columns = @($register.columns | ForEach-Object { [string]$_ })
    $traceName = 'description'
    if ($RegisterKey -eq 'risk_register') { $traceName = 'risk_name' }
    $traceCol = [array]::IndexOf($columns, $traceName) + 1
    $fixed = @($grid.fixed_columns).Count
    $initial = [double]$grid.year_initial_value
    $regRows = $script:Tables[[string]$register.table_name]
    $gridRows = $script:Tables[[string]$grid.table_name]
    $yearCols = @($gridRows[0]).Count - $fixed

    $held = @{}
    foreach ($row in $gridRows) {
        $key = [string]$row[0]
        if (($key.Length -gt 0) -and (-not $held.ContainsKey($key))) {
            $values = @()
            for ($c = 0; $c -lt $yearCols; $c++) { $values += ,$row[$fixed + $c] }
            $held[$key] = $values
        }
    }
    $writeRow = 0
    foreach ($regRow in $regRows) {
        $driverId = [string]$regRow[0]
        if ($driverId.Length -eq 0) { continue }
        $writeRow++
        if ($writeRow -gt $gridRows.Count) { $null = $gridRows.Add((New-BlankRow -Width ($fixed + $yearCols))) }
        $row = @($gridRows[$writeRow - 1])
        $row[0] = $driverId
        $row[1] = [string]$regRow[$traceCol - 1]
        for ($c = 0; $c -lt $yearCols; $c++) {
            if ($held.ContainsKey($driverId)) { $row[$fixed + $c] = ($held[$driverId])[$c] }
            else { $row[$fixed + $c] = $initial }
        }
        $gridRows[$writeRow - 1] = $row
    }
    for ($r = $writeRow; $r -lt $gridRows.Count; $r++) { $gridRows[$r] = New-BlankRow -Width ($fixed + $yearCols) }
}

# modDrivers.AddDriver: first genuinely free row (an unkeyed row holding data is an
# orphan and a refusal), allocate the next identifier, write ONLY the identifier,
# then SyncRows. The Description is not written by production.
function Add-FakeDriver {
    param([string]$RegisterKey)
    $register = $script:RegisterByKey[$RegisterKey]
    $counter = $script:CounterByRegister[$RegisterKey]
    $rows = $script:Tables[[string]$register.table_name]
    $target = 0
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $row = @($rows[$r])
        if ([string]$row[0] -ne '') { continue }
        $holdsData = $false
        for ($c = 1; $c -lt $row.Count; $c++) { if ([string]$row[$c] -ne '') { $holdsData = $true } }
        if ($holdsData) { throw ('Row ' + [string]($r + 1) + ' already contains data but has no permanent identifier.') }
        $target = $r + 1; break
    }
    $script:Seq[$RegisterKey] = [int]$script:Seq[$RegisterKey] + 1
    $newId = Get-BenchmarkPermanentId -Counter $counter -Sequence ([int]$script:Seq[$RegisterKey])
    $script:Names[[string]$counter.defined_name] = [double]$script:Seq[$RegisterKey]
    if ($target -eq 0) { $null = $rows.Add((New-BlankRow -Width @($register.columns).Count)); $target = $rows.Count }
    $row = @($rows[$target - 1]); $row[0] = $newId; $rows[$target - 1] = $row
    Sync-FakeRows -RegisterKey $RegisterKey
    return $newId
}

function Invoke-Phase5ProductionOperation {
    param($Excel, [string]$Operation, [string]$Stage, [string]$Argument = '', [switch]$WithArgument)
    $null = $script:Ops.Add($Operation)
    if ($Operation -eq 'PCCM_AddCostLine') { $id = Add-FakeDriver -RegisterKey 'cost_lines'; return ('OK|Cost Line ' + $id + ' added.') }
    if ($Operation -eq 'PCCM_AddRisk')     { $id = Add-FakeDriver -RegisterKey 'risk_register'; return ('OK|Risk ' + $id + ' added.') }
    if ($Operation -eq 'PCCM_ApplyTimeline') {
        # SetYearColumns with the entered timeline unchanged deletes and adds no
        # column; then SyncRows for both kinds, in that order.
        Sync-FakeRows -RegisterKey 'cost_lines'
        Sync-FakeRows -RegisterKey 'risk_register'
        return 'OK|Timeline applied.'
    }
    return ('OK|' + $Operation)
}

# --- THE SCENARIOS ------------------------------------------------------------------
function Get-Scenario { param([string]$Id) foreach ($s in @($plan.scenarios)) { if ([string]$s.id -eq $Id) { return $s } }; throw ('no scenario ' + $Id) }
$spec = Get-Scenario -Id 'PERF-SMALL'

function Write-GridReport {
    param([string]$Scenario, $Model)
    $syncOps = @('PCCM_AddCostLine', 'PCCM_AddRisk', 'PCCM_ApplyTimeline')
    Write-Output ('OPS|' + $Scenario + '|' + (@($script:Ops) -join ','))
    foreach ($pair in @(@{ key = 'cost_lines'; drivers = @($Model.cost_lines) },
                        @{ key = 'risk_register'; drivers = @($Model.risks) })) {
        $register = $script:RegisterByKey[$pair.key]
        $grid = $script:GridByRegister[$pair.key]
        $columns = @($register.columns | ForEach-Object { [string]$_ })
        $traceName = 'description'; $ownSync = 'PCCM_AddCostLine'
        if ($pair.key -eq 'risk_register') { $traceName = 'risk_name'; $ownSync = 'PCCM_AddRisk' }
        $traceCol = [array]::IndexOf($columns, $traceName) + 1
        $regRows = @(Get-TableBody -Workbook $null -SheetName $register.sheet -TableName $register.table_name)
        # For each driver: was its trace text written after the Add that keyed its row,
        # and did ANY synchronisation of this grid run after that write?
        for ($i = 0; $i -lt $pair.drivers.Count; $i++) {
            $id = [string]($pair.drivers[$i]).permanent_id
            $rowIndex = 0
            for ($r = 0; $r -lt $regRows.Count; $r++) { if ([string]$regRows[$r][0] -eq $id) { $rowIndex = $r + 1 } }
            # The last write that covered this row's trace cell: a Set-TableCell on
            # that column, or a rectangular block write spanning it.
            $write = $null
            foreach ($w in $script:Writes) {
                if (($w.Table -eq [string]$register.table_name) -and ($w.Row -eq $rowIndex) -and
                    ($w.Column -le $traceCol) -and (($w.Column + $w.Width - 1) -ge $traceCol)) { $write = $w }
            }
            # The (i+1)-th Add of this kind is the one that keyed this driver's row.
            $addIndex = -1; $seen = 0
            for ($k = 0; $k -lt $script:Ops.Count; $k++) {
                if ($script:Ops[$k] -eq $ownSync) { $seen++; if ($seen -eq ($i + 1)) { $addIndex = $k } }
            }
            $writeAfterAdd = ($null -ne $write) -and ($addIndex -ge 0) -and ($write.AfterOp -gt $addIndex)
            $syncedAfter = $false
            if ($null -ne $write) {
                for ($k = $write.AfterOp; $k -lt $script:Ops.Count; $k++) {
                    if (($script:Ops[$k] -eq $ownSync) -or ($script:Ops[$k] -eq 'PCCM_ApplyTimeline')) { $syncedAfter = $true }
                }
            }
            Write-Output ('DESCWRITE|' + $Scenario + '|' + $id + '|write-after-add=' + [string]$writeAfterAdd + '|synced-after-write=' + [string]$syncedAfter)
            Write-Output ('MODELWEIGHTS|' + [string]$grid.table_name + '|' + $id + '|' + ((@(($pair.drivers[$i]).profile_weights) | ForEach-Object { [string]$_ }) -join ','))
        }
        $gridRows = @(Get-TableBody -Workbook $null -SheetName $grid.sheet -TableName $grid.table_name)
        $fixed = @($grid.fixed_columns).Count
        $lastKeyed = 0
        for ($r = 0; $r -lt $gridRows.Count; $r++) {
            $row = @($gridRows[$r])
            if ([string]$row[0] -ne '') {
                $lastKeyed = $r + 1
                Write-Output ('DESC|' + $Scenario + '|' + [string]$grid.table_name + '|' + [string]($r + 1) + '|' + [string]$row[0] + '|' + [string]$row[1])
                $cells = @(); for ($c = $fixed; $c -lt $row.Count; $c++) { $cells += [string]$row[$c] }
                Write-Output ('WEIGHTS|' + $Scenario + '|' + [string]$grid.table_name + '|' + [string]$row[0] + '|' + ($cells -join ','))
            }
        }
        $allBlank = $true
        for ($r = $lastKeyed; $r -lt $gridRows.Count; $r++) { foreach ($v in @($gridRows[$r])) { if ([string]$v -ne '') { $allBlank = $false } } }
        Write-Output ('SUFFIX|' + $Scenario + '|' + [string]$grid.table_name + '|' + [string]($gridRows.Count - $lastKeyed) + '|' + [string]$allBlank)
        # THE LOWEST GRID COLUMN ANY FIXTURE WROTE. The identifier and the trace
        # column are production's; a fixture that wrote either would show here.
        $lowest = 0
        foreach ($w in $script:Writes) {
            if (($w.Table -eq [string]$grid.table_name) -and (($lowest -eq 0) -or ($w.Column -lt $lowest))) { $lowest = [int]$w.Column }
        }
        Write-Output ('GRIDWRITES|' + $Scenario + '|' + [string]$grid.table_name + '|lowest-column=' + [string]$lowest + '|fixed=' + [string]$fixed)
        $lines = @()
        foreach ($row in $gridRows) { $lines += ((@($row) | ForEach-Object { [string]$_ }) -join '|') }
        Write-Output ('BODY|' + $Scenario + '|' + [string]$grid.table_name + '|' + ($lines -join ' ;; '))
    }
}

foreach ($scenario in @('endpoints-nosync', 'endpoints', 'bulk')) {
    Reset-FakeWorkbook -Years ([int]$spec.years)
    $model = New-BenchmarkModel -ScenarioSpec $spec
    $caught = $null
    try {
        if ($scenario -eq 'bulk') {
            Reset-BulkOp
            Set-BulkOp 'bulk.window.open'
            try {
                $null = Set-BenchmarkBulkFixture -Excel $null -Workbook $null -Manifest $manifest `
                    -Inspection $inspection -Model $model -ScenarioSpec $spec
            } finally { Set-BulkOp 'bulk.window.close' }
        } else {
            $null = Set-Phase5Fixture -Excel $null -Workbook $null -Manifest $manifest `
                -Inspection $inspection -Model $model
            if ($scenario -eq 'endpoints') {
                # EXACTLY THE RUNNER'S ENDPOINTS BRANCH: the accepted fixture, then
                # the lifted resynchronisation, nothing between and nothing after.
                $null = Invoke-BenchmarkEndpointsResync -Excel $null
            }
        }
    } catch { $caught = $_ }
    Write-GridReport -Scenario $scenario -Model $model
    if ($null -eq $caught) { Write-Output ('RESULT|' + $scenario + '|completed|' + [string]$script:Ops.Count + ' production operations') }
    else { Write-Output ('RESULT|' + $scenario + '|RAISED|' + [string]$caught.Exception.Message) }
}

exit 0
