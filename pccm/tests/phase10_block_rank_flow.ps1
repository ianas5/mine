<#
.SYNOPSIS
    PCCM test-only harness: THE FIXTURE BLOCKS ARE RECTANGULAR, executed.

.DESCRIPTION
    WHY THIS EXISTS. A Windows Bulk fixture run raised

      the bulk write for tblCostProfiling was handed a rank-1 array;
      Excel accepts only a rectangular two-dimensional block

    `New-BenchmarkWeightBlock` ended with `return $block` on a rank-2 object[,], and
    PowerShell ENUMERATES a multidimensional array into its elements in row-major
    order - so the caller's assignment received an Object[] of rows x years scalars.
    The register builder never had the defect because its matrix travels as a
    PROPERTY of a record, which the pipeline cannot enumerate.

    "What CLR type, what rank, what dimensions" is not a property of text. So this
    lifts the real builders out of `bootstrap/windows/phase10_benchmark.ps1` BY AST
    and asks the objects they hand back.

    Excel is never started and no workbook is opened.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      RANK|<case>|<clr type>|<rank>|<rows>|<cols>|<expected rows>|<expected cols>
      PROP|<case>|<what a bare return would have produced>|<what the record gives>
    Exit 0 always.
#>
param(
    [string]$Runner
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path (Split-Path -Parent $here) 'bootstrap/windows/phase10_benchmark.ps1'
}
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}
foreach ($name in @('Get-BenchmarkPermanentId', 'New-BenchmarkRegisterBlock',
                    'New-BenchmarkWeightBlock')) {
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

# --- THE FAKES ---------------------------------------------------------------
# The weight builder reads the grid's keyed rows back through Get-TableBody, because
# production's SyncRows owns their order. The fake returns exactly the keys the
# model declares, in order, with a blank reserved suffix.
$script:GridRows = @()
function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @($script:GridRows)
}

function New-FakeDrivers {
    param([int]$Count, [int]$Years, [string]$Prefix)
    $drivers = @()
    for ($i = 1; $i -le $Count; $i++) {
        $weights = @()
        for ($y = 0; $y -lt $Years; $y++) { $weights += [double](1 / $Years) }
        $drivers += [pscustomobject]@{
            permanent_id      = ($Prefix + '-' + ([string]$i).PadLeft(3, '0'))
            description       = ('GateB ' + [string]$i)
            quantity          = [double]$i
            min_value         = [double]1
            most_likely       = [double]2
            max_value         = [double]3
            probability       = [double]0.5
            currency          = 'GBP'
            inflation_profile = 'CPI'
            distribution      = 'TRIANGULAR'
            profile_weights   = $weights
        }
    }
    # COMMA-WRAPPED so a one-driver model stays an array - and therefore NOT
    # collected with @() at the call site, because @( ,@(x) ) double-wraps and the
    # builder would see one element that is itself the array. That is the same trap
    # the Phase-10 shape defects were made of.
    return ,$drivers
}

$costColumns = @('permanent_id', 'description', 'category', 'quantity', 'uom',
                 'unit_cost_min', 'unit_cost_most_likely', 'unit_cost_max',
                 'currency', 'inflation_profile', 'distribution')
$riskColumns = @('permanent_id', 'risk_name', 'category', 'probability',
                 'impact_min', 'impact_most_likely', 'impact_max',
                 'currency', 'inflation_profile', 'distribution')

function Show-Block {
    param([string]$Case, $Block, [int]$ExpectedRows, [int]$ExpectedCols)
    $type = '<null>'; $rank = -1; $rows = -1; $cols = -1
    if ($null -ne $Block) {
        $type = $Block.GetType().FullName
        try { $rank = [int]$Block.Rank } catch { $rank = -1 }
        if ($rank -eq 2) {
            $rows = [int]$Block.GetLength(0)
            $cols = [int]$Block.GetLength(1)
        } elseif ($rank -eq 1) {
            $rows = [int]$Block.Length
            $cols = -1
        }
    }
    Write-Output ('RANK|' + $Case + '|' + $type + '|' + [string]$rank + '|' + [string]$rows +
                  '|' + [string]$cols + '|' + [string]$ExpectedRows + '|' + [string]$ExpectedCols)
}

# --- REGISTERS ---------------------------------------------------------------
foreach ($c in @(
    @{ n='SMALL cost lines'; count=12;  years=5;  risk=$false; prefix='CL'; cols=$costColumns },
    @{ n='SMALL risks';      count=8;   years=5;  risk=$true;  prefix='RK'; cols=$riskColumns },
    @{ n='LARGE cost lines'; count=180; years=30; risk=$false; prefix='CL'; cols=$costColumns })) {
    $drivers = New-FakeDrivers -Count $c.count -Years $c.years -Prefix $c.prefix
    $register = [pscustomobject]@{ table_name = ('tbl' + $c.prefix); sheet = 'Registers'
                                   columns = $c.cols }
    $prepared = New-BenchmarkRegisterBlock -Register $register `
        -Counter ([pscustomobject]@{ prefix = ($c.prefix + '-'); pad_width = 3
                                     defined_name = ('PCCM_Counter_' + $c.prefix) }) `
        -Drivers $drivers -IsRisk ([bool]$c.risk)
    Show-Block -Case ('register ' + $c.n) -Block $prepared.Block `
        -ExpectedRows $c.count -ExpectedCols @($c.cols).Count
}

# --- PROFILING GRIDS ---------------------------------------------------------
foreach ($c in @(
    @{ n='SMALL cost profiling'; count=12;  years=5;  prefix='CL' },
    @{ n='SMALL risk profiling'; count=8;   years=5;  prefix='RK' },
    @{ n='LARGE cost profiling'; count=180; years=30; prefix='CL' })) {
    $drivers = New-FakeDrivers -Count $c.count -Years $c.years -Prefix $c.prefix
    # The grid body: the same keys, in order, then a blank reserved suffix.
    $body = @()
    foreach ($d in $drivers) { $body += ,@([string]$d.permanent_id, '', '') }
    for ($i = 0; $i -lt 4; $i++) { $body += ,@('', '', '') }
    $script:GridRows = $body
    $grid = [pscustomobject]@{ table_name = ('tblProfiling' + $c.prefix); sheet = 'Profiling'
                               fixed_columns = @('permanent_id', 'description') }
    $prepared = New-BenchmarkWeightBlock -Workbook $null -Grid $grid `
        -Drivers $drivers -Years $c.years
    Show-Block -Case ('weights ' + $c.n) -Block $prepared.Block `
        -ExpectedRows $c.count -ExpectedCols $c.years
    Write-Output ('PROP|' + $c.n + '|rows=' + [string]$prepared.Rows + '|cols=' +
                  [string]$prepared.Columns + '|keys=' + [string]@($prepared.Keys).Count)
}

# --- WHAT A BARE RETURN WOULD HAVE DONE -------------------------------------
# THE DEFECT, DEMONSTRATED. Not a claim about PowerShell: two functions, identical
# but for how they hand the SAME matrix back.
function Get-MatrixByEmit {
    $m = New-Object 'object[,]' 3, 4
    for ($r = 0; $r -lt 3; $r++) { for ($c = 0; $c -lt 4; $c++) { $m[$r, $c] = ($r * 4 + $c) } }
    return $m
}
function Get-MatrixByRecord {
    $m = New-Object 'object[,]' 3, 4
    for ($r = 0; $r -lt 3; $r++) { for ($c = 0; $c -lt 4; $c++) { $m[$r, $c] = ($r * 4 + $c) } }
    return [pscustomobject]@{ Block = $m }
}
$emitted = Get-MatrixByEmit
$recorded = (Get-MatrixByRecord).Block
Show-Block -Case 'bare return (the defect)' -Block $emitted -ExpectedRows 3 -ExpectedCols 4
Show-Block -Case 'record property (the fix)' -Block $recorded -ExpectedRows 3 -ExpectedCols 4

exit 0
