<#
.SYNOPSIS
    Executes the runner's repair.grids-restored diagnosis - Format-FaGridDifferences
    and the two pure body helpers beneath it - lifted out of the runner by AST and
    driven over synthetic plain-data bodies under Set-StrictMode 2.0. Excel is never
    started. IT ASSERTS NOTHING: it prints one DIAG|<case>|<text> line per case and
    the Python control decides.
#>
param([string]$Runner)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('ConvertTo-FaBodyKey', 'Get-FaBodyCell', 'Get-FaBodyWidth', 'Format-FaGridDifferences')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the runner defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}

function New-Body {
    param([object[]]$Rows)
    $out = @()
    foreach ($row in $Rows) { $out += , @($row) }
    return , $out
}
$columns = @('Cost Line ID', 'Description', '2028', '2029', '2030', '2031')
$baseline = New-Body @(@('CL-001', 'Civils', '0.25', '0.25', '0.25', '0.25'),
                       @('CL-002', 'Rail', '0.5', '0.5', '0', '0'),
                       @('', '', '', '', '', ''))
$rowTwoBlankTail = New-Body @(@('CL-001', 'Civils', '0.25', '0.25', '0.25', '0.25'),
                              @('CL-002', 'Rail', '0.5', '0.5', '', ''),
                              @('', '', '', '', '', ''))
$rowTwoZeroed = New-Body @(@('CL-001', 'Civils', '0.25', '0.25', '0.25', '0.25'),
                           @('CL-002', 'Rail', '0.5', '0.5', '0', '0'),
                           @('', '', '0', '0', '0', '0'))
$narrow = New-Body @(@('CL-001', 'Civils', '0.25', '0.25', '0.25'),
                     @('CL-002', 'Rail', '0.5', '0.5', '0'),
                     @('', '', '', '', ''))
$narrowColumns = @('Cost Line ID', 'Description', '2028', '2029', '2030')
$manyDiffs = New-Body @(@('CL-001', 'Civils', '1', '1', '1', '1'),
                        @('CL-002', 'Rail', '1', '1', '1', '1'),
                        @('', '', '', '', '', ''))

function Emit { param([string]$Case, [string]$Text) Write-Output ('DIAG|' + $Case + '|' + $Text) }

Emit 'identical' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $baseline -AfterResync $baseline -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
Emit 'case1.blank-tail' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $rowTwoBlankTail -AfterResync $rowTwoBlankTail -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
Emit 'case2.zeroed-by-resync' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $baseline -AfterResync $rowTwoZeroed -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
Emit 'mixed' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $rowTwoBlankTail -AfterResync $rowTwoZeroed -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
Emit 'note.restored-by-resync' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $rowTwoBlankTail -AfterResync $baseline -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
Emit 'width.narrow-before-resync' (Format-FaGridDifferences -Label 'Risk Profiling!tblRiskProfiling' -Baseline $baseline -BeforeResync $narrow -AfterResync $baseline -BaselineColumns $columns -BeforeColumns $narrowColumns -AfterColumns $columns)
Emit 'limit.many' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $manyDiffs -AfterResync $manyDiffs -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns -Limit 2)
Emit 'rows.fewer-after-resync' (Format-FaGridDifferences -Label 'Cost Profiling!tblCostProfiling' -Baseline $baseline -BeforeResync $baseline -AfterResync (New-Body @(, @('CL-001', 'Civils', '0.25', '0.25', '0.25', '0.25'))) -BaselineColumns $columns -BeforeColumns $columns -AfterColumns $columns)
