<#
.SYNOPSIS
    Executes the runner's semantic post-reset verifiers - Get-FaSemanticBlockProblems,
    Get-FaTableBodyProblems, ConvertTo-FaRectCells and Get-FaCellOffset - lifted out
    of the runner by AST and driven over plain data under Set-StrictMode 2.0. Excel
    is never started. IT ASSERTS NOTHING: it prints RESET|<case>|<problems> lines
    and the Python control decides.
#>
param([string]$Runner)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('ConvertTo-FaRectCells', 'Get-FaSemanticBlockProblems', 'Get-FaTableBodyProblems', 'Get-FaCellOffset')) {
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
function Emit { param([string]$Case, [string[]]$Problems) Write-Output ('RESET|' + $Case + '|' + [string]@($Problems).Count + '|' + (@($Problems) -join '; ')) }

# TABLE BODIES: blank content with retained rows passes; one populated cell fails; geometry drift fails.
$blankFour = New-Body @(@('', '', ''), @('', '', ''), @('', '', ''), @('', '', ''))
Emit 'table.blank-with-rows' (Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 4 -ExpectedColumns 3)
Emit 'table.one-populated' (Get-FaTableBodyProblems -Label 'tblCalcYears' -Body (New-Body @(@('', '', ''), @('', '2029', ''), @('', '', ''), @('', '', ''))) -ExpectedRows 4 -ExpectedColumns 3)
Emit 'table.rows-drifted' (Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 5 -ExpectedColumns 3)
Emit 'table.columns-drifted' (Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 4 -ExpectedColumns 4)
Emit 'table.zero-rows-would-be-drift' (Get-FaTableBodyProblems -Label 'tblCalcYears' -Body @() -ExpectedRows 4 -ExpectedColumns 3)

# CALC STATE C13:C20: NONE at the last-attempt field (row 17 = index 5), blank elsewhere.
$calcState = @('', '', '', '', 'NONE', '', '', '')
Emit 'calc.exact' (Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells $calcState -SentinelIndex (Get-FaCellOffset -Address 'C13:C20' -Cell 'C17') -Sentinel 'NONE')
Emit 'calc.sentinel-missing' (Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', '', '', '', '', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.sentinel-wrong' (Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', '', '', '', 'SUCCESS', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.fingerprint-left' (Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', 'abc123', '', '', 'NONE', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')

# SIM RECORD D23:D30: NONE at D23 (index 1), blank elsewhere; the active bank is the last field.
Emit 'sim.exact' (Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells @('NONE', '', '', '', '', '', '', '') -SentinelIndex (Get-FaCellOffset -Address 'D23:D30' -Cell 'D23') -Sentinel 'NONE')
Emit 'sim.active-bank-populated' (Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells @('NONE', '', '', '', '', '', '', 'A') -SentinelIndex 1 -Sentinel 'NONE')
Emit 'sim.sentinel-blank' (Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells @('', '', '', '', '', '', '', '') -SentinelIndex 1 -Sentinel 'NONE')
Emit 'offset.wrong-column' (Get-FaSemanticBlockProblems -Label 'x' -Cells @('NONE', '') -SentinelIndex (Get-FaCellOffset -Address 'D23:D30' -Cell 'E23') -Sentinel 'NONE')

# THE RECT READER: a 2-D COM rectangle becomes a flat string list, nulls as blanks.
# ONE-BASED, as a COM Value2 rectangle is.
$rect = [Array]::CreateInstance([object], @(3, 1), @(1, 1))
$rect.SetValue('NONE', 1, 1); $rect.SetValue($null, 2, 1); $rect.SetValue(4, 3, 1)
Emit 'rect.flatten' (ConvertTo-FaRectCells -Rect $rect)
