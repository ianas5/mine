<#
.SYNOPSIS
    Executes the runner's semantic post-reset verifiers - Get-FaSemanticBlockProblems,
    Get-FaTableBodyProblems, Get-FaResetProblems, ConvertTo-FaRectCells and
    Get-FaCellOffset - lifted out of the runner by AST and driven over plain data
    under Set-StrictMode 2.0. Excel is never started. IT ASSERTS NOTHING: it prints
    RESET|<case>|<count>|<problems> and RESET.SHAPE|<case>|count=..|nested=..|types=..
    lines and the Python control decides.

    RESTATED AT FINAL ACCEPTANCE RUN 17 (24edbcb). Every case now takes its result
    the way the runner's own callers do - `$x = @(<verifier> ...)` - and hands it
    to Emit WITHOUT a [string[]] coercion, which had masked a verifier emitting a
    no-problem result as one nested System.Object[] instead of zero objects. The
    shape line reports the count, whether any element is itself an array (checked
    recursively), and the element types. Get-FaResetProblems is composed over
    stubbed readers of plain data.
#>
param([string]$Runner)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('ConvertTo-FaRectCells', 'Get-FaSemanticBlockProblems', 'Get-FaTableBodyProblems', 'Get-FaCellOffset', 'Get-FaResetProblems')) {
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
# TRUE when any element, at any depth, is itself an array: the nested-collection defect.
function Test-Nested {
    param($Value)
    foreach ($item in @($Value)) {
        if ($item -is [System.Array]) { return $true }
    }
    return $false
}
# NO COERCION: $Problems arrives exactly as the caller holds it.
function Emit {
    param([string]$Case, $Problems)
    $items = @($Problems)
    $types = @()
    foreach ($item in $items) { if ($null -eq $item) { $types += 'null' } else { $types += $item.GetType().Name } }
    Write-Output ('RESET|' + $Case + '|' + [string]$items.Count + '|' + ($items -join '; '))
    Write-Output ('RESET.SHAPE|' + $Case + '|count=' + [string]$items.Count + '|nested=' + [string](Test-Nested $items) + '|types=' + ($types -join ','))
}

# TABLE BODIES: blank content with retained rows passes; one populated cell fails; geometry drift fails.
$blankFour = New-Body @(@('', '', ''), @('', '', ''), @('', '', ''), @('', '', ''))
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 4 -ExpectedColumns 3)
Emit 'table.blank-with-rows' $x
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body (New-Body @(@('', '', ''), @('', '2029', ''), @('', '', ''), @('', '', ''))) -ExpectedRows 4 -ExpectedColumns 3)
Emit 'table.one-populated' $x
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 5 -ExpectedColumns 3)
Emit 'table.rows-drifted' $x
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 4 -ExpectedColumns 4)
Emit 'table.columns-drifted' $x
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body @() -ExpectedRows 4 -ExpectedColumns 3)
Emit 'table.zero-rows-would-be-drift' $x
$x = @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body (New-Body @(@('1', '2', '3'), @('4', '5', '6'), @('7', '', ''))) -ExpectedRows 2 -ExpectedColumns 2)
Emit 'table.many-populated' $x

# CALC STATE C13:C20: NONE at the last-attempt field (row 17 = index 5), blank elsewhere.
$calcState = @('', '', '', '', 'NONE', '', '', '')
$x = @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells $calcState -SentinelIndex (Get-FaCellOffset -Address 'C13:C20' -Cell 'C17') -Sentinel 'NONE')
Emit 'calc.exact' $x
$x = @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', '', '', '', '', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.sentinel-missing' $x
$x = @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', '', '', '', 'SUCCESS', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.sentinel-wrong' $x
$x = @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', 'abc123', '', '', 'NONE', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.fingerprint-left' $x
$x = @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('a', 'b', '', '', 'NONE', '', '', 'c') -SentinelIndex 5 -Sentinel 'NONE')
Emit 'calc.three-left' $x

# SIM RECORD D23:D30: NONE at D23 (index 1), blank elsewhere; the active bank is the last field.
$simRecord = @('NONE', '', '', '', '', '', '', '')
$x = @(Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells $simRecord -SentinelIndex (Get-FaCellOffset -Address 'D23:D30' -Cell 'D23') -Sentinel 'NONE')
Emit 'sim.exact' $x
$x = @(Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells @('NONE', '', '', '', '', '', '', 'A') -SentinelIndex 1 -Sentinel 'NONE')
Emit 'sim.active-bank-populated' $x
$x = @(Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells @('', '', '', '', '', '', '', '') -SentinelIndex 1 -Sentinel 'NONE')
Emit 'sim.sentinel-blank' $x
$x = @(Get-FaSemanticBlockProblems -Label 'x' -Cells @('NONE', '') -SentinelIndex (Get-FaCellOffset -Address 'D23:D30' -Cell 'E23') -Sentinel 'NONE')
Emit 'offset.wrong-column' $x

# THE ASSEMBLY, exactly as Get-FaResetProblems accumulates: `$problems += @(...)`
# over the valid calc block, the valid sim record and a blank body must stay empty.
$assembled = @()
$assembled += @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells $calcState -SentinelIndex 5 -Sentinel 'NONE')
$assembled += @(Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells $simRecord -SentinelIndex 1 -Sentinel 'NONE')
$assembled += @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body $blankFour -ExpectedRows 4 -ExpectedColumns 3)
Emit 'assembled.valid' $assembled
$assembled = @()
$assembled += @(Get-FaSemanticBlockProblems -Label 'calculation state _Calc!C13:C20' -Cells @('', 'abc123', '', '', 'NONE', '', '', '') -SentinelIndex 5 -Sentinel 'NONE')
$assembled += @(Get-FaSemanticBlockProblems -Label 'simulation record _SimData!D23:D30' -Cells $simRecord -SentinelIndex 1 -Sentinel 'NONE')
$assembled += @(Get-FaTableBodyProblems -Label 'tblCalcYears' -Body (New-Body @(@('', '', ''), @('', '2029', ''), @('', '', ''), @('', '', ''))) -ExpectedRows 4 -ExpectedColumns 3)
Emit 'assembled.two-problems' $assembled

# THE RECT READER: a 2-D COM rectangle becomes a flat string list, nulls as blanks.
# ONE-BASED, as a COM Value2 rectangle is. Taken as the runner takes it: by assignment.
$rect = [Array]::CreateInstance([object], @(3, 1), @(1, 1))
$rect.SetValue('NONE', 1, 1); $rect.SetValue($null, 2, 1); $rect.SetValue(4, 3, 1)
$x = ConvertTo-FaRectCells -Rect $rect
Emit 'rect.flatten' $x

# THE COMPOSITION: Get-FaResetProblems over stubbed readers of plain data. The
# stubs stand where the COM readers stand and answer with the same shapes: a
# block's Rect is a one-based 2-D Value2 rectangle (nulls for blanks), a table
# body is one row object per row, an ordinary rectangle is blank when every
# cell is. The first state is the exact contracted post-reset state Windows
# observed at 24edbcb; the second leaves a fingerprint, a populated cell and a
# populated ordinary rectangle behind.
function New-Column {
    param([string[]]$Values)
    $out = [Array]::CreateInstance([object], @($Values.Count, 1), @(1, 1))
    for ($i = 1; $i -le $Values.Count; $i++) {
        if ($Values[$i - 1] -eq '') { $out.SetValue($null, $i, 1) } else { $out.SetValue($Values[$i - 1], $i, 1) }
    }
    return , $out
}
$script:Cells = @{}
$script:Bodies = @{}
function Get-FaClearedRectangles {
    param($Reset, [int]$Iterations)
    return @([pscustomobject]@{ Sheet = '_Calc'; Address = 'C13:C20' },
             [pscustomobject]@{ Sheet = '_Calc'; Address = 'E5:E9' },
             [pscustomobject]@{ Sheet = '_SimData'; Address = 'D23:D30' },
             [pscustomobject]@{ Sheet = '_SimData'; Address = 'B40:B45' })
}
function Get-FaBlock {
    param($Workbook, [string]$SheetName, [string]$Address)
    return [pscustomobject]@{ Sheet = $SheetName; Address = $Address; Rect = $script:Cells[($SheetName + '!' + $Address)] }
}
function Test-FaBlockBlank {
    param($Workbook, [string]$SheetName, [string]$Address)
    $cells = ConvertTo-FaRectCells -Rect ((Get-FaBlock -Workbook $Workbook -SheetName $SheetName -Address $Address).Rect)
    foreach ($cell in $cells) { if ($cell -ne '') { return $false } }
    return $true
}
function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    foreach ($row in @($script:Bodies[$TableName])) { Write-Output -NoEnumerate ([object[]]$row) }
}
$reset = [pscustomobject]@{
    publications = [pscustomobject]@{
        calculation = [pscustomobject]@{ sheet = '_Calc'; cleared = [pscustomobject]@{ state = 'C13:C20'; tables = @('tblCalcYears', 'tblCalcTotals') } }
        simulation  = [pscustomobject]@{ sheet = '_SimData'; cleared = [pscustomobject]@{ attempt_and_selector = 'D23:D30' } }
    }
}
$shapes = @{ 'tblCalcYears' = [pscustomobject]@{ Rows = 4; Columns = 3 }; 'tblCalcTotals' = [pscustomobject]@{ Rows = 1; Columns = 2 } }

$script:Cells = @{
    '_Calc!C13:C20'    = (New-Column @('', '', '', '', 'NONE', '', '', ''))
    '_Calc!E5:E9'      = (New-Column @('', '', '', '', ''))
    '_SimData!D23:D30' = (New-Column @('NONE', '', '', '', '', '', '', ''))
    '_SimData!B40:B45' = (New-Column @('', '', '', '', '', ''))
}
$script:Bodies = @{ 'tblCalcYears' = @(, @('', '', '')) * 4; 'tblCalcTotals' = @(, @('', '')) }
$x = @(Get-FaResetProblems -Workbook 'stub' -Reset $reset -Iterations 1000 -CalcAttemptCell 'C17' -SimAttemptCell 'D23' -Sentinel 'NONE' -TableShapes $shapes)
Emit 'reset.valid' $x

$script:Cells['_Calc!C13:C20'] = (New-Column @('', 'abc123', '', '', 'NONE', '', '', ''))
$script:Cells['_SimData!B40:B45'] = (New-Column @('', '', '0.42', '', '', ''))
$script:Bodies['tblCalcYears'] = @(@('', '', ''), @('', '2029', ''), @('', '', ''), @('', '', ''))
$x = @(Get-FaResetProblems -Workbook 'stub' -Reset $reset -Iterations 1000 -CalcAttemptCell 'C17' -SimAttemptCell 'D23' -Sentinel 'NONE' -TableShapes $shapes)
Emit 'reset.three-problems' $x
