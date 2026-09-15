<#
.SYNOPSIS
    Executes the Windows chart-binding gate's PURE decision functions -
    Split-SeriesFormula, Test-CategoryBinding and Test-CellRange - lifted out of
    bootstrap/windows/phase10_chart_polish_inspect.ps1 by AST and driven over
    real Excel SERIES formulas under Set-StrictMode 2.0. Excel is never started.
    IT ASSERTS NOTHING: it prints GATE|<case>|<text> lines and the Python control
    decides.

    WHY THIS EXISTS. The gate is the only thing that can catch the defect it
    exists for - Excel dropping a category binding is invisible in the chart
    XML - so the gate's own reading of a SERIES formula has to be proved, and
    the formula Windows actually returned at a2da277 is one of the cases.
#>
param([string]$Inspector)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Inspector -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('Split-SeriesFormula', 'Get-SeriesField', 'Test-CategoryBinding', 'Test-CellRange')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the inspector defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}
# The two script variables the lifted functions read, bound exactly as the
# inspector binds them.
$script:CategoryName = 'chartAnnual_calendar_year'

function Emit { param([string]$Case, [string]$Text) Write-Output ('GATE|' + $Case + '|' + $Text) }
# THE LITERAL CONTRACT, read the same way: a fixed-population chart plots an
# ordinary cell range and must never come back blank either.
function Read-Literal {
    param([string]$Case, [string]$Formula)
    $parts = Split-SeriesFormula -Formula $Formula
    $categories = Get-SeriesField -Parts $parts -Index 1
    $verdict = 'accepted'
    if ([string]::IsNullOrWhiteSpace($categories)) { $verdict = 'blank-categories' }
    elseif (-not (Test-CellRange -Reference $categories)) { $verdict = 'not-a-range' }
    Emit $Case ('categories=' + $(if ($categories -eq '') { '<blank>' } else { $categories }) + '|verdict=' + $verdict)
}

function Read-Series {
    param([string]$Case, [string]$Formula)
    $parts = Split-SeriesFormula -Formula $Formula
    $categories = Get-SeriesField -Parts $parts -Index 1
    $values = Get-SeriesField -Parts $parts -Index 2
    $verdict = 'accepted'
    if ([string]::IsNullOrWhiteSpace($categories)) { $verdict = 'blank-categories' }
    elseif (Test-CellRange -Reference $categories) { $verdict = 'cell-range' }
    elseif (-not (Test-CategoryBinding -Reference $categories)) { $verdict = 'wrong-source' }
    Emit $Case ('categories=' + $(if ($categories -eq '') { '<blank>' } else { $categories }) +
                '|values=' + $(if ($values -eq '') { '<blank>' } else { $values }) + '|verdict=' + $verdict)
}

# 1. THE DEFECT EXCEL RETURNED AT a2da277, both year charts, verbatim.
Read-Series 'windows.a2da277.s-curve-nominal' '=SERIES("Cumulative Nominal",,Results!chartAnnual_cumulative_nominal,1)'
Read-Series 'windows.a2da277.s-curve-pv' '=SERIES("Cumulative PV",,Results!chartAnnual_cumulative_pv,2)'
Read-Series 'windows.a2da277.cash-flow' '=SERIES("Annual Nominal",,Results!chartAnnual_annual_nominal,1)'
# 2. THE SHAPE THE CORRECTION MUST PRODUCE, in both spellings Excel uses.
Read-Series 'bound.plain' '=SERIES("Cumulative Nominal",Results!chartAnnual_calendar_year,Results!chartAnnual_cumulative_nominal,1)'
Read-Series 'bound.quoted' "=SERIES(`"Cumulative Nominal`",'Results'!chartAnnual_calendar_year,'Results'!chartAnnual_cumulative_nominal,1)"
# 3. THE RESERVED WINDOW COMING BACK.
Read-Series 'regression.window-range' '=SERIES("Cumulative Nominal",Results!$D$279:$D$478,Results!chartAnnual_cumulative_nominal,1)'
Read-Series 'regression.short-range' '=SERIES("Annual Nominal",Results!$D$279:$D$288,Results!chartAnnual_annual_nominal,1)'
# 4. A CATEGORY SOURCE THIS WORKBOOK DOES NOT DECLARE.
Read-Series 'regression.other-name' '=SERIES("Annual Nominal",Results!chartAnnual_project_index,Results!chartAnnual_annual_nominal,1)'
Read-Series 'regression.other-sheet' '=SERIES("Annual Nominal",Dashboard!chartAnnual_calendar_year,Results!chartAnnual_annual_nominal,1)'
# 5. A SERIES NAME CARRYING A COMMA, which a naive split would misread as the
#    category argument - and a reference carrying one inside brackets.
Read-Series 'parsing.comma-in-name' '=SERIES("Cumulative, Nominal",Results!chartAnnual_calendar_year,Results!chartAnnual_cumulative_nominal,1)'
Read-Series 'parsing.comma-in-reference' '=SERIES("Annual Nominal",Results!chartAnnual_calendar_year,OFFSET(Results!$F$279,0,0,MAX(1,2),1),1)'
# 6. THE FIXED-POPULATION CHARTS, whose categories Windows also found blank at
#    10f5e62 and whose accepted source is an ordinary range.
Read-Literal 'literal.histogram-accepted' '=SERIES("Iterations",Results!$B$487:$B$506,Results!$F$487:$F$506,1)'
Read-Literal 'literal.tornado-accepted' '=SERIES("Rho",Results!$B$515:$B$524,Results!$D$515:$D$524,1)'
Read-Literal 'literal.windows.10f5e62' '=SERIES("Iterations",,Results!$F$487:$F$506,1)'
Read-Literal 'literal.name-instead-of-range' '=SERIES("Iterations",Results!chartAnnual_calendar_year,Results!$F$487:$F$506,1)'
