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
foreach ($name in @('Split-SeriesFormula', 'Get-SeriesField', 'Get-CategoryRows', 'Test-CellRange')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the inspector defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}
# The two script variables the lifted functions read, bound exactly as the
# inspector binds them.
$script:CategorySheet = 'Results'
$script:CategoryColumn = 'D'
$script:CategoryFirstRow = 279
$script:CategoryLastRow = 478

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

# THE YEAR-CHART CONTRACT, read exactly as the gate reads it: the categories
# must cover the first annual row through the row the published count implies.
function Read-Series {
    param([string]$Case, [string]$Formula, [int]$PublishedYears)
    $parts = Split-SeriesFormula -Formula $Formula
    $categories = Get-SeriesField -Parts $parts -Index 1
    $values = Get-SeriesField -Parts $parts -Index 2
    $wantFirst = $script:CategoryFirstRow
    $wantLast = $wantFirst
    if ($PublishedYears -ge 1) {
        $wantLast = $wantFirst + $PublishedYears - 1
        if ($wantLast -gt $script:CategoryLastRow) { $wantLast = $script:CategoryLastRow }
    }
    $verdict = 'accepted'
    if ([string]::IsNullOrWhiteSpace($categories)) { $verdict = 'blank-categories' }
    else {
        $rows = Get-CategoryRows -Reference $categories
        if (@($rows)[0] -eq 0) { $verdict = 'wrong-source' }
        elseif (@($rows)[0] -ne $wantFirst) { $verdict = 'wrong-first-row' }
        elseif (@($rows)[1] -ne $wantLast) { $verdict = 'wrong-last-row' }
        elseif ((@($rows)[1] -eq $script:CategoryLastRow) -and ($PublishedYears -lt ($script:CategoryLastRow - $script:CategoryFirstRow + 1))) {
            $verdict = 'whole-window'
        }
    }
    Emit $Case ('categories=' + $(if ($categories -eq '') { '<blank>' } else { $categories }) +
                '|values=' + $(if ($values -eq '') { '<blank>' } else { $values }) + '|verdict=' + $verdict)
}

# 1. THE DEFECT EXCEL RETURNED AT a2da277 AND AGAIN AT 10f5e62, verbatim.
Read-Series 'windows.blank.s-curve-nominal' '=SERIES("Cumulative Nominal",,Results!chartAnnual_cumulative_nominal,1)' 10
Read-Series 'windows.blank.cash-flow' '=SERIES("Annual Nominal",,Results!chartAnnual_annual_nominal,1)' 10
# 2. THE SHAPE THE RUNTIME BINDING PRODUCES, which Windows proved by assigning
#    XValues and reading the SERIES formula back. Both spellings Excel uses.
Read-Series 'bound.ten-years' '=SERIES("Cumulative Nominal",Results!$D$279:$D$288,Results!chartAnnual_cumulative_nominal,1)' 10
Read-Series 'bound.quoted' "=SERIES(`"Cumulative Nominal`",'Results'!`$D`$279:`$D`$288,'Results'!chartAnnual_cumulative_nominal,1)" 10
Read-Series 'bound.one-year' '=SERIES("Annual Nominal",Results!$D$279,Results!chartAnnual_annual_nominal,1)' 1
Read-Series 'bound.nothing-published' '=SERIES("Annual Nominal",Results!$D$279,Results!chartAnnual_annual_nominal,1)' 0
Read-Series 'bound.whole-window-published' '=SERIES("Annual Nominal",Results!$D$279:$D$478,Results!chartAnnual_annual_nominal,1)' 200
# 3. THE RESERVED WINDOW COMING BACK while less than that is published.
Read-Series 'regression.whole-window' '=SERIES("Cumulative Nominal",Results!$D$279:$D$478,Results!chartAnnual_cumulative_nominal,1)' 10
# 4. THE WRONG ROWS, either end.
Read-Series 'regression.wrong-first' '=SERIES("Annual Nominal",Results!$D$280:$D$289,Results!chartAnnual_annual_nominal,1)' 10
Read-Series 'regression.wrong-last' '=SERIES("Annual Nominal",Results!$D$279:$D$300,Results!chartAnnual_annual_nominal,1)' 10
# 5. A CATEGORY SOURCE THAT IS NOT THE CALENDAR-YEAR COLUMN AT ALL - including
#    the failed defined name, which must no longer be accepted.
Read-Series 'regression.named-category' '=SERIES("Annual Nominal",Results!chartAnnual_calendar_year,Results!chartAnnual_annual_nominal,1)' 10
Read-Series 'regression.other-column' '=SERIES("Annual Nominal",Results!$B$279:$B$288,Results!chartAnnual_annual_nominal,1)' 10
Read-Series 'regression.other-sheet' '=SERIES("Annual Nominal",Dashboard!$D$279:$D$288,Results!chartAnnual_annual_nominal,1)' 10
# 5b. A SERIES NAME CARRYING A COMMA, which a naive split would misread as the
#     category argument.
Read-Series 'parsing.comma-in-name' '=SERIES("Cumulative, Nominal",Results!$D$279:$D$288,Results!chartAnnual_cumulative_nominal,1)' 10
# 6. THE FIXED-POPULATION CHARTS, whose categories Windows also found blank at
#    10f5e62 and whose accepted source is an ordinary range.
Read-Literal 'literal.histogram-accepted' '=SERIES("Iterations",Results!$B$487:$B$506,Results!$F$487:$F$506,1)'
Read-Literal 'literal.tornado-accepted' '=SERIES("Rho",Results!$B$515:$B$524,Results!$D$515:$D$524,1)'
Read-Literal 'literal.windows.10f5e62' '=SERIES("Iterations",,Results!$F$487:$F$506,1)'
Read-Literal 'literal.name-instead-of-range' '=SERIES("Iterations",Results!chartAnnual_calendar_year,Results!$F$487:$F$506,1)'
