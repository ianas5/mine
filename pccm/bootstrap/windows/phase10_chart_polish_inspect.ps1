<#
.SYNOPSIS
    Reads the four Dashboard chart properties the final-delivery chart polish
    declares, from a workbook that already holds a published result, and prints
    them as CHART|<title>|<property>=<value> lines. Excel is opened by this
    script, the workbook is opened read-only, nothing is written and nothing is
    saved. A presentation review is done by a person looking at the Dashboard;
    this exports the properties so that review can be recorded exactly.

    IT IS ALSO A GATE, NOT ONLY A REPORT. Final acceptance of the chart package
    built at a2da277 found both year charts back from Excel with an EMPTY
    second SERIES argument - the applied-year VALUE names had survived and the
    applied-year CATEGORY name had not - which no amount of reading the chart
    XML would have shown. So this script now decides: it FAILS, with exit code
    1 and a CHART.FAIL line naming the chart and the reason, when a year chart
    plots blank categories, plots the whole reserved window, or plots any
    category source other than the applied-year calendar name.

.PARAMETER WorkbookPath
    The .xlsm to inspect - a Stage-B build, or the copy the presentation
    scenario was run in. The chart series formulas and axis settings are in the
    file whether or not a result is published; the Values/XValues counts are
    only meaningful after Calculate, Simulation and Annual Stochastic have run.

.PARAMETER ReportPath
    Optional. Where the lines are also written, UTF-8, LF.

.NOTES
    Reads only. No production endpoint is run, no cell is written, and the
    workbook is closed with SaveChanges:=False. Windows PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$WorkbookPath,
    [string]$ReportPath = ''
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# THE CONTRACT THIS GATE HOLDS THE WORKBOOK TO. The two year charts and the one
# category source they must plot. The name is the builder's, from the manifest's
# declared prefix and the calendar-year column key; it is spelled once here
# because a Windows gate cannot import the projection.
$script:YearCharts = @('Cumulative Cost Profile', 'Annual Cash Flow')
$script:CategoryName = 'chartAnnual_calendar_year'
$script:Lines = New-Object System.Collections.ArrayList
$script:Failures = New-Object System.Collections.ArrayList

function Emit-Line {
    param([string]$Text)
    $null = $script:Lines.Add($Text)
    Write-Host $Text
}
function Add-Failure {
    param([string]$Chart, [string]$Reason)
    $null = $script:Failures.Add($Chart + ': ' + $Reason)
    Emit-Line ('CHART.FAIL|' + $Chart + '|' + $Reason)
}
function Release-Object {
    param($Obj)
    if ($null -ne $Obj) {
        try { $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
    }
}

# THE SERIES FORMULA, SPLIT AT ITS TOP-LEVEL COMMAS. =SERIES(name, categories,
# values, order): a series NAME can carry a comma inside quotes and a reference
# can carry one inside brackets, so a naive split would misread both.
function Split-SeriesFormula {
    param([string]$Formula)
    $inner = $Formula
    $open = $inner.IndexOf('(')
    if ($open -lt 0) { return @() }
    $inner = $inner.Substring($open + 1, $inner.Length - $open - 2)
    $parts = New-Object System.Collections.ArrayList
    $depth = 0
    $quoted = $false
    $current = ''
    foreach ($ch in $inner.ToCharArray()) {
        if ($ch -eq '"') { $quoted = -not $quoted; $current = $current + $ch; continue }
        if ($quoted) { $current = $current + $ch; continue }
        if (($ch -eq '(') -or ($ch -eq '[')) { $depth = $depth + 1 }
        if (($ch -eq ')') -or ($ch -eq ']')) { $depth = $depth - 1 }
        if (($ch -eq ',') -and ($depth -eq 0)) { $null = $parts.Add($current.Trim()); $current = ''; continue }
        $current = $current + $ch
    }
    $null = $parts.Add($current.Trim())
    return @($parts)
}
function Get-SeriesField {
    param($Parts, [int]$Index)
    if (@($Parts).Count -le $Index) { return '' }
    return [string]@($Parts)[$Index]
}
function Get-PointCount {
    # HOW MANY POINTS THE SERIES ACTUALLY CARRIES - the applied years, once a
    # result is published - not how many rows the reserved window holds.
    param($Series)
    try { return [string]@($Series.Values).Count } catch { return '<unreadable>' }
}
# TRUE when a reference names the applied-year category, however Excel spells
# the sheet: Results!name and 'Results'!name are the same reference.
function Test-CategoryBinding {
    param([string]$Reference)
    $plain = $Reference.Replace("'", '')
    return ($plain -eq ('Results!' + $script:CategoryName))
}
function Test-CellRange {
    param([string]$Reference)
    return ($Reference -match '\$?[A-Z]{1,3}\$?\d+\s*:\s*\$?[A-Z]{1,3}\$?\d+')
}

$resolved = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = $null; $workbooks = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($resolved, 0, $true)
    Emit-Line ('WORKBOOK|' + [string]$workbook.FullName)

    $sheets = $null; $dashboard = $null; $results = $null; $names = $null
    try {
        $sheets = $workbook.Worksheets
        $dashboard = $sheets.Item('Dashboard')
        $results = $sheets.Item('Results')
        # THE APPLIED-YEAR NAMES, as Excel resolves them right now.
        $names = $results.Names
        $nameCount = [int]$names.Count
        $seenCategoryName = $false
        for ($n = 1; $n -le $nameCount; $n++) {
            $name = $null
            try {
                $name = $names.Item($n)
                $shortName = [string]$name.Name
                if ($shortName -like '*chartAnnual_*') {
                    $rowsNow = '<no range>'
                    try { $rowsNow = [string]$name.RefersToRange.Rows.Count } catch { $rowsNow = '<not a range now>' }
                    Emit-Line ('NAME|' + $shortName + '|refers_to=' + [string]$name.RefersTo + '|rows_now=' + $rowsNow)
                    if ($shortName -like ('*' + $script:CategoryName)) { $seenCategoryName = $true }
                }
            } finally { Release-Object $name }
        }
        if (-not $seenCategoryName) {
            Add-Failure 'Results' ('the applied-year category name ' + $script:CategoryName + ' is not defined on the sheet')
        }
        $objects = $null
        try {
            $objects = $dashboard.ChartObjects()
            $count = [int]$objects.Count
            Emit-Line ('CHARTS|count=' + [string]$count)
            $seenYearCharts = @()
            for ($i = 1; $i -le $count; $i++) {
                $object = $null; $chart = $null; $collection = $null
                try {
                    $object = $objects.Item($i)
                    $chart = $object.Chart
                    $title = '<untitled>'
                    try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = '<unreadable>' }
                    Emit-Line ('CHART|' + $title + '|type=' + [string]$chart.ChartType)
                    $isYearChart = ($script:YearCharts -contains $title)
                    if ($isYearChart) { $seenYearCharts = $seenYearCharts + $title }
                    $collection = $chart.SeriesCollection()
                    $seriesCount = [int]$collection.Count
                    if ($isYearChart -and ($seriesCount -lt 1)) {
                        Add-Failure $title 'plots no series at all'
                    }
                    for ($s = 1; $s -le $seriesCount; $s++) {
                        $series = $null
                        try {
                            $series = $collection.Item($s)
                            $formula = '<unreadable>'
                            try { $formula = [string]$series.Formula } catch { $formula = '<unreadable>' }
                            $parts = Split-SeriesFormula -Formula $formula
                            $categories = Get-SeriesField -Parts $parts -Index 1
                            $values = Get-SeriesField -Parts $parts -Index 2
                            $points = Get-PointCount -Series $series
                            $label = 'CHART|' + $title + '|series' + [string]$s
                            Emit-Line ($label + '.formula=' + $formula)
                            Emit-Line ($label + '.categories=' + $(if ($categories -eq '') { '<blank>' } else { $categories }))
                            Emit-Line ($label + '.values=' + $(if ($values -eq '') { '<blank>' } else { $values }))
                            Emit-Line ($label + '.points=' + $points)
                            if ($isYearChart) {
                                # THE GATE, AND IT IS THREE SEPARATE FAILURES. A blank
                                # categories argument is what Excel leaves behind when it
                                # drops the binding; a cell range is the reserved window
                                # coming back; anything else is a category source this
                                # workbook does not declare.
                                if ([string]::IsNullOrWhiteSpace($categories)) {
                                    Add-Failure $title ('series ' + [string]$s +
                                        ' has BLANK XValues/categories; Excel dropped the applied-year binding and is numbering the categories 1, 2, 3')
                                } elseif (Test-CellRange -Reference $categories) {
                                    Add-Failure $title ('series ' + [string]$s +
                                        ' plots the cell range ' + $categories +
                                        '; the whole reserved year window, not the applied years')
                                } elseif (-not (Test-CategoryBinding -Reference $categories)) {
                                    Add-Failure $title ('series ' + [string]$s +
                                        ' plots categories from ' + $categories + ', not Results!' + $script:CategoryName)
                                }
                                if ([string]::IsNullOrWhiteSpace($values)) {
                                    Add-Failure $title ('series ' + [string]$s + ' has BLANK values')
                                }
                            }
                        } finally { Release-Object $series }
                    }
                    # xlValue = 2, xlCategory = 1. On the horizontal bar chart the
                    # VALUE axis is the one drawn along the bottom.
                    $valueAxis = $null; $categoryAxis = $null
                    try {
                        $valueAxis = $chart.Axes(2)
                        Emit-Line ('CHART|' + $title + '|value_axis.min=' + [string]$valueAxis.MinimumScale + '|auto=' + [string]$valueAxis.MinimumScaleIsAuto)
                        Emit-Line ('CHART|' + $title + '|value_axis.max=' + [string]$valueAxis.MaximumScale + '|auto=' + [string]$valueAxis.MaximumScaleIsAuto)
                        Emit-Line ('CHART|' + $title + '|value_axis.major_unit=' + [string]$valueAxis.MajorUnit + '|auto=' + [string]$valueAxis.MajorUnitIsAuto)
                        Emit-Line ('CHART|' + $title + '|value_axis.number_format=' + [string]$valueAxis.TickLabels.NumberFormat)
                    } finally { Release-Object $valueAxis }
                    try {
                        $categoryAxis = $chart.Axes(1)
                        Emit-Line ('CHART|' + $title + '|category_axis.label_spacing=' + [string]$categoryAxis.TickLabelSpacing + '|auto=' + [string]$categoryAxis.TickLabelSpacingIsAuto)
                        Emit-Line ('CHART|' + $title + '|category_axis.number_format=' + [string]$categoryAxis.TickLabels.NumberFormat)
                    } finally { Release-Object $categoryAxis }
                } finally {
                    Release-Object $collection
                    Release-Object $chart
                    Release-Object $object
                }
            }
            foreach ($wanted in $script:YearCharts) {
                if (-not ($seenYearCharts -contains $wanted)) {
                    Add-Failure $wanted 'is not on the Dashboard at all'
                }
            }
        } finally { Release-Object $objects }
    } finally {
        Release-Object $names
        Release-Object $results
        Release-Object $dashboard
        Release-Object $sheets
    }
} finally {
    if ($null -ne $workbook) { try { $workbook.Close($false) } catch { }; Release-Object $workbook }
    Release-Object $workbooks
    if ($null -ne $excel) { try { $excel.Quit() } catch { }; Release-Object $excel }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
}
Emit-Line ''
if ($script:Failures.Count -eq 0) {
    Emit-Line 'CHART BINDING PASS'
} else {
    Emit-Line ('CHART BINDING FAIL: ' + [string]$script:Failures.Count + ' problem(s)')
    foreach ($problem in $script:Failures) { Emit-Line ('    ' + [string]$problem) }
}
if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($ReportPath, (($script:Lines -join "`n") + "`n"), $utf8)
    Write-Host ('Report written to ' + $ReportPath)
}
if ($script:Failures.Count -eq 0) { exit 0 } else { exit 1 }
