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

# THE CONTRACT THIS GATE HOLDS THE WORKBOOK TO, AND IT IS TWO CONTRACTS.
#
# THE TWO YEAR CHARTS plot the applied years, and their categories are bound at
# RUNTIME by modChartPresentation - Excel returns a named category blank, and
# Windows proved Series.XValues instead. So their categories must be a RANGE on
# the bridge sheet that starts at the first annual row and ends at the row the
# published year count implies. Never a name, never blank, and never the whole
# reserved window.
#
# THE HISTOGRAM AND THE TORNADO plot fixed populations - the declared bin count
# and the top-N drivers - so their categories must be an ordinary cell RANGE,
# the accepted literal source every Windows round before the chart polish
# proved. A name there would mean the correction had wandered into a chart that
# never needed one.
#
# NEITHER may be blank. A blank category argument is what Excel leaves behind
# when it drops a binding, and it is the defect this gate exists for.
$script:YearCharts = @('Cumulative Cost Profile', 'Annual Cash Flow')
$script:LiteralCharts = @('Total Cost Distribution', 'Top Drivers by Rank Correlation')
# The bridge sheet, the first annual row and the reserved window, read from the
# workbook rather than assumed: the presentation owner's own window name is
# what says where the categories may start and how far they may reach.
$script:CategoryWindowName = 'chartAnnual_category_window'
$script:YearCountName = 'chartAnnual_year_count'
$script:CategoryFirstRow = 0
$script:CategoryLastRow = 0
$script:CategoryColumn = ''
$script:CategorySheet = ''
$script:PublishedYears = 0
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
# THE ROWS A CATEGORY REFERENCE COVERS, as Excel spells it back:
# Results!$D$279:$D$288 -> 279, 288. Zero rows when it is not a range on the
# bridge sheet's category column at all.
function Get-CategoryRows {
    param([string]$Reference)
    $plain = $Reference.Replace("'", '').Trim()
    $pattern = '^' + [regex]::Escape($script:CategorySheet) + '!\$?' +
               [regex]::Escape($script:CategoryColumn) + '\$?(\d+)(?::\$?' +
               [regex]::Escape($script:CategoryColumn) + '\$?(\d+))?$'
    $match = [regex]::Match($plain, $pattern)
    if (-not $match.Success) { return @(0, 0) }
    $first = [int]$match.Groups[1].Value
    $last = $first
    if ($match.Groups[2].Success) { $last = [int]$match.Groups[2].Value }
    return @($first, $last)
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
        # THE APPLIED-YEAR NAMES, as Excel resolves them right now. They are
        # WORKBOOK-scoped, so they are read from the workbook's own collection:
        # looking only inside Results.Names would report every one of them
        # missing while they were all present.
        $names = $workbook.Names
        $nameCount = [int]$names.Count
        $seenWindow = $false
        $seenCount = $false
        for ($n = 1; $n -le $nameCount; $n++) {
            $name = $null
            try {
                $name = $names.Item($n)
                $shortName = [string]$name.Name
                if ($shortName -like '*chartAnnual_*') {
                    $rowsNow = '<no range>'
                    try { $rowsNow = [string]$name.RefersToRange.Rows.Count } catch { $rowsNow = '<not a range now>' }
                    Emit-Line ('NAME|' + $shortName + '|refers_to=' + [string]$name.RefersTo + '|rows_now=' + $rowsNow)
                    if ($shortName -like ('*' + $script:CategoryWindowName)) {
                        $seenWindow = $true
                        # THE RESERVED CATEGORY WINDOW, WHICH IS THE CONTRACT the
                        # bound range is judged against: its sheet, its column,
                        # its first row and its last.
                        $windowRange = $null
                        try {
                            $windowRange = $name.RefersToRange
                            $script:CategorySheet = [string]$windowRange.Worksheet.Name
                            $script:CategoryFirstRow = [int]$windowRange.Row
                            $script:CategoryLastRow = [int]$windowRange.Row + [int]$windowRange.Rows.Count - 1
                            $script:CategoryColumn = [string]$windowRange.Cells(1, 1).Address($true, $true).Split([char]36)[1]
                        } catch {
                            Add-Failure 'Results' ('the category window name does not resolve to a range: ' + $_.Exception.Message)
                        } finally { Release-Object $windowRange }
                    }
                    if ($shortName -like ('*' + $script:YearCountName)) {
                        $seenCount = $true
                        # THE PUBLISHED YEAR COUNT, from the one authority, read
                        # the same way the presentation owner reads it.
                        $countRange = $null
                        try {
                            $countRange = $name.RefersToRange
                            $reported = $countRange.Cells(1, 1).Value2
                            if ($null -ne $reported) {
                                $parsed = 0
                                if ([double]::TryParse([string]$reported, [ref]$parsed)) {
                                    $script:PublishedYears = [int]$parsed
                                }
                            }
                            Emit-Line ('YEARS|published=' + [string]$script:PublishedYears + '|reported=' +
                                       $(if ($null -eq $reported) { '<blank>' } else { [string]$reported }))
                        } catch {
                            Add-Failure 'Results' ('the year count name does not resolve to a range: ' + $_.Exception.Message)
                        } finally { Release-Object $countRange }
                    }
                }
            } finally { Release-Object $name }
        }
        if (-not $seenWindow) {
            Add-Failure 'Results' ('the category window name ' + $script:CategoryWindowName + ' is not defined in the workbook')
        }
        if (-not $seenCount) {
            Add-Failure 'Results' ('the year count name ' + $script:YearCountName + ' is not defined in the workbook')
        }
        $objects = $null
        try {
            $objects = $dashboard.ChartObjects()
            $count = [int]$objects.Count
            Emit-Line ('CHARTS|count=' + [string]$count)
            $seenYearCharts = @()
            $seenLiteralCharts = @()
            for ($i = 1; $i -le $count; $i++) {
                $object = $null; $chart = $null; $collection = $null
                try {
                    $object = $objects.Item($i)
                    $chart = $object.Chart
                    $title = '<untitled>'
                    try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = '<unreadable>' }
                    Emit-Line ('CHART|' + $title + '|type=' + [string]$chart.ChartType)
                    $isYearChart = ($script:YearCharts -contains $title)
                    $isLiteralChart = ($script:LiteralCharts -contains $title)
                    if ($isYearChart) { $seenYearCharts = $seenYearCharts + $title }
                    if ($isLiteralChart) { $seenLiteralCharts = $seenLiteralCharts + $title }
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
                            # THE GATE. A BLANK CATEGORY ARGUMENT FAILS EVERY CHART -
                            # it is what Excel leaves behind when it drops a binding,
                            # and the histogram and the tornado came back blank too.
                            if ([string]::IsNullOrWhiteSpace($categories)) {
                                Add-Failure $title ('series ' + [string]$s +
                                    ' has BLANK XValues/categories; Excel dropped the binding and is numbering the categories 1, 2, 3')
                            } elseif ($isYearChart) {
                                # A YEAR CHART'S CATEGORIES ARE BOUND AT RUNTIME to the
                                # reserved calendar-year column, resized to the published
                                # year count. This is the live Excel object's own answer,
                                # after the presentation owner has run: the first row must
                                # be the block's first row, and the last must be the one
                                # the published count implies - one row when nothing is
                                # published, and never the whole reserved window.
                                $rows = Get-CategoryRows -Reference $categories
                                $wantFirst = $script:CategoryFirstRow
                                $wantLast = $wantFirst
                                if ($script:PublishedYears -ge 1) {
                                    $wantLast = $wantFirst + $script:PublishedYears - 1
                                    if ($wantLast -gt $script:CategoryLastRow) { $wantLast = $script:CategoryLastRow }
                                }
                                if (@($rows)[0] -eq 0) {
                                    Add-Failure $title ('series ' + [string]$s +
                                        ' plots categories from ' + $categories +
                                        '; a year chart is bound at runtime to ' + $script:CategorySheet +
                                        '!' + $script:CategoryColumn + ' and this is not that column')
                                } else {
                                    if (@($rows)[0] -ne $wantFirst) {
                                        Add-Failure $title ('series ' + [string]$s +
                                            ' starts its categories at row ' + [string]@($rows)[0] +
                                            ', not the first annual row ' + [string]$wantFirst)
                                    }
                                    if (@($rows)[1] -ne $wantLast) {
                                        Add-Failure $title ('series ' + [string]$s +
                                            ' ends its categories at row ' + [string]@($rows)[1] +
                                            ', not row ' + [string]$wantLast + ' implied by ' +
                                            [string]$script:PublishedYears + ' published year(s)')
                                    }
                                    if (@($rows)[1] -eq $script:CategoryLastRow -and $script:PublishedYears -lt ($script:CategoryLastRow - $script:CategoryFirstRow + 1)) {
                                        Add-Failure $title ('series ' + [string]$s +
                                            ' is bound to the whole reserved year window')
                                    }
                                }
                            } elseif ($isLiteralChart) {
                                # A FIXED-POPULATION CHART PLOTS ITS ACCEPTED LITERAL
                                # RANGE. A name here would mean the applied-year
                                # correction had reached a chart that never needed it.
                                if (-not (Test-CellRange -Reference $categories)) {
                                    Add-Failure $title ('series ' + [string]$s +
                                        ' plots categories from ' + $categories +
                                        '; this chart plots a fixed population and its categories are an ordinary cell range')
                                }
                            }
                            if ([string]::IsNullOrWhiteSpace($values)) {
                                Add-Failure $title ('series ' + [string]$s + ' has BLANK values')
                            }
                            # AND THE POINT COUNT IS REPORTED FOR EVERY SERIES, so a
                            # year chart drawing 200 points instead of the produced
                            # years is visible in the record even where it passes.
                            if ($isYearChart -and ($points -ne '<unreadable>') -and ([int]$points -gt 200)) {
                                Add-Failure $title ('series ' + [string]$s +
                                    ' draws ' + $points + ' points, more than the reserved year window holds')
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
            foreach ($wanted in $script:LiteralCharts) {
                if (-not ($seenLiteralCharts -contains $wanted)) {
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
