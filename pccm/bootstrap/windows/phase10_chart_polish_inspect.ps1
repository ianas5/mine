<#
.SYNOPSIS
    Reads the four Dashboard chart properties the final-delivery chart polish
    declares, from a workbook that already holds a published result, and prints
    them as CHART|<title>|<property>=<value> lines. Excel is opened by this
    script, the workbook is opened read-only, nothing is written and nothing is
    saved. A presentation review is done by a person looking at the Dashboard;
    this exports the properties so that review can be recorded exactly.

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

$script:Lines = New-Object System.Collections.ArrayList
function Emit-Line {
    param([string]$Text)
    $null = $script:Lines.Add($Text)
    Write-Host $Text
}
function Release-Object {
    param($Obj)
    if ($null -ne $Obj) {
        try { $null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch { }
    }
}
function Get-SeriesFormula {
    param($Series)
    try { return [string]$Series.Formula } catch { return '<unreadable>' }
}
function Get-PointCount {
    # HOW MANY POINTS THE SERIES ACTUALLY CARRIES - the applied years, once a
    # result is published - not how many rows the reserved window holds.
    param($Series)
    try { return [string]@($Series.Values).Count } catch { return '<unreadable>' }
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
        for ($n = 1; $n -le $nameCount; $n++) {
            $name = $null
            try {
                $name = $names.Item($n)
                $shortName = [string]$name.Name
                if ($shortName -like '*chartAnnual_*') {
                    $rowsNow = '<no range>'
                    try { $rowsNow = [string]$name.RefersToRange.Rows.Count } catch { $rowsNow = '<not a range now>' }
                    Emit-Line ('NAME|' + $shortName + '|refers_to=' + [string]$name.RefersTo + '|rows_now=' + $rowsNow)
                }
            } finally { Release-Object $name }
        }
        $objects = $null
        try {
            $objects = $dashboard.ChartObjects()
            $count = [int]$objects.Count
            Emit-Line ('CHARTS|count=' + [string]$count)
            for ($i = 1; $i -le $count; $i++) {
                $object = $null; $chart = $null; $collection = $null
                try {
                    $object = $objects.Item($i)
                    $chart = $object.Chart
                    $title = '<untitled>'
                    try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = '<unreadable>' }
                    Emit-Line ('CHART|' + $title + '|type=' + [string]$chart.ChartType)
                    $collection = $chart.SeriesCollection()
                    $seriesCount = [int]$collection.Count
                    for ($s = 1; $s -le $seriesCount; $s++) {
                        $series = $null
                        try {
                            $series = $collection.Item($s)
                            Emit-Line ('CHART|' + $title + '|series' + [string]$s + '.formula=' + (Get-SeriesFormula -Series $series))
                            Emit-Line ('CHART|' + $title + '|series' + [string]$s + '.points=' + (Get-PointCount -Series $series))
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
if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($ReportPath, (($script:Lines -join "`n") + "`n"), $utf8)
    Write-Host ('Report written to ' + $ReportPath)
}
