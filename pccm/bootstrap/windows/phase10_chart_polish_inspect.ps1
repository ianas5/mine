<#
.SYNOPSIS
    Reads the four Dashboard charts as Excel has them LOADED, and decides
    whether each one presents the categories its contract says it should.
    Excel is opened by this script, the workbook is opened read-only, nothing is
    written and nothing is saved.

    THE ORACLE IS Series.XValues, NOT THE SERIES FORMULA. An earlier version of
    this gate parsed the second argument of `=SERIES(...)` and treated a blank
    there as proof the categories were gone. That was never valid, and Windows
    proved it twice over: the runtime binding works - a disposable copy opened
    with events on returned XVALUES|2099 with the Dashboard protected again -
    while the formula argument can read blank for a series whose categories are
    plainly present, which is also why this gate once reported the histogram and
    the tornado broken when their accepted literal populations were intact. So
    the formula is REPORTED, as a diagnostic, and Series.XValues is what the
    verdict is taken on.

    WHAT EACH CHART IS HELD TO.

    THE TWO YEAR CHARTS present the applied years. Their expected categories are
    derived from the workbook itself - the reserved category window and the
    published year count, both named by the builder - and the live XValues must
    equal the first N cells of that window, in order. With nothing published
    that is exactly one item, the first bridge cell's own current value, which
    is normally blank. Never the whole reserved window while fewer years are
    published, and the two series of the cumulative chart must present identical
    categories.

    THE HISTOGRAM AND THE TORNADO present fixed populations. Their expected
    categories are the cells of the literal source ranges the accepted chart
    projection declares, read from that projection so this script declares no
    address of its own.

.PARAMETER WorkbookPath
    The .xlsm to inspect - a Stage-B build, or the copy the presentation
    scenario was run in. The category binding is applied when the workbook
    opens, so the workbook must be opened with events ENABLED, which is what
    this script does.

.PARAMETER ChartsProjection
    The accepted Phase-8 chart projection, phase8_charts_inspection.json, which
    is where the literal charts' declared category ranges come from. Defaults to
    the build directory of this script's own repository.

.PARAMETER ReportPath
    Optional. Where the lines are also written, UTF-8, LF.

.NOTES
    Reads only. No production endpoint is run, no cell is written, and the
    workbook is closed with SaveChanges:=False. Windows PowerShell 5.1.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)] [string]$WorkbookPath,
    [string]$ChartsProjection = '',
    [string]$ReportPath = ''
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

# THE TWO CONTRACTS, BY CHART. The titles are the manifest's; everything else -
# the category window, the year count, the literal ranges - is read from the
# workbook or from the accepted projection, so this script spells no address.
$script:YearCharts = @('Cumulative Cost Profile', 'Annual Cash Flow')
$script:LiteralCharts = @('Total Cost Distribution', 'Top Drivers by Rank Correlation')
$script:CategoryWindowName = 'chartAnnual_category_window'
$script:YearCountName = 'chartAnnual_year_count'

$script:Lines = New-Object System.Collections.ArrayList
$script:Failures = New-Object System.Collections.ArrayList
$script:ExpectedYear = @()
$script:ExpectedLiteral = @{}
$script:PublishedYears = 0
$script:WindowRows = 0

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

# ONE CELL OR ONE XVALUE, AS TEXT, so two of them can be compared without
# caring whether Excel handed back a Double, a String or Empty. A number is
# compared as a number - 2027 and 2027.0 are the same year - and everything
# else as its trimmed text.
function ConvertTo-ComparableValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.DBNull]) { return '' }
    $text = ''
    try { $text = [string]$Value } catch { return '<unreadable>' }
    $text = $text.Trim()
    if ($text -eq '') { return '' }
    $number = 0.0
    if ([double]::TryParse($text, [ref]$number)) {
        return $number.ToString('R', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    return $text
}

# THE COM SHAPE, NORMALISED. Series.XValues answers a 1-based Variant array for
# a multi-point series, a bare scalar for a single point, and - for a series
# whose source cannot be read at all - it raises. A 2-D rectangle is flattened
# in row order, which is the order Excel plots it in.
function ConvertTo-XValueList {
    param($Raw)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Raw) { return @() }
    if ($Raw -is [System.Array]) {
        $rank = 1
        try { $rank = [int]$Raw.Rank } catch { $rank = 1 }
        if ($rank -eq 2) {
            $rows = $Raw.GetLength(0); $columns = $Raw.GetLength(1)
            $firstRow = $Raw.GetLowerBound(0); $firstColumn = $Raw.GetLowerBound(1)
            for ($r = 0; $r -lt $rows; $r++) {
                for ($c = 0; $c -lt $columns; $c++) {
                    $null = $out.Add((ConvertTo-ComparableValue $Raw.GetValue($firstRow + $r, $firstColumn + $c)))
                }
            }
        } else {
            foreach ($item in $Raw) { $null = $out.Add((ConvertTo-ComparableValue $item)) }
        }
        return @($out)
    }
    $null = $out.Add((ConvertTo-ComparableValue $Raw))
    return @($out)
}

# THE CELLS OF ONE RANGE, in row order, normalised the same way.
function ConvertTo-CellList {
    param($Target)
    $out = New-Object System.Collections.ArrayList
    $rows = [int]$Target.Rows.Count
    $columns = [int]$Target.Columns.Count
    for ($r = 1; $r -le $rows; $r++) {
        for ($c = 1; $c -le $columns; $c++) {
            $cell = $null
            try {
                $cell = $Target.Cells($r, $c)
                $null = $out.Add((ConvertTo-ComparableValue $cell.Value2))
            } finally { Release-Object $cell }
        }
    }
    return @($out)
}

# A BOUNDED RENDERING OF A PAYLOAD, so a 200-item list does not fill the report.
function Format-Payload {
    param($Items, [int]$Limit = 8)
    $all = @($Items)
    if ($all.Count -eq 0) { return '<empty>' }
    $shown = $all
    $suffix = ''
    if ($all.Count -gt $Limit) {
        $shown = $all[0..($Limit - 1)]
        $suffix = ',... ' + [string]($all.Count - $Limit) + ' more'
    }
    $rendered = @()
    foreach ($item in $shown) { $rendered += $(if ([string]$item -eq '') { '<blank>' } else { [string]$item }) }
    return (($rendered -join ',') + $suffix)
}

# TWO PAYLOADS, COMPARED ITEM FOR ITEM AND IN ORDER. Returns '' when they
# agree, or the first disagreement, which is what a reader needs.
function Compare-Payload {
    param($Actual, $Expected)
    $left = @($Actual); $right = @($Expected)
    if ($left.Count -ne $right.Count) {
        return ('presents ' + [string]$left.Count + ' categories where ' + [string]$right.Count + ' are expected')
    }
    for ($i = 0; $i -lt $left.Count; $i++) {
        if ([string]$left[$i] -cne [string]$right[$i]) {
            return ('category ' + [string]($i + 1) + ' is ' +
                    $(if ([string]$left[$i] -eq '') { '<blank>' } else { [string]$left[$i] }) + ', expected ' +
                    $(if ([string]$right[$i] -eq '') { '<blank>' } else { [string]$right[$i] }))
        }
    }
    return ''
}

# The series formula, for the record only. It is not the oracle.
function Get-SeriesFormula {
    param($Series)
    try { return [string]$Series.Formula } catch { return '<unreadable>' }
}

# =============================================================================
# THE DECLARED LITERAL CATEGORY RANGES, from the accepted chart projection.
# =============================================================================
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($ChartsProjection)) {
    $ChartsProjection = Join-Path (Split-Path -Parent (Split-Path -Parent $scriptDir)) `
        (Join-Path 'build' 'phase8_charts_inspection.json')
}
$projection = $null
if (Test-Path -LiteralPath $ChartsProjection) {
    $projection = Get-Content -LiteralPath $ChartsProjection -Raw | ConvertFrom-Json
    Emit-Line ('PROJECTION|' + $ChartsProjection)
} else {
    Emit-Line ('PROJECTION|<not found> ' + $ChartsProjection)
}

$resolved = (Resolve-Path -LiteralPath $WorkbookPath).Path
$excel = $null; $workbooks = $null; $workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # EVENTS ON: the category binding is applied when the workbook opens, and a
    # gate that suppressed that would be inspecting a workbook nobody will have.
    $excel.EnableEvents = $true
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Open($resolved, 0, $true)
    Emit-Line ('WORKBOOK|' + [string]$workbook.FullName)

    $sheets = $null; $dashboard = $null; $names = $null
    try {
        $sheets = $workbook.Worksheets
        $dashboard = $sheets.Item('Dashboard')
        # THE WORKBOOK'S OWN NAMES. They are WORKBOOK-scoped, so looking inside
        # one sheet's collection would report every one of them missing while
        # they were all present.
        $names = $workbook.Names
        $nameCount = [int]$names.Count
        $windowRange = $null
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
                        try { $windowRange = $name.RefersToRange }
                        catch { Add-Failure 'Results' ('the category window name does not resolve to a range: ' + $_.Exception.Message) }
                    }
                    if ($shortName -like ('*' + $script:YearCountName)) {
                        $countRange = $null
                        try {
                            $countRange = $name.RefersToRange
                            $reported = $countRange.Cells(1, 1).Value2
                            $parsed = 0.0
                            if ($null -ne $reported) {
                                if ([double]::TryParse([string]$reported, [ref]$parsed)) {
                                    if ($parsed -ge 1) { $script:PublishedYears = [int][Math]::Floor($parsed) }
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
        if ($null -eq $windowRange) {
            Add-Failure 'Results' ('the category window name ' + $script:CategoryWindowName + ' is not defined in the workbook')
        } else {
            # THE EXPECTED YEAR CATEGORIES: the first N cells of the reserved
            # window, capped at the window, and exactly one when nothing is
            # published - the first cell's own current value, whatever it is.
            try {
                $script:WindowRows = [int]$windowRange.Rows.Count
                $take = 1
                if ($script:PublishedYears -ge 1) {
                    $take = $script:PublishedYears
                    if ($take -gt $script:WindowRows) { $take = $script:WindowRows }
                }
                $slice = $null
                try {
                    $slice = $windowRange.Resize($take, 1)
                    $script:ExpectedYear = ConvertTo-CellList -Target $slice
                } finally { Release-Object $slice }
                Emit-Line ('EXPECTED|year|count=' + [string]@($script:ExpectedYear).Count +
                           '|payload=' + (Format-Payload $script:ExpectedYear))
            } catch {
                Add-Failure 'Results' ('the expected year categories could not be read: ' + $_.Exception.Message)
            } finally { Release-Object $windowRange }
        }
        # THE EXPECTED LITERAL CATEGORIES, from the accepted projection's own
        # declared ranges - no address is spelled here.
        if ($null -ne $projection) {
            foreach ($chart in @($projection.charts)) {
                $title = [string]$chart.title
                if (-not ($script:LiteralCharts -contains $title)) { continue }
                $reference = [string]$chart.categories.range
                $parts = $reference.Split([char]33)
                if ($parts.Count -ne 2) {
                    Add-Failure $title ('the projected category range ' + $reference + ' names no sheet')
                    continue
                }
                $sourceSheet = $null; $sourceRange = $null
                try {
                    $sourceSheet = $sheets.Item($parts[0])
                    $sourceRange = $sourceSheet.Range($parts[1])
                    $script:ExpectedLiteral[$title] = ConvertTo-CellList -Target $sourceRange
                    Emit-Line ('EXPECTED|' + $title + '|count=' + [string]@($script:ExpectedLiteral[$title]).Count +
                               '|payload=' + (Format-Payload $script:ExpectedLiteral[$title]))
                } catch {
                    Add-Failure $title ('the projected category range ' + $reference + ' could not be read: ' + $_.Exception.Message)
                } finally {
                    Release-Object $sourceRange
                    Release-Object $sourceSheet
                }
            }
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
                    if (($isYearChart -or $isLiteralChart) -and ($seriesCount -lt 1)) {
                        Add-Failure $title 'plots no series at all'
                    }
                    $perSeries = @()
                    for ($s = 1; $s -le $seriesCount; $s++) {
                        $series = $null
                        try {
                            $series = $collection.Item($s)
                            $label = 'CHART|' + $title + '|series' + [string]$s
                            Emit-Line ($label + '.formula=' + (Get-SeriesFormula -Series $series))
                            # THE ORACLE. Series.XValues, read directly and
                            # normalised; unreadable is a failure, not a blank.
                            $readable = $true
                            $xvalues = @()
                            try { $xvalues = ConvertTo-XValueList -Raw $series.XValues }
                            catch { $readable = $false }
                            $values = @()
                            try { $values = @($series.Values) } catch { $values = @() }
                            if (-not $readable) {
                                Emit-Line ($label + '.xvalues=<unreadable>')
                                if ($isYearChart -or $isLiteralChart) {
                                    Add-Failure $title ('series ' + [string]$s + ' XValues could not be read')
                                }
                            } else {
                                Emit-Line ($label + '.xvalues.count=' + [string]@($xvalues).Count)
                                Emit-Line ($label + '.xvalues=' + (Format-Payload $xvalues))
                                $perSeries = $perSeries + ,@($xvalues)
                            }
                            Emit-Line ($label + '.points=' + [string]@($values).Count)
                            if ($readable) {
                                if ($isYearChart) {
                                    if (@($script:ExpectedYear).Count -eq 0) {
                                        Add-Failure $title ('series ' + [string]$s +
                                            ' cannot be judged: the expected year categories were not established')
                                    } else {
                                        $problem = Compare-Payload -Actual $xvalues -Expected $script:ExpectedYear
                                        if ($problem -ne '') {
                                            Add-Failure $title ('series ' + [string]$s + ' ' + $problem)
                                        }
                                        # AND NEVER THE WHOLE RESERVED WINDOW while
                                        # fewer years than that are published.
                                        if ((@($xvalues).Count -eq $script:WindowRows) -and
                                            ($script:PublishedYears -lt $script:WindowRows)) {
                                            Add-Failure $title ('series ' + [string]$s +
                                                ' presents the whole reserved year window, ' +
                                                [string]$script:WindowRows + ' categories, with ' +
                                                [string]$script:PublishedYears + ' year(s) published')
                                        }
                                    }
                                } elseif ($isLiteralChart) {
                                    if (-not $script:ExpectedLiteral.ContainsKey($title)) {
                                        Add-Failure $title ('series ' + [string]$s +
                                            ' cannot be judged: no declared category range was read for this chart')
                                    } else {
                                        $problem = Compare-Payload -Actual $xvalues -Expected $script:ExpectedLiteral[$title]
                                        if ($problem -ne '') {
                                            Add-Failure $title ('series ' + [string]$s + ' ' + $problem)
                                        }
                                    }
                                }
                            }
                        } finally { Release-Object $series }
                    }
                    # THE CUMULATIVE CHART'S TWO SERIES SHARE ONE AXIS: they must
                    # present the same categories, or one of them is drawn against
                    # years it was not produced for.
                    if ($isYearChart -and (@($perSeries).Count -gt 1)) {
                        for ($s = 1; $s -lt @($perSeries).Count; $s++) {
                            $problem = Compare-Payload -Actual @($perSeries)[$s] -Expected @($perSeries)[0]
                            if ($problem -ne '') {
                                Add-Failure $title ('series ' + [string]($s + 1) +
                                    ' presents different categories from series 1: ' + $problem)
                            }
                        }
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
