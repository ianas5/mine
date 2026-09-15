<#
.SYNOPSIS
    Executes the Windows chart-binding gate's PURE decision functions -
    ConvertTo-ComparableValue, ConvertTo-XValueList, Compare-Payload and
    Format-Payload - lifted out of bootstrap/windows/phase10_chart_polish_inspect.ps1
    by AST and driven over the COM shapes Series.XValues really returns, under
    Set-StrictMode 2.0. Excel is never started. IT ASSERTS NOTHING: it prints
    GATE|<case>|<text> lines and the Python control decides.

    WHY THIS EXISTS. The gate's oracle is Series.XValues, and Windows proved
    that is the live truth - a disposable copy returned XVALUES|2099 with the
    Dashboard protected again, while the SERIES formula's category argument can
    read blank for a series whose categories are plainly present. So the gate
    must normalise whatever COM hands back and compare it, in order, against
    the cells the contract expects. That normalisation and that comparison are
    what is exercised here.
#>
param([string]$Inspector)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$source = Get-Content -LiteralPath $Inspector -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('ConvertTo-ComparableValue', 'ConvertTo-XValueList', 'ConvertTo-CellList',
                    'Compare-Payload', 'Format-Payload')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the inspector defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}

function Emit { param([string]$Case, [string]$Text) Write-Output ('GATE|' + $Case + '|' + $Text) }

# THE CONTRACT THE YEAR CHARTS ARE JUDGED BY, in the gate's own terms: the first
# N cells of the reserved window, one when nothing is published, capped at the
# window. This mirrors what the inspector derives from the workbook.
$script:WindowRows = 200
function Expected-Year {
    param([int]$PublishedYears, $WindowCells)
    $take = 1
    if ($PublishedYears -ge 1) {
        $take = $PublishedYears
        if ($take -gt $script:WindowRows) { $take = $script:WindowRows }
    }
    return @(@($WindowCells)[0..($take - 1)])
}
# The reserved window as the bridge holds it with N years published: N calendar
# years, then blanks to the end of the window.
function Window-Cells {
    param([int]$PublishedYears)
    $cells = @()
    for ($i = 0; $i -lt $script:WindowRows; $i++) {
        if ($i -lt $PublishedYears) { $cells += (ConvertTo-ComparableValue (2026 + $i)) }
        else { $cells += '' }
    }
    return @($cells)
}
function Read-Year {
    param([string]$Case, $RawXValues, [int]$PublishedYears)
    $window = Window-Cells -PublishedYears $PublishedYears
    $expected = Expected-Year -PublishedYears $PublishedYears -WindowCells $window
    $actual = ConvertTo-XValueList -Raw $RawXValues
    $verdict = 'accepted'
    $problem = Compare-Payload -Actual $actual -Expected $expected
    if ($problem -ne '') { $verdict = 'refused' }
    elseif ((@($actual).Count -eq $script:WindowRows) -and ($PublishedYears -lt $script:WindowRows)) {
        $verdict = 'whole-window'
        $problem = 'presents the whole reserved year window'
    }
    Emit $Case ('count=' + [string]@($actual).Count + '|payload=' + (Format-Payload $actual) +
                '|verdict=' + $verdict + '|problem=' + $(if ($problem -eq '') { '<none>' } else { $problem }))
}
function Read-Literal {
    param([string]$Case, $RawXValues, $ExpectedCells)
    $actual = ConvertTo-XValueList -Raw $RawXValues
    $problem = Compare-Payload -Actual $actual -Expected @($ExpectedCells)
    Emit $Case ('count=' + [string]@($actual).Count + '|payload=' + (Format-Payload $actual) +
                '|verdict=' + $(if ($problem -eq '') { 'accepted' } else { 'refused' }) +
                '|problem=' + $(if ($problem -eq '') { '<none>' } else { $problem }))
}
# A 1-BASED VARIANT ARRAY, which is what Excel hands back for a multi-point
# series, and a 2-D rectangle, which is what a range-shaped answer looks like.
function New-ComArray {
    param([object[]]$Items)
    $out = [Array]::CreateInstance([object], @($Items.Count), @(1))
    for ($i = 0; $i -lt $Items.Count; $i++) { $out.SetValue($Items[$i], $i + 1) }
    return , $out
}
function New-ComRectangle {
    param([object[]]$Items)
    $out = [Array]::CreateInstance([object], @($Items.Count, 1), @(1, 1))
    for ($i = 0; $i -lt $Items.Count; $i++) { $out.SetValue($Items[$i], $i + 1, 1) }
    return , $out
}

# 1. NOTHING PUBLISHED: exactly one category, the first bridge cell's own value,
#    which is blank. A scalar is what Excel returns for a single point.
Read-Year 'year.nothing-published.scalar' '' 0
Read-Year 'year.nothing-published.array' (New-ComArray @('')) 0
# 2. ONE YEAR, then three, then two hundred.
Read-Year 'year.one' (New-ComArray @(2026)) 1
Read-Year 'year.three' (New-ComArray @(2026, 2027, 2028)) 3
Read-Year 'year.ten' (New-ComArray @(2026, 2027, 2028, 2029, 2030, 2031, 2032, 2033, 2034, 2035)) 10
Read-Year 'year.whole-window-published' (New-ComArray (Window-Cells -PublishedYears 200)) 200
# 3. THE WHOLE RESERVED WINDOW WHILE TEN YEARS ARE PUBLISHED.
Read-Year 'year.whole-window-refused' (New-ComArray (Window-Cells -PublishedYears 10)) 10
# 4. THE WRONG PAYLOAD, each way it can be wrong.
Read-Year 'year.wrong-first' (New-ComArray @(2025, 2027, 2028)) 3
Read-Year 'year.wrong-last' (New-ComArray @(2026, 2027, 2099)) 3
Read-Year 'year.wrong-order' (New-ComArray @(2028, 2027, 2026)) 3
Read-Year 'year.too-few' (New-ComArray @(2026, 2027)) 3
Read-Year 'year.too-many' (New-ComArray @(2026, 2027, 2028, 2029)) 3
Read-Year 'year.fabricated-when-empty' (New-ComArray @(2026)) 0
# 5. THE SHAPES COM CAN RETURN, all normalising to the same ordered payload.
Read-Year 'shape.rectangle' (New-ComRectangle @(2026, 2027, 2028)) 3
Read-Year 'shape.plain-array' (@(2026, 2027, 2028)) 3
Read-Year 'shape.numeric-text' (New-ComArray @('2026', '2027', '2028')) 3
# 6. THE LITERAL CHARTS, whose populations are fixed and whose SERIES formula
#    category argument Windows found blank while the categories were present.
$bins = @(); for ($i = 0; $i -lt 20; $i++) { $bins += (ConvertTo-ComparableValue (1000000 + $i * 50000)) }
Read-Literal 'literal.histogram.correct' (New-ComArray $bins) $bins
Read-Literal 'literal.histogram.short' (New-ComArray @($bins[0..17])) $bins
$drivers = @('Steel price', 'Labour rate', 'Ground conditions')
Read-Literal 'literal.tornado.correct' (New-ComArray $drivers) $drivers
Read-Literal 'literal.tornado.reordered' (New-ComArray @('Labour rate', 'Steel price', 'Ground conditions')) $drivers
Read-Literal 'literal.tornado.fabricated' (New-ComArray @('Steel price', 'Labour rate', 'Invented driver')) $drivers
