<#
.SYNOPSIS
    PCCM Phase 8 - P8-Z: one question, in real Excel. Is a zero-variance driver
    absent from the tornado?

.DESCRIPTION
    THIS IS NOT THE P8-3 SUITE AND MUST NOT BE READ AS ONE. The 189-check P8-3
    acceptance at 9e3c141 stands; nothing here re-tests any of it. Four charts,
    seven parts, every state transition, every preservation comparison - all of
    that is settled and is not repeated. This file asks ONE question that the
    accepted W4 fixture could not: it has five drivers and all five vary, so it
    never produced a diagnostic row for the tornado window to pick up.

    THE QUESTION. sim_contract says a driver with no variance to correlate is

        zero_variance:
          status_label: "n/a - no variance"
          rho_reported: false
          reported_as_zero_rho: false
          excluded_from_ranking: true
          excluded_from_tornado_input: true
          retained_diagnostically: true

    Two of those are what this proves in Excel: RETAINED DIAGNOSTICALLY, so the
    Sensitivity sheet still shows the driver and says why it has no rho; and
    EXCLUDED FROM TORNADO INPUT, which the source settlement read as absent
    from the chart entirely - not present with an empty bar.

    HOW THE DRIVER IS MADE. Not by a new fixture. The accepted W4 model is
    applied exactly as P8-3 applies it, and then ONE cost line is given an
    identical minimum, most likely and maximum through the register. A driver
    whose unit cost cannot vary contributes the same amount in every iteration,
    so its rank vector is constant, so `sxx = 0` in the accepted kernel and the
    record comes back SIM_SENSITIVITY_NO_VARIANCE. That is the production path
    reaching the state on its own; nothing here fabricates a status.

    WHAT IS READ, AND WHAT IS NOT. The Sensitivity row is read as cells. The
    tornado is read out of Excel's own ChartObjects and SeriesCollection - the
    category range and the value range as Excel resolved them - and compared
    against the bridge cells frozen BEFORE the chart was inspected. No ranking
    is recomputed, no order is derived, and no other chart is touched.

    ONE SESSION, ONE DISPOSABLE WORKBOOK, and the same COM lifecycle discipline
    the accepted runners use.

.NOTES
    THIS FILE HAS NOT RUN. There is no Windows and no Excel where it was
    written, and nothing in it claims otherwise.
#>
[CmdletBinding()]
param(
    [string]$BuildDir
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# ALL THREE ARE DEFINITION-ONLY AT TOP LEVEL. Dot-sourcing them defines
# functions and script variables and runs no scenario.
. (Join-Path $scriptDir 'com_lifecycle.ps1')
. (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')
. (Join-Path $scriptDir 'phase6_gate_b_scenarios.ps1')

# EVERY PATH THIS RUNNER NEEDS, DERIVED FROM WHERE THE SCRIPT IS - the accepted
# pattern, and now the same three lines the accepted runners use rather than a
# variation on them.
#
#   bootstrap/windows  ->  pccm  ->  the repository root
#
# NOT THE WORKING DIRECTORY. The documented invocation is from the repository
# root, but nothing here may depend on that: a runner that only worked when
# launched from one place would be a runner whose evidence depended on how it
# was started.
$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$repoRoot = Split-Path -Parent $pccmRoot
if ([string]::IsNullOrWhiteSpace($BuildDir)) { $BuildDir = Join-Path $pccmRoot 'build' }

$script:P8ZChecks = New-Object System.Collections.ArrayList
$script:P8ZLines = New-Object System.Collections.ArrayList

$script:P8ZErrorCodes = @{
    -2146826288 = '#NULL!'; -2146826281 = '#DIV/0!'; -2146826273 = '#VALUE!'
    -2146826265 = '#REF!';  -2146826259 = '#NAME?';  -2146826252 = '#NUM!'
    -2146826246 = '#N/A'
}

function Write-RowObject {
    param([object[]]$Row)
    Write-Output -NoEnumerate $Row
}

function Get-NamedValue {
    param($Workbook, [string]$DefinedName)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        $v = $rng.Value2
        if ($null -eq $v) { return '' }
        return [string]$v
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Set-NamedValue {
    param($Workbook, [string]$DefinedName, $Value)
    $names = $null; $nm = $null; $rng = $null
    try {
        $names = $Workbook.Names
        $nm = $names.Item($DefinedName)
        $rng = $nm.RefersToRange
        if ($null -eq $Value) { $null = $rng.ClearContents() } else { $rng.Value2 = [double]$Value }
    } finally {
        if ($null -ne $rng)   { Release-Transient $rng   'Range(name)'; $rng   = $null }
        if ($null -ne $nm)    { Release-Transient $nm    'Name';        $nm    = $null }
        if ($null -ne $names) { Release-Transient $names 'Names';       $names = $null }
    }
}

function Get-TableColumnNames {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $cols = $null
    $out = @()
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $cols = $lo.ListColumns
        $colCount = [int]$cols.Count
        for ($i = 1; $i -le $colCount; $i++) {
            $c = $null
            try { $c = $cols.Item($i); $out += [string]$c.Name }
            finally { if ($null -ne $c) { Release-Transient $c 'ListColumn'; $c = $null } }
        }
    } finally {
        if ($null -ne $cols)            { Release-Transient $cols            'ListColumns'; $cols            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
    return $out
}

function Set-TableCell {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex, [int]$ColumnIndex, $Value)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null; $cell = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        $cell = $body.Cells($RowIndex, $ColumnIndex)
        if ($null -eq $Value) {
            # A genuine blank, not zero. The two are different assumptions and the
            # harness has to be able to create each of them deliberately.
            $null = $cell.ClearContents()
        } elseif ($Value -is [string]) {
            $cell.Value2 = [string]$Value
        } else {
            $cell.Value2 = [double]$Value
        }
    } finally {
        if ($null -ne $cell)            { Release-Transient $cell            'Range(cell)'; $cell            = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)'; $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Get-TableBody {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $body = $null
    $rowsObj = $null; $colsObj = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $body = $lo.DataBodyRange
        # An empty body is a valid outcome: emit NOTHING. The caller's @(...) turns
        # zero pipeline objects into an empty collection, which is exactly right.
        if ($null -eq $body) { return }

        # Row and column counts are read through named, released objects rather than
        # through $body.Rows.Count, which would mint an unowned Range on every
        # iteration of the loop.
        $rowsObj = $body.Rows
        $colsObj = $body.Columns
        $rowCount = [int]$rowsObj.Count
        $colCount = [int]$colsObj.Count
        Release-Transient $rowsObj 'Range(rows)'; $rowsObj = $null
        Release-Transient $colsObj 'Range(columns)'; $colsObj = $null

        for ($r = 1; $r -le $rowCount; $r++) {
            $line = @()
            for ($c = 1; $c -le $colCount; $c++) {
                $cell = $null
                try {
                    $cell = $body.Cells($r, $c)
                    $v = $cell.Value2
                    if ($null -eq $v) { $line += '' } else { $line += [string]$v }
                } finally {
                    if ($null -ne $cell) { Release-Transient $cell 'Range(cell)'; $cell = $null }
                }
            }
            Write-RowObject $line
        }
    } finally {
        if ($null -ne $rowsObj)         { Release-Transient $rowsObj         'Range(rows)';    $rowsObj         = $null }
        if ($null -ne $colsObj)         { Release-Transient $colsObj         'Range(columns)'; $colsObj         = $null }
        if ($null -ne $body)            { Release-Transient $body            'Range(body)';    $body            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';     $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects';    $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';      $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';     $localWorksheets = $null }
    }
    # No trailing return: every row has already been emitted, one object each.
}

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return @(Get-TableBody -Workbook $Workbook -SheetName $SheetName -TableName $TableName).Count
}

function Add-BlankTableRow {
    param($Workbook, [string]$SheetName, [string]$TableName)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $added = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $added = $rows.Add()
        return [int]$added.Index
    } finally {
        if ($null -ne $added)           { Release-Transient $added           'ListRow';     $added           = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Remove-TableRow {
    param($Workbook, [string]$SheetName, [string]$TableName, [int]$RowIndex)
    $localWorksheets = $null; $ws = $null; $los = $null; $lo = $null; $rows = $null; $victim = $null
    try {
        $localWorksheets = $Workbook.Worksheets
        $ws = $localWorksheets.Item($SheetName)
        $los = $ws.ListObjects
        $lo = $los.Item($TableName)
        $rows = $lo.ListRows
        $victim = $rows.Item($RowIndex)
        $victim.Delete()
    } finally {
        if ($null -ne $victim)          { Release-Transient $victim          'ListRow';     $victim          = $null }
        if ($null -ne $rows)            { Release-Transient $rows            'ListRows';    $rows            = $null }
        if ($null -ne $lo)              { Release-Transient $lo              'ListObject';  $lo              = $null }
        if ($null -ne $los)             { Release-Transient $los             'ListObjects'; $los             = $null }
        if ($null -ne $ws)              { Release-Transient $ws              'Worksheet';   $ws              = $null }
        if ($null -ne $localWorksheets) { Release-Transient $localWorksheets 'Worksheets';  $localWorksheets = $null }
    }
}

function Get-IdColumnValues {
    param($Workbook, $Info)
    $out = @()
    foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $Info.sheet -TableName $Info.table_name)) {
        if ($row[0] -ne '') { $out += $row[0] }
    }
    return $out
}

function Write-P8ZLine {
    param([string]$Text = '')
    $null = $script:P8ZLines.Add($Text)
    Write-Host $Text
    # Written through on every line, so a run that is stopped still leaves every
    # observation it had already taken on disk.
    if (-not [string]::IsNullOrWhiteSpace($script:P8ZPath)) {
        try {
            Set-Content -LiteralPath $script:P8ZPath `
                -Value ($script:P8ZLines -join "`r`n") -Encoding UTF8
        } catch { }
    }
}

function Add-P8ZCheck {
    param([string]$Label, [bool]$Ok, [string]$Detail = '', [string]$Kind = 'RESULT')
    $null = $script:P8ZChecks.Add([pscustomobject]@{
        Kind = $Kind; Label = $Label; Ok = $Ok; Detail = $Detail
    })
    $verdict = 'FAIL'
    if ($Ok) { $verdict = 'PASS' }
    $line = '  [' + $verdict + '] ' + $Label
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $line = $line + ' -- ' + $Detail }
    Write-P8ZLine $line
    return $Ok
}

function Invoke-P8ZRelease {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(
            ("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
        if ([int]$rec.Count -ne 0) {
            $null = $script:P8ZResidual.Add(
                ([string]$rec.Label + ' left ' + [string]$rec.Count + ' outstanding'))
        }
    } elseif ($rec.Status -eq 'FAIL') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $null = $Ledger.Failed.Add($rec.Label)
        $null = $Ledger.Lines.Add(("      {0,-24} | FAIL    | {1}" -f $rec.Label, $rec.Error))
    } else {
        $null = $Ledger.Lines.Add(("      {0,-24} | SKIPPED | {1}" -f $rec.Label, $rec.Error))
    }
}

function Get-P8ZSourceRevision {
    param([string]$RepoRoot)
    $head = ''
    try { $head = [string](& git -C $RepoRoot rev-parse HEAD 2>$null) } catch { $head = '' }
    $head = $head.Trim()
    if ([string]::IsNullOrWhiteSpace($head)) {
        throw ('git could not report HEAD for ' + $RepoRoot +
               '; a P8Z result could not be attributed to a source revision')
    }
    $dirty = @()
    foreach ($pathspec in @('pccm/src', 'pccm/spec', 'pccm/builder')) {
        $lines = @()
        try { $lines = @(& git -C $RepoRoot status --porcelain -- $pathspec 2>$null) } catch { $lines = @() }
        foreach ($line in $lines) {
            if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $dirty += [string]$line }
        }
    }
    return [pscustomobject]@{ Head = $head; Dirty = $dirty }
}

function Invoke-P8ZEndpoint {
    param($Excel, [string]$Endpoint)
    $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null
    $Excel.Run($Endpoint) | Out-Null
    return [string]$Excel.Run('PCCM_AutomationResult')
}

function Get-P8ZRegister {
    param($Manifest, [string]$Key)
    foreach ($register in @($Manifest.registers)) {
        if (([string]$register.key) -ceq $Key) { return $register }
    }
    throw ('the Stage-B manifest projects no ' + [char]39 + $Key + [char]39 + ' register')
}

function Get-P8ZRegisterColumnIndex {
    param($Register, [string]$ColumnKey)
    $ordinal = [array]::IndexOf(@($Register.columns), $ColumnKey) + 1
    if ($ordinal -lt 1) {
        throw ('the register carries no ' + [char]39 + $ColumnKey + [char]39 + ' column')
    }
    return $ordinal
}

function Get-P8ZRegisterRowIndex {
    param($Workbook, $Register, [string]$PermanentId)
    $body = @(Get-TableBody -Workbook $Workbook -SheetName ([string]$Register.sheet) `
        -TableName ([string]$Register.table_name))
    for ($index = 0; $index -lt $body.Count; $index++) {
        $cells = @($body[$index])
        if ($cells.Count -lt 1) { continue }
        if (([string]$cells[0]) -ceq $PermanentId) { return ($index + 1) }
    }
    return 0
}

function Get-P8ZCell {
    param($Workbook, [string]$SheetName, [string]$Address)
    $sheets = $null; $sheet = $null; $range = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($SheetName)
        $range = $sheet.Range($Address)
        $value = $range.Value2
        $text = ''
        try { $text = [string]$range.Text } catch { $text = '<unreadable>' }
        $formula = ''
        try { $formula = [string]$range.Formula } catch { $formula = '<unreadable>' }
        $format = ''
        try { $format = [string]$range.NumberFormat } catch { $format = '<unreadable>' }
        $errorName = ''
        if (($value -is [int]) -or ($value -is [long])) {
            $code = [int]$value
            if ($script:P8ZErrorCodes.ContainsKey($code)) {
                $errorName = [string]$script:P8ZErrorCodes[$code]
            }
        }
        # A CELL WHOSE TEXT IS AN ERROR NAME IS AN ERROR whatever its Value2
        # arrived as. Both routes are kept because neither is guaranteed alone.
        if ([string]::IsNullOrEmpty($errorName)) {
            foreach ($name in $script:P8ZErrorCodes.Values) {
                if (([string]$text) -ceq [string]$name) { $errorName = [string]$name }
            }
        }
        return [pscustomobject]@{
            Sheet = $SheetName; Address = $Address; Value = $value; Text = $text
            Formula = $formula; NumberFormat = $format; ErrorName = $errorName
            IsError = (-not [string]::IsNullOrEmpty($errorName))
        }
    } finally {
        if ($null -ne $range)  { Release-Transient $range  ('Range(' + $SheetName + ')');      $range  = $null }
        if ($null -ne $sheet)  { Release-Transient $sheet  ('Worksheet(' + $SheetName + ')');  $sheet  = $null }
        if ($null -ne $sheets) { Release-Transient $sheets 'Worksheets';                       $sheets = $null }
    }
}

function Format-P8ZCell {
    param($Cell)
    if ($null -eq $Cell) { return '<not read>' }
    if ($Cell.IsError) { return ([string]$Cell.ErrorName + ' (error)') }
    return ([string]$Cell.Sheet + '!' + [string]$Cell.Address + ' = ' +
            (Format-SimValue $Cell.Value) + ' shown as ' + [char]39 + [string]$Cell.Text + [char]39)
}

function Test-P8ZBlank {
    param($Cell)
    if ($null -eq $Cell) { return $false }
    if ($Cell.IsError) { return $false }
    return (Test-SimBlank -Value $Cell.Value)
}

function Test-P8ZSameCellValue {
    param($Left, $Right)
    if (($null -eq $Left) -or ($null -eq $Right)) { return $false }
    # 1. ERRORS ARE COMPARED AS ERRORS, BY IDENTITY.
    if ($Left.IsError -or $Right.IsError) {
        if (-not ($Left.IsError -and $Right.IsError)) { return $false }
        return ([string]$Left.ErrorName -ceq [string]$Right.ErrorName)
    }
    # 2-4. EVERYTHING THAT IS NOT AN ERROR, decided in one place so a chart
    #      point and a worksheet cell are read by the same rule.
    return (Test-P8ZSameValue -A $Left.Value -B $Right.Value)
}

function Test-P8ZSameValue {
    param($A, $B)
    # BLANK IS ITS OWN THING, and is not zero.
    $leftBlank = Test-SimBlank -Value $A
    $rightBlank = Test-SimBlank -Value $B
    if ($leftBlank -or $rightBlank) { return ($leftBlank -and $rightBlank) }
    # TEXT AND NUMBERS DO NOT MEET, and neither is stringified to make them.
    $leftText = ($A -is [string])
    $rightText = ($B -is [string])
    if ($leftText -ne $rightText) { return $false }
    if ($leftText) { return ([string]$A -ceq [string]$B) }
    if (($A -is [bool]) -or ($B -is [bool])) {
        if (-not (($A -is [bool]) -and ($B -is [bool]))) { return $false }
        return ([bool]$A -eq [bool]$B)
    }
    # NUMERIC, ACROSS THE COM SUBTYPE. COM chooses Int32 or Double for the same
    # reading of the same cell; the workbook does not. No tolerance is applied,
    # so a real change of any size is still a change.
    return ([double]$A -eq [double]$B)
}

function Invoke-P8ZRecalculate {
    param($Excel, [string]$Stage)
    $mode = '<unreadable>'
    try { $mode = [string]$Excel.Calculation } catch { $mode = '<unreadable>' }
    $failure = ''
    try { $Excel.Calculate() } catch { $failure = (Format-Err $_) }
    Write-P8ZLine ('    recalculated at ' + $Stage + ' (Application.Calculation = ' + $mode + ')')
    return (Add-P8ZCheck ($Stage + ': the workbook recalculated') `
        ([string]::IsNullOrWhiteSpace($failure)) ($failure + ' calculation mode ' + $mode) `
        'PREREQUISITE')
}

function ConvertTo-P8ZColumnNumber {
    param([string]$Letters)
    $number = 0
    foreach ($character in $Letters.ToUpperInvariant().ToCharArray()) {
        $number = ($number * 26) + ([int]$character - 64)
    }
    return $number
}

function Get-P8ZCharts {
    param($Workbook, [string]$SheetName)
    $sheets = $null; $sheet = $null; $objects = $null
    $out = New-Object System.Collections.ArrayList
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($SheetName)
        $objects = $sheet.ChartObjects()
        $count = [int]$objects.Count
        for ($index = 1; $index -le $count; $index++) {
            $object = $null; $chart = $null; $series = $null; $anchor = $null
            try {
                $object = $objects.Item($index)
                $chart = $object.Chart
                $title = ''
                try { if ($chart.HasTitle) { $title = [string]$chart.ChartTitle.Text } } catch { $title = '<unreadable>' }
                $chartType = -1
                try { $chartType = [int]$chart.ChartType } catch { $chartType = -1 }
                $anchor = $object.TopLeftCell
                $anchorRow = [int]$anchor.Row
                $anchorColumn = [int]$anchor.Column
                $widthCm = ConvertTo-P8ZCentimetres -Points ([double]$object.Width)
                $heightCm = ConvertTo-P8ZCentimetres -Points ([double]$object.Height)
                $hasLegend = $false
                try { $hasLegend = [bool]$chart.HasLegend } catch { $hasLegend = $false }
                $series = $chart.SeriesCollection()
                $seriesCount = [int]$series.Count
                $collected = New-Object System.Collections.ArrayList
                for ($s = 1; $s -le $seriesCount; $s++) {
                    $item = $null
                    try {
                        $item = $series.Item($s)
                        $formula = [string]$item.Formula
                        $parts = Get-P8ZSeriesParts -Formula $formula
                        # THE PLOTTED VALUES THEMSELVES. A #N/A arrives as the
                        # CVErr code, which is how "no point" is told apart from
                        # "a point at zero" - the whole reason the bridge exists.
                        $values = @()
                        try { $values = @($item.Values) } catch { $values = @() }
                        $null = $collected.Add([pscustomobject]@{
                            Index = $s; Formula = $formula; Name = $parts.Name
                            Categories = (ConvertTo-P8ZNormalRange -Reference $parts.Categories)
                            ValuesRange = (ConvertTo-P8ZNormalRange -Reference $parts.Values)
                            Values = $values
                        })
                    } finally {
                        if ($null -ne $item) { Release-Transient $item 'Series'; $item = $null }
                    }
                }
                $null = $out.Add([pscustomobject]@{
                    Index = $index; Title = $title; ChartType = $chartType
                    AnchorRow = $anchorRow; AnchorColumn = $anchorColumn
                    WidthCm = $widthCm; HeightCm = $heightCm; HasLegend = $hasLegend
                    SeriesCount = $seriesCount; Series = @($collected)
                })
            } finally {
                if ($null -ne $anchor)  { Release-Transient $anchor  'TopLeftCell';       $anchor  = $null }
                if ($null -ne $series)  { Release-Transient $series  'SeriesCollection';  $series  = $null }
                if ($null -ne $chart)   { Release-Transient $chart   'Chart';             $chart   = $null }
                if ($null -ne $object)  { Release-Transient $object  'ChartObject';       $object  = $null }
            }
        }
    } finally {
        if ($null -ne $objects) { Release-Transient $objects 'ChartObjects'; $objects = $null }
        if ($null -ne $sheet)   { Release-Transient $sheet   'Worksheet';    $sheet   = $null }
        if ($null -ne $sheets)  { Release-Transient $sheets  'Worksheets';   $sheets  = $null }
    }
    return @($out)
}

function Test-P8ZNoPoint {
    param($Value)
    if ($null -eq $Value) { return $true }
    if (($Value -is [int]) -or ($Value -is [long])) {
        return ($script:P8ZErrorCodes.ContainsKey([int]$Value))
    }
    return $false
}

function Format-P8ZPoint {
    param($Value)
    if ($null -eq $Value) { return '<empty>' }
    if (Test-P8ZNoPoint -Value $Value) {
        $code = 0
        if (($Value -is [int]) -or ($Value -is [long])) { $code = [int]$Value }
        if ($script:P8ZErrorCodes.ContainsKey($code)) { return [string]$script:P8ZErrorCodes[$code] }
        return '<no point>'
    }
    return (Format-SimValue $Value)
}

function Get-P8ZSeriesParts {
    param([string]$Formula)
    # =SERIES(name, categories, values, order). Split at the TOP level only: a
    # range can carry no comma, but a quoted name can, and a naive split would
    # tear one in half.
    $inner = $Formula
    $open = $inner.IndexOf('(')
    if ($open -ge 0) { $inner = $inner.Substring($open + 1) }
    if ($inner.EndsWith(')')) { $inner = $inner.Substring(0, $inner.Length - 1) }
    $parts = New-Object System.Collections.ArrayList
    $depth = 0
    $quoted = $false
    $token = ''
    foreach ($character in $inner.ToCharArray()) {
        if ($character -eq '"') { $quoted = -not $quoted; $token = $token + [string]$character; continue }
        if (-not $quoted) {
            if ($character -eq '(') { $depth = $depth + 1 }
            elseif ($character -eq ')') { $depth = $depth - 1 }
            elseif (($character -eq ',') -and ($depth -eq 0)) {
                $null = $parts.Add($token); $token = ''; continue
            }
        }
        $token = $token + [string]$character
    }
    $null = $parts.Add($token)
    $all = @($parts)
    $name = ''; $categories = ''; $values = ''
    if ($all.Count -ge 1) { $name = ([string]$all[0]).Trim() }
    if ($all.Count -ge 2) { $categories = ([string]$all[1]).Trim() }
    if ($all.Count -ge 3) { $values = ([string]$all[2]).Trim() }
    return [pscustomobject]@{ Name = $name; Categories = $categories; Values = $values }
}

function ConvertTo-P8ZNormalRange {
    param([string]$Reference)
    $out = [string]$Reference
    $out = $out.Replace("'", '')
    $out = $out.Replace('=', '')
    $out = $out.Trim()
    return $out
}

function ConvertTo-P8ZCentimetres {
    param([double]$Points)
    return ($Points / 72.0 * 2.54)
}


# THE ONE READER THIS SCENARIO ADDS, and it is added rather than folded into the
# copied helper so that helper stays byte-identical to the accepted P8-3 one.
#
# `Get-P8ZCharts` captures each series' Values - the plotted numbers - and the
# category RANGE, but not the category VALUES. This question needs those: "the
# chart contains no category for that driver" is a statement about what Excel
# resolved the category range to, not about the range's address. Series.XValues
# is where Excel keeps them, and a #N/A arrives as the CVErr code exactly as it
# does in Values, so an absent category is told apart from an empty one.
function Get-P8ZCategoryValues {
    param($Workbook, [string]$SheetName, [string]$Title)
    $sheets = $null; $sheet = $null; $objects = $null
    $object = $null; $chart = $null; $series = $null; $item = $null
    try {
        $sheets = $Workbook.Worksheets
        $sheet = $sheets.Item($SheetName)
        $objects = $sheet.ChartObjects()
        for ($index = 1; $index -le $objects.Count; $index++) {
            $object = $objects.Item($index)
            try {
                $chart = $object.Chart
                $text = ''
                try { if ($chart.HasTitle) { $text = [string]$chart.ChartTitle.Text } } catch { $text = '' }
                if ($text -cne $Title) { continue }
                $series = $chart.SeriesCollection()
                $item = $series.Item(1)
                return @($item.XValues)
            } finally {
                if ($null -ne $item)   { Release-Transient $item   'Series';           $item   = $null }
                if ($null -ne $series) { Release-Transient $series 'SeriesCollection'; $series = $null }
                if ($null -ne $chart)  { Release-Transient $chart  'Chart';            $chart  = $null }
                if ($null -ne $object) { Release-Transient $object 'ChartObject';      $object = $null }
            }
        }
        return @()
    } finally {
        if ($null -ne $objects) { Release-Transient $objects 'ChartObjects'; $objects = $null }
        if ($null -ne $sheet)   { Release-Transient $sheet   'Worksheet';    $sheet   = $null }
        if ($null -ne $sheets)  { Release-Transient $sheets  'Worksheets';   $sheets  = $null }
    }
}


# ===========================================================================
# THE PROJECTIONS, READ THE ONE WAY THEY ARE MEANT TO BE READ
# ===========================================================================
$manifestPath   = Join-Path $BuildDir 'stage_b_manifest.json'
$inspectPath    = Join-Path $BuildDir 'phase5_gate_b_inspection.json'
$simInspectPath = Join-Path $BuildDir 'phase6_gate_b_inspection.json'
$casesPath      = Join-Path $BuildDir 'phase7_acceptance_cases.json'
$chartPath      = Join-Path $BuildDir 'phase8_charts_inspection.json'

$manifest   = Get-Content -LiteralPath $manifestPath   -Raw | ConvertFrom-Json
$inspection = Get-Content -LiteralPath $inspectPath    -Raw | ConvertFrom-Json
$simInspect = Get-Content -LiteralPath $simInspectPath -Raw | ConvertFrom-Json
$cases      = Get-Content -LiteralPath $casesPath      -Raw | ConvertFrom-Json
$charts     = Get-Content -LiteralPath $chartPath      -Raw | ConvertFrom-Json
$zeroVarianceStatus = [string]$charts.zero_variance_status

$case = $null
foreach ($scenario in @($cases.scenarios)) { if ([string]$scenario.id -ceq 'W4') { $case = $scenario } }
if ($null -eq $case) { throw 'the acceptance corpus carries no W4 scenario' }
$model = $case.model

$source = $charts.sensitivity_source
$eligibilityColumn = [string]$source.eligibility.column
$sensitivitySheet = [string]$charts.sensitivity_sheet
$sourceFirstRow = [int]$source.first_row
$drivers = $charts.bridge.drivers
$tornado = $null
foreach ($spec in @($charts.charts)) { if ([string]$spec.key -ceq 'tornado') { $tornado = $spec } }
if ($null -eq $tornado) { throw 'the chart projection carries no tornado' }

# THE SOURCE REVISION, BEFORE ANYTHING ELSE HAPPENS.
#
# RUN 1 DIED HERE and never reached Excel: this called the helper with NO
# -RepoRoot, so it took the parameter's empty default and asked git to report
# HEAD for ''. The accepted runners pass the derived root explicitly and that is
# what this does now.
#
# AND THE REFUSAL THAT WAS MISSING ENTIRELY. The accepted runners will not run
# against a modified pccm/src, pccm/spec or pccm/builder, because a result from
# a tree that is not a commit cannot be attributed to one. This runner had no
# such check at all - a worse defect than the empty root, because it would have
# produced a green report rather than an error.
$revision = $null
try { $revision = Get-P8ZSourceRevision -RepoRoot $repoRoot }
catch { Write-Host (Format-Err $_) -ForegroundColor Red; exit 1 }
if ($revision.Dirty.Count -gt 0) {
    Write-Host 'REFUSED, BEFORE EXCEL WAS STARTED.' -ForegroundColor Red
    Write-Host ''
    Write-Host ('pccm/src, pccm/spec or pccm/builder is modified, so a P8-Z result ' +
                'could not be attributed to a source revision:') -ForegroundColor Red
    foreach ($line in $revision.Dirty) { Write-Host ('    ' + $line) -ForegroundColor Red }
    exit 1
}

$stageBPath = Join-Path $BuildDir 'PCCM_stageA.xlsx'

# AND EVERY FILE IT IS ABOUT TO READ EXISTS, named one at a time. A missing
# artefact is a Stage A that was not built, and saying so here costs a second
# rather than a Windows turn.
foreach ($required in @($manifestPath, $inspectPath, $simInspectPath, $casesPath,
                        $chartPath, $stageBPath)) {
    if (-not (Test-Path -LiteralPath $required)) {
        Write-Host ('REFUSED, BEFORE EXCEL WAS STARTED: ' + $required +
                    ' does not exist. Build Stage A first.') -ForegroundColor Red
        exit 1
    }
}

Write-P8ZLine 'PCCM - PHASE 8 P8-Z: A ZERO-VARIANCE DRIVER IS NOT A TORNADO CATEGORY'
Write-P8ZLine '====================================================================='
Write-P8ZLine ''
Write-P8ZLine ('source revision   : ' + $revision.Head)
Write-P8ZLine ('fixture           : ' + [string]$case.id + ', one cost line made constant')
Write-P8ZLine ('eligibility field : ' + [string]$source.eligibility.key +
               ' at ' + $sensitivitySheet + '!' + $eligibilityColumn)
Write-P8ZLine ''

$ledger = New-ReleaseLedger
$excel = $null; $wb = $null; $workbooks = $null
$comAcquired = 0
$fatal = ''
try {
    $excel = New-Object -ComObject Excel.Application
    $comAcquired = $comAcquired + 1
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false

    $workbooks = $excel.Workbooks
    $comAcquired = $comAcquired + 1
    $wb = $workbooks.Open($stageBPath)
    $comAcquired = $comAcquired + 1

    $compileFailure = ''
    try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $compileFailure = (Format-Err $_) }
    $null = Add-P8ZCheck 'the current VBAProject compiles in real Excel' `
        ([string]::IsNullOrWhiteSpace($compileFailure)) $compileFailure 'PREREQUISITE'
    if (-not [string]::IsNullOrWhiteSpace($compileFailure)) { throw $compileFailure }

    $null = Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection

    # -------------------------------------------------------------------
    # THE MODEL: the accepted W4 fixture, then ONE DRIVER MADE CONSTANT
    # -------------------------------------------------------------------
    Write-P8ZLine 'THE MODEL - W4, WITH ONE COST LINE THAT CANNOT VARY'
    Write-P8ZLine '---------------------------------------------------'
    $applied = [string](Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest `
        -Inspection $inspection -Model $model)
    $null = Add-P8ZCheck 'the accepted W4 fixture applied' ($applied -like 'OK|*') $applied 'PREREQUISITE'

    $costRegister = Get-P8ZRegister -Manifest $manifest -Key 'cost_lines'
    $constant = @($model.cost_lines)[0]
    $constantId = [string]$constant.permanent_id
    $fixedValue = [double]$constant.min_value
    $constantRow = Get-P8ZRegisterRowIndex -Workbook $wb -Register $costRegister -PermanentId $constantId
    $null = Add-P8ZCheck 'the driver to be made constant is in the register' `
        ($constantRow -ge 1) ($constantId + ' at body row ' + [string]$constantRow) 'PREREQUISITE'
    if ($constantRow -lt 1) { throw ('P8-Z could not find ' + $constantId) }

    # MINIMUM, MOST LIKELY AND MAXIMUM ALL THE SAME. The unit cost cannot move,
    # so the contribution is identical in every iteration, so the accepted
    # kernel finds sxx = 0 and returns SIM_SENSITIVITY_NO_VARIANCE. Production
    # reaches the state; this only removes the driver's uncertainty.
    foreach ($columnKey in @('unit_cost_min', 'unit_cost_most_likely', 'unit_cost_max')) {
        $ordinal = 0
        try { $ordinal = Get-P8ZRegisterColumnIndex -Register $costRegister -ColumnKey $columnKey }
        catch { $ordinal = 0 }
        if ($ordinal -lt 1) { continue }
        Set-TableCell -Workbook $wb -SheetName $costRegister.sheet `
            -TableName $costRegister.table_name -RowIndex $constantRow `
            -ColumnIndex $ordinal -Value $fixedValue
    }
    Write-P8ZLine ('    ' + $constantId + ' fixed at ' + [string]$fixedValue +
                   ' across min, most likely and max')

    # -------------------------------------------------------------------
    # THE RUN
    # -------------------------------------------------------------------
    Write-P8ZLine ''
    Write-P8ZLine 'THE RUN - CALCULATE, SIMULATE, SENSITIVITY'
    Write-P8ZLine '-----------------------------------------'
    foreach ($endpoint in @('PCCM_Calculate', 'PCCM_RunSimulation')) {
        $result = Invoke-P8ZEndpoint -Excel $excel -Endpoint $endpoint
        $null = Add-P8ZCheck ($endpoint + ' succeeded') ($result -like 'OK|*') $result 'PREREQUISITE'
        if ($result -notlike 'OK|*') { throw ($endpoint + ' refused: ' + $result) }
    }
    $sensitivity = Invoke-P8ZEndpoint -Excel $excel -Endpoint ([string]$charts.sensitivity_endpoint)
    # (1) THE ENDPOINT SUCCEEDS.
    $null = Add-P8ZCheck ([string]$charts.sensitivity_endpoint + ' succeeded') `
        ($sensitivity -like 'OK|*') $sensitivity
    if ($sensitivity -notlike 'OK|*') { throw ('sensitivity refused: ' + $sensitivity) }
    $null = Invoke-P8ZRecalculate -Excel $excel -Stage 'after sensitivity'

    # -------------------------------------------------------------------
    # THE SENSITIVITY SHEET - THE DIAGNOSTIC ROW IS THERE, AND IT IS BLANK
    # -------------------------------------------------------------------
    Write-P8ZLine ''
    Write-P8ZLine 'THE SHEET - THE DIAGNOSTIC ROW IS RETAINED'
    Write-P8ZLine '------------------------------------------'
    $columns = @{}
    foreach ($column in @($source.columns)) { $columns[[string]$column.key] = [string]$column.column }
    $columns[[string]$source.eligibility.key] = $eligibilityColumn

    # THE WHOLE PUBLISHED WINDOW IS WALKED. Which row the diagnostic driver
    # landed on is the publication's business, not this runner's: it is found by
    # its NAME, not assumed to be last.
    $constantName = ''
    $constantSheetRow = 0
    $rankedNames = New-Object System.Collections.ArrayList
    $blankMeasures = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt [int]$drivers.row_count + 4; $index++) {
        $sheetRow = $sourceFirstRow + $index
        $status = Get-P8ZCell -Workbook $wb -SheetName $sensitivitySheet `
            -Address ($columns['status'] + [string]$sheetRow)
        if (Test-P8ZBlank -Cell $status) { continue }
        $name = Get-P8ZCell -Workbook $wb -SheetName $sensitivitySheet `
            -Address ($columns['driver_name'] + [string]$sheetRow)
        $rank = Get-P8ZCell -Workbook $wb -SheetName $sensitivitySheet `
            -Address ($columns['rank'] + [string]$sheetRow)
        Write-P8ZLine ('    row ' + [string]$sheetRow + ': ' + (Format-P8ZCell $name) +
                       ', status ' + (Format-P8ZCell $status) + ', rank ' + (Format-P8ZCell $rank))
        if ([string]$status.Value -ceq $zeroVarianceStatus) {
            $constantName = [string]$name.Value
            $constantSheetRow = $sheetRow
        } else {
            $null = $rankedNames.Add([string]$name.Value)
        }
    }

    # (2) AND (3) THE DIAGNOSTIC ROW IS PRESENT, AND SAYS WHY.
    $null = Add-P8ZCheck 'the zero-variance driver is still published on the Sensitivity sheet' `
        ($constantSheetRow -ge 1) ('found at row ' + [string]$constantSheetRow)
    if ($constantSheetRow -lt 1) {
        throw 'no row carried the zero-variance status; the fixture did not produce one'
    }
    $statusCell = Get-P8ZCell -Workbook $wb -SheetName $sensitivitySheet `
        -Address ($columns['status'] + [string]$constantSheetRow)
    $null = Add-P8ZCheck ('its status is ' + $zeroVarianceStatus) `
        ([string]$statusCell.Value -ceq $zeroVarianceStatus) `
        (Format-P8ZCell $statusCell)

    # (4) ITS MEASURES DISPLAY BLANK, NOT ZERO.
    foreach ($measure in @('rho', 'abs_rho', 'rank', 'direction')) {
        if (-not $columns.ContainsKey($measure)) { continue }
        $cell = Get-P8ZCell -Workbook $wb -SheetName $sensitivitySheet `
            -Address ($columns[$measure] + [string]$constantSheetRow)
        $isBlank = (Test-P8ZBlank -Cell $cell)
        if (-not $isBlank) { $null = $blankMeasures.Add((Format-P8ZCell $cell)) }
    }
    $null = Add-P8ZCheck 'its rho, |rho|, rank and direction are blank and not zero' `
        ($blankMeasures.Count -eq 0) (($blankMeasures -join '; '))

    # -------------------------------------------------------------------
    # THE BRIDGE, FROZEN BEFORE THE CHART IS LOOKED AT
    # -------------------------------------------------------------------
    $bridge = @{}
    foreach ($column in @($drivers.columns)) {
        $cells = New-Object System.Collections.ArrayList
        for ($index = 0; $index -lt [int]$drivers.row_count; $index++) {
            $null = $cells.Add((Get-P8ZCell -Workbook $wb -SheetName ([string]$charts.bridge_sheet) `
                -Address ([string]$column.column + [string]([int]$drivers.first_row + $index))))
        }
        $bridge[[string]$column.key] = @($cells)
    }

    # -------------------------------------------------------------------
    # THE CHART - EXCEL'S OWN CATEGORIES AND VALUES
    # -------------------------------------------------------------------
    Write-P8ZLine ''
    Write-P8ZLine 'THE CHART - WHAT EXCEL ACTUALLY PLOTS'
    Write-P8ZLine '-------------------------------------'
    $objects = Get-P8ZCharts -Workbook $wb -SheetName ([string]$charts.chart_sheet)
    $found = $null
    foreach ($object in @($objects)) {
        if ([string]$object.Title -ceq [string]$tornado.title) { $found = $object }
    }
    $null = Add-P8ZCheck 'the tornado chart object is on the Dashboard' `
        ($null -ne $found) ([string]@($objects).Count + ' chart objects')
    if ($null -eq $found) { throw 'the tornado chart object was not found' }

    # (8) THE WIRING IS THE PROJECTED BRIDGE, unchanged by this correction.
    $series = @($found.Series)[0]
    $null = Add-P8ZCheck 'the tornado category range is the projected bridge range' `
        ([string]$series.Categories -ceq [string]$tornado.categories.range) `
        ([string]$series.Categories)
    $null = Add-P8ZCheck 'the tornado value range is the projected bridge range' `
        ([string]$series.ValuesRange -ceq [string](@($tornado.series)[0].range)) `
        ([string]$series.ValuesRange)

    # (5), (6) AND (7). A category is drawn only where the bridge carries one,
    # and the bridge carries one only for a ranked driver.
    $drawnCategories = New-Object System.Collections.ArrayList
    $drawnValues = 0
    $names = @($bridge['driver_name'])
    $rhos = @($bridge['rho'])
    $points = @($series.Values)
    $categories = @(Get-P8ZCategoryValues -Workbook $wb `
        -SheetName ([string]$charts.chart_sheet) -Title ([string]$tornado.title))
    $mismatched = New-Object System.Collections.ArrayList
    for ($index = 0; $index -lt $names.Count; $index++) {
        $cellAbsent = ((Test-P8ZBlank -Cell $names[$index]) -or $names[$index].IsError)
        $pointAbsent = $true
        if ($index -lt $categories.Count) { $pointAbsent = (Test-P8ZNoPoint -Value $categories[$index]) }
        if ($cellAbsent -ne $pointAbsent) {
            $null = $mismatched.Add('row ' + [string]($index + 1) + ': bridge ' +
                                    (Format-P8ZCell $names[$index]) + ', plotted ' +
                                    (Format-P8ZPoint $categories[$index]))
        }
        if (-not $cellAbsent) { $null = $drawnCategories.Add([string]$names[$index].Value) }
        if ($index -lt $points.Count) {
            if (-not (Test-P8ZNoPoint -Value $points[$index])) { $drawnValues = $drawnValues + 1 }
        }
    }
    $null = Add-P8ZCheck 'every drawn category matches a bridge cell that carries one' `
        ($mismatched.Count -eq 0) (($mismatched -join '; '))
    Write-P8ZLine ('    categories drawn : ' + ($drawnCategories -join ', '))
    Write-P8ZLine ('    ranked drivers   : ' + ($rankedNames -join ', '))

    # (5) NOT A CATEGORY.
    $null = Add-P8ZCheck 'the zero-variance driver is not a tornado category' `
        (-not ($drawnCategories -ccontains $constantName)) `
        ($constantName + ' among: ' + ($drawnCategories -join ', '))
    # (6) NOT A BAR.
    $barred = $false
    for ($index = 0; $index -lt $names.Count; $index++) {
        if ([string]$names[$index].Value -cne $constantName) { continue }
        if ($index -lt $points.Count) {
            if (-not (Test-P8ZNoPoint -Value $points[$index])) { $barred = $true }
        }
    }
    $null = Add-P8ZCheck 'the zero-variance driver has no bar' (-not $barred) `
        ($constantName + ', ' + [string]$drawnValues + ' bars drawn')
    # (7) EVERY CATEGORY IS AN ELIGIBLE RANKED DRIVER, and there are exactly as
    #     many as the sheet ranked.
    $unexpected = New-Object System.Collections.ArrayList
    foreach ($drawn in $drawnCategories) {
        if (-not ($rankedNames -ccontains $drawn)) { $null = $unexpected.Add($drawn) }
    }
    $null = Add-P8ZCheck 'every tornado category is an eligible ranked driver' `
        ($unexpected.Count -eq 0) (($unexpected -join '; '))
    $null = Add-P8ZCheck 'the tornado draws exactly as many categories as were ranked' `
        ($drawnCategories.Count -eq $rankedNames.Count) `
        ([string]$drawnCategories.Count + ' drawn, ' + [string]$rankedNames.Count + ' ranked')
    $null = Add-P8ZCheck 'fewer than the projected top N are drawn, and nothing fills the rest' `
        (($drawnCategories.Count -lt [int]$drivers.row_count) -and
         ($drawnValues -eq $drawnCategories.Count)) `
        ([string]$drawnCategories.Count + ' of ' + [string]$drivers.row_count +
         ', ' + [string]$drawnValues + ' bars')

    if ($null -ne $series) { Release-Transient $series 'Series'; $series = $null }
    [System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()
    try { $excel.Run('PCCM_AutomationEnd') | Out-Null } catch { }
} catch {
    $fatal = (Format-Err $_)
    Write-P8ZLine ''
    Write-P8ZLine ('THE P8-Z SESSION DID NOT COMPLETE: ' + $fatal)
} finally {
    Invoke-P8ZRelease -Ledger $ledger -Excel ([ref]$excel) -Workbook ([ref]$wb) `
        -Workbooks ([ref]$workbooks) -Acquired $comAcquired
}

$failed = 0
foreach ($check in @($script:P8ZChecks)) { if ([string]$check.Status -ceq 'FAIL') { $failed = $failed + 1 } }
Write-P8ZLine ''
Write-P8ZLine ('P8-Z: ' + [string]@($script:P8ZChecks).Count + ' checked, ' +
               [string]$failed + ' failed')
$ok = (($failed -eq 0) -and [string]::IsNullOrWhiteSpace($fatal))
Write-P8ZLine ('P8-Z ' + $(if ($ok) { 'PASS' } else { 'FAIL' }))
Write-P8ZLine ''
Write-P8ZLine 'THIS RUNNER ANSWERS ONE QUESTION. It is not the P8-3 acceptance suite and'
Write-P8ZLine 'no result here re-establishes or disturbs that acceptance.'
if ($ok) { exit 0 } else { exit 1 }
