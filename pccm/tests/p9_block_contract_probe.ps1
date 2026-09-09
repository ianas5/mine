<#
.SYNOPSIS
    PCCM test-only probe: the Phase-9 block reader's scalar contract.

.DESCRIPTION
    WHY THIS EXISTS. Windows run 2 died on
    `Cannot convert the "System.Object[]" value ... to type "System.Double"`
    after nine passing checks. `Get-P9Block` returned Value2 directly; PowerShell
    ENUMERATES a rank-2 array on its way out of a function, so the caller held a
    FLAT array, and `$block[$row, 1]` on a flat array is multi-index SELECTION -
    it returns two elements, silently. Every reading looked plausible until the
    first numeric cast.

    THE FUNCTIONS UNDER TEST ARE EXTRACTED FROM THE RUNNER ITSELF at test time
    and spliced in below, so this cannot drift into testing a copy. Excel is
    modelled by a fake workbook whose Value2 is a NOTE property - a script
    property would emit through the pipeline and flatten the array inside the
    fixture, which would make the probe prove itself.

.NOTES
    Print PASS/FAIL per assertion. Exit 0 and print ALL CLEAN, or exit 1.
#>
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
function Release-Transient { param($Obj, [string]$Label) }

# A FAKE WORKBOOK THAT MODELS COM PROPERTY READS. Value2 and Worksheets are
# NOTE properties, not script properties: a script property emits through the
# pipeline and would flatten a rank-2 array inside the fixture itself, which
# would make this control prove the fixture rather than the runner.
function New-FakeWorkbook {
    param($Value)
    $range = New-Object psobject
    $range | Add-Member -MemberType NoteProperty -Name Value2 -Value $Value
    $ws = New-Object psobject
    $ws | Add-Member -MemberType ScriptMethod -Name Range -Value { param($a) $script:TheRange }
    $sheets = New-Object psobject
    $sheets | Add-Member -MemberType ScriptMethod -Name Item -Value { param($n) $script:TheSheet }
    $wb = New-Object psobject
    $wb | Add-Member -MemberType NoteProperty -Name Worksheets -Value $sheets
    $script:TheRange = $range
    $script:TheSheet = $ws
    return $wb
}

function New-Rect {
    param([int]$Rows, [int]$Cols)
    $a = [Array]::CreateInstance([object], [int[]]@($Rows, $Cols), [int[]]@(1, 1))
    for ($r = 1; $r -le $Rows; $r++) {
        for ($c = 1; $c -le $Cols; $c++) { $a.SetValue(($r * 10 + $c), $r, $c) }
    }
    return ,$a
}

#<FUNCTIONS-FROM-THE-RUNNER>

$script:P9ErrorCodes = @{ -2146826273 = '#VALUE!'; -2146826246 = '#N/A' }
$fails = 0
function Check { param([string]$Label, [bool]$Ok, [string]$Detail = '')
    if ($Ok) { "PASS $Label" } else { $script:fails++; "FAIL $Label -- $Detail" } }

# 1. A MULTI-CELL RECTANGLE SURVIVES THE RETURN WITH ITS RANK.
$wb = New-FakeWorkbook (New-Rect -Rows 7 -Cols 1)
$block = Get-P9Block -Workbook $wb -SheetName 'S' -Address 'D8:D14'
Check 'the block keeps rank 2 through the return' ($block.Rect.Rank -eq 2) ([string]$block.Rect.Rank)
Check 'the first row reads its own value' ((Get-P9BlockCell -Block $block -Row 1 -Column 1) -eq 11)
Check 'the last row reads its own value' ((Get-P9BlockCell -Block $block -Row 7 -Column 1) -eq 71)
$scalarOk = $true
try { $null = [double](Get-P9BlockCell -Block $block -Row 3 -Column 1) } catch { $scalarOk = $false }
Check 'a numeric cell converts to Double' $scalarOk

# 2. A WIDE RECTANGLE READS BY ROW AND COLUMN.
$wide = New-FakeWorkbook (New-Rect -Rows 3 -Cols 6)
$wideBlock = Get-P9Block -Workbook $wide -SheetName 'S' -Address 'B18:G20'
Check 'a wide block reads column 6' ((Get-P9BlockCell -Block $wideBlock -Row 2 -Column 6) -eq 26)

# 3. THE RUN-2 SHAPE IS REFUSED, NOT SILENTLY MIS-READ.
$flat = New-FakeWorkbook (New-Object 'object[]' 7)
$flatBlock = Get-P9Block -Workbook $flat -SheetName 'S' -Address 'D8:D14'
$refused = ''
try { $null = Get-P9BlockCell -Block $flatBlock -Row 3 -Column 1 } catch { $refused = $_.Exception.Message }
Check 'a flattened block is refused' ($refused -like '*rank 1*') $refused
Check 'the refusal names the sheet and address' ($refused -like '*S!D8:D14*') $refused

# 4. A SINGLE CELL IS A SCALAR, AND ONLY (1,1) EXISTS IN IT.
$single = New-FakeWorkbook 'Check ID'
$singleBlock = Get-P9Block -Workbook $single -SheetName 'S' -Address 'B17'
Check 'a single cell reads back as itself' ((Get-P9BlockCell -Block $singleBlock -Row 1 -Column 1) -ceq 'Check ID')
$outside = ''
try { $null = Get-P9BlockCell -Block $singleBlock -Row 2 -Column 1 } catch { $outside = $_.Exception.Message }
Check 'a single cell refuses a second coordinate' ($outside -like '*only (1,1) exists*') $outside

# 5. AN EXCEL ERROR KEEPS ITS IDENTITY AND IS NOT A NUMBER.
$errRect = [Array]::CreateInstance([object], [int[]]@(1,1), [int[]]@(1,1))
$errRect.SetValue([int](-2146826273), 1, 1)
$errBlock = Get-P9Block -Workbook (New-FakeWorkbook $errRect) -SheetName 'S' -Address 'D9'
$errCell = Get-P9BlockCell -Block $errBlock -Row 1 -Column 1
Check 'an Excel error survives as an error' (Test-P9Error $errCell) ([string]$errCell)
Check 'an Excel error formats as its word' ((Format-P9Cell $errCell) -ceq '#VALUE!') (Format-P9Cell $errCell)

# 6. A BLANK IS A BLANK, NOT A ZERO.
$blankRect = [Array]::CreateInstance([object], [int[]]@(1,1), [int[]]@(1,1))
$blankBlock = Get-P9Block -Workbook (New-FakeWorkbook $blankRect) -SheetName 'S' -Address 'D10'
$blankCell = Get-P9BlockCell -Block $blankBlock -Row 1 -Column 1
Check 'a blank comes back as null' ($null -eq $blankCell) ([string]$blankCell)

# 7. A CELL HOLDING AN ARRAY IS REFUSED RATHER THAN HAVING ITS FIRST ITEM TAKEN.
$nested = [Array]::CreateInstance([object], [int[]]@(1,1), [int[]]@(1,1))
$nested.SetValue((New-Object 'object[]' 2), 1, 1)
$nestedBlock = Get-P9Block -Workbook (New-FakeWorkbook $nested) -SheetName 'S' -Address 'D11'
$many = ''
try { $null = Get-P9BlockCell -Block $nestedBlock -Row 1 -Column 1 } catch { $many = $_.Exception.Message }
Check 'more than one value is refused' ($many -like '*2 values where exactly one*') $many

# 8. THE OLD SHAPE REPRODUCES THE RUN-2 ERROR EXACTLY, so this is the defect.
function Get-P9BlockOld {
    param($Workbook, [string]$SheetName, [string]$Address)
    $rng = $Workbook.Worksheets.Item($SheetName).Range($Address)
    return $rng.Value2
}
$old = Get-P9BlockOld -Workbook (New-FakeWorkbook (New-Rect -Rows 7 -Cols 1)) -SheetName 'S' -Address 'D8:D14'
Check 'the old return flattened the rectangle' ($old.Rank -eq 1) ([string]$old.Rank)
$selection = $old[3, 1]
Check 'the old index selected two values' ((@($selection).Count -eq 2)) ([string]@($selection).Count)
$run2 = ''
try { $null = [double]$selection } catch { $run2 = $_.Exception.Message }
Check 'the old shape raises the run-2 message' `
    ($run2 -eq 'Cannot convert the "System.Object[]" value of type "System.Object[]" to type "System.Double".') $run2

if ($fails -gt 0) { "FAILURES $fails"; exit 1 }
"ALL CLEAN"
exit 0
