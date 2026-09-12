<#
.SYNOPSIS
    PCCM test-only harness: RESERVED CAPACITY vs SEMANTIC COUNT, executed.

.DESCRIPTION
    WHY THIS EXISTS. Equivalence run 2 failed in the Bulk pass with

      tblCostLines already holds 25 body rows where the fixture needs 12

    because `Set-BenchmarkRegisterRowCount` read its argument as the number of
    body rows the table should END UP with. Stage A builds the register with
    `reserved_rows: 25`, and twelve Cost Lines occupying twelve of them is not an
    error - it is the state production reaches. `modDrivers.AddDriver` takes a
    blank RESERVED row and grows the table only when none is left.

    "How many rows does this end up with, and how many Adds did it take" is a
    COUNT, not a property of text. So this RUNS the real functions: it lifts
    `Set-BenchmarkRegisterRowCount`, `Get-BenchmarkPermanentId` and
    `New-BenchmarkRegisterBlock` out of `bootstrap/windows/phase10_benchmark.ps1`
    BY AST and drives them against a fake whose `ListRows.Add()` is counted.

    Excel is never started and no workbook is opened.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      ROWS|<case>|<capacity>|<needed>|<returned>|<adds>|<outcome>
      BLOCK|<case>|<rows>|<cols>|<first id>|<last id>|<counter>|<ids beyond N>
      BLANK|<case>|<blank cells in the block's unpopulated columns>
    Exit 0 always.
#>
param(
    [string]$Runner
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path (Split-Path -Parent $here) 'bootstrap/windows/phase10_benchmark.ps1'
}
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}
$wanted = @('Set-BenchmarkRegisterRowCount', 'Get-BenchmarkPermanentId',
            'New-BenchmarkRegisterBlock')
foreach ($name in $wanted) {
    $body = $null
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($fn.Name -eq $name) { $body = $fn.Extent.Text }
    }
    if ($null -eq $body) {
        Write-Output ('MISSING|' + $name + ' is not defined in the runner')
        exit 0
    }
    Invoke-Expression $body
}

# --- THE FAKE ----------------------------------------------------------------
# The grower reads the row count through `Get-TableRowCount` and grows through
# COM. Both are stubbed: the count comes from a script variable the fake's
# `Add()` increments, so the returned value and the number of Adds are both
# observable without Excel.
$script:PhysicalRows = 0
$script:Adds = 0

function Get-TableRowCount {
    param($Workbook, [string]$SheetName, [string]$TableName)
    return [int]$script:PhysicalRows
}

function Release-Transient {
    param($Target, [string]$Description)
}

function New-FakeWorkbook {
    $listRows = New-Object psobject
    Add-Member -InputObject $listRows -MemberType ScriptMethod -Name Add -Value {
        $script:Adds = $script:Adds + 1
        $script:PhysicalRows = $script:PhysicalRows + 1
        return (New-Object psobject)
    }
    $listObject = New-Object psobject
    Add-Member -InputObject $listObject -MemberType NoteProperty -Name ListRows -Value $listRows
    $listObjects = New-Object psobject
    Add-Member -InputObject $listObjects -MemberType ScriptMethod -Name Item -Value {
        param($Key)
        return $script:FakeListObject
    }
    $script:FakeListObject = $listObject
    $sheet = New-Object psobject
    Add-Member -InputObject $sheet -MemberType NoteProperty -Name ListObjects -Value $listObjects
    $sheets = New-Object psobject
    Add-Member -InputObject $sheets -MemberType ScriptMethod -Name Item -Value {
        param($Key)
        return $script:FakeSheet
    }
    $script:FakeSheet = $sheet
    $workbook = New-Object psobject
    Add-Member -InputObject $workbook -MemberType NoteProperty -Name Worksheets -Value $sheets
    return $workbook
}

# --- SECTION A: CAPACITY vs SEMANTIC COUNT -----------------------------------
# The contracted reserved capacity of every benchmark-populated table, and the
# semantic counts the three scenarios need.
$cases = @(
    @{ Case = 'small-cost-12-into-25';    Capacity = 25; Needed = 12 }
    @{ Case = 'small-risk-8-into-25';     Capacity = 25; Needed = 8 }
    @{ Case = 'exactly-at-capacity';      Capacity = 25; Needed = 25 }
    @{ Case = 'one-past-capacity';        Capacity = 25; Needed = 26 }
    @{ Case = 'medium-cost-60-into-25';   Capacity = 25; Needed = 60 }
    @{ Case = 'large-cost-180-into-25';   Capacity = 25; Needed = 180 }
    @{ Case = 'large-risk-120-into-25';   Capacity = 25; Needed = 120 }
)
foreach ($case in $cases) {
    $script:PhysicalRows = [int]$case.Capacity
    $script:Adds = 0
    $workbook = New-FakeWorkbook
    $returned = '<raised>'
    $outcome = 'returned'
    try {
        $returned = [string](Set-BenchmarkRegisterRowCount -Workbook $workbook `
            -SheetName 'Cost Lines' -TableName 'tblCostLines' -MinimumRows ([int]$case.Needed))
    } catch {
        $outcome = 'RAISED: ' + (($_.Exception.Message) -replace '\s+', ' ')
    }
    Write-Output ('ROWS|' + [string]$case.Case + '|' + [string]$case.Capacity + '|' +
                  [string]$case.Needed + '|' + $returned + '|' + [string]$script:Adds +
                  '|' + $outcome)
}

# --- SECTION B: THE REGISTER BLOCK ------------------------------------------
# The block must cover exactly the semantic rows, carry identifiers only on them,
# and leave the counter at the last issued sequence. The manifest's own counter
# projection supplies the prefix and pad width.
function New-FakeRegister {
    param([string]$Table, [string[]]$Columns)
    return [pscustomobject]@{ table_name = $Table; sheet = 'sheet'; columns = $Columns }
}
$costColumns = @('cost_line_id', 'category', 'description', 'uom', 'quantity', 'currency',
                 'inflation_profile', 'unit_cost_min', 'unit_cost_most_likely',
                 'unit_cost_max', 'distribution')
$costCounter = [pscustomobject]@{ defined_name = 'nmCounterCostLine'; prefix = 'CL-'
                                  pad_width = 3 }

function New-FakeDrivers {
    param([int]$Count)
    $drivers = @()
    for ($i = 1; $i -le $Count; $i++) {
        $drivers += [pscustomobject]@{
            permanent_id = ('CL-' + ([string]$i).PadLeft(3, [char]48))
            distribution = 'Triangular'; currency = 'SAR'; inflation_profile = 'Standard'
            min_value = [double]100; most_likely = [double]135; max_value = [double]210
            quantity = [double]1
        }
    }
    return $drivers
}

foreach ($count in @(12, 8, 180)) {
    $drivers = @(New-FakeDrivers -Count $count)
    $prepared = New-BenchmarkRegisterBlock -Register (New-FakeRegister -Table 'tblCostLines' `
        -Columns $costColumns) -Counter $costCounter -Drivers $drivers -IsRisk $false
    $block = $prepared.Block
    $rows = [int]$block.GetLength(0)
    $cols = [int]$block.GetLength(1)
    $ids = @($prepared.Ids)
    $beyond = 0
    foreach ($id in $ids) {
        $sequence = [int]($id.Substring(3))
        if ($sequence -gt $count) { $beyond = $beyond + 1 }
    }
    Write-Output ('BLOCK|' + [string]$count + '|' + [string]$rows + '|' + [string]$cols + '|' +
                  [string]$ids[0] + '|' + [string]$ids[$ids.Count - 1] + '|' +
                  [string]$prepared.CounterValue + '|' + [string]$beyond)

    # THE COLUMNS NO DRIVER FILLS MUST BE GENUINELY BLANK. `category` and `uom`
    # are blank in the accepted endpoint-built register too, and writing '' there
    # instead of $null would make them populated.
    $blank = 0
    $populated = 0
    foreach ($name in @('category', 'uom')) {
        $ordinal = [array]::IndexOf($costColumns, $name)
        for ($r = 0; $r -lt $rows; $r++) {
            if ($null -eq $block[$r, $ordinal]) { $blank = $blank + 1 }
            else { $populated = $populated + 1 }
        }
    }
    Write-Output ('BLANK|' + [string]$count + '|' + [string]$blank + '|' + [string]$populated)
}
exit 0
