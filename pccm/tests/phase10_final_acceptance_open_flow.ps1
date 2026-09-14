<#
.SYNOPSIS
    Executes the runner's OPEN/READY boundary - Wait-FaWorkbookReady - and its
    fatal diagnostics - Format-FaFatalLines - lifted out of the runner by AST,
    beside the real shared retry authority (com_lifecycle.ps1). Excel is never
    started. IT ASSERTS NOTHING: it prints OPEN|<case>|<text> lines and the
    Python control decides.

    WHAT LINUX CAN AND CANNOT MODEL. On Windows a refused COM property read
    surfaces as a terminating COMException, which is what the authority classifies
    and retries. PowerShell 7 on Linux answers null for a throwing .NET or script
    property getter and raises nothing, so the barrier's property reads cannot be
    made to throw here. What IS executed: the classification of a transient
    rejection against a non-transient failure; the authority's own retry and
    rethrow through a throwing METHOD read; the barrier's bounded outer retry on
    a workbook that answers nothing; the identity and not-a-count errors that
    escape the loop; readiness at once; and the fatal formatter on a real record.
#>
param([string]$Runner, [string]$Lifecycle)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

. $Lifecycle
$source = Get-Content -LiteralPath $Runner -Raw
$ast = [System.Management.Automation.Language.Parser]::ParseInput($source, [ref]$null, [ref]$null)
foreach ($name in @('Wait-FaWorkbookReady', 'Format-FaFatalLines')) {
    $definition = $ast.FindAll({ param($node)
        ($node -is [System.Management.Automation.Language.FunctionDefinitionAst]) -and ($node.Name -eq $name) }, $true)
    if (@($definition).Count -ne 1) { throw ('the runner defines ' + $name + ' ' + [string]@($definition).Count + ' times') }
    . ([scriptblock]::Create(@($definition)[0].Extent.Text))
}

function Emit { param([string]$Case, [string]$Text) Write-Output ('OPEN|' + $Case + '|' + $Text) }

# A STAND-IN WORKBOOK: FullName answers at once; Worksheets.Count answers nothing
# for the first $Silent reads, then the count.
function New-StandInWorkbook {
    param([string]$Path, $Sheets, [int]$Silent)
    $state = @{ Sheets = $Sheets; Silent = $Silent; Calls = 0 }
    $wb = New-Object PSObject
    $wb | Add-Member -MemberType NoteProperty -Name FullName -Value $Path
    $wb | Add-Member -MemberType NoteProperty -Name State -Value $state
    $wb | Add-Member -MemberType ScriptProperty -Name Worksheets -Value {
        $this.State.Calls = $this.State.Calls + 1
        $sheets = New-Object PSObject
        if ($this.State.Silent -gt 0) { $this.State.Silent = $this.State.Silent - 1; $sheets | Add-Member -MemberType NoteProperty -Name Count -Value $null }
        else { $sheets | Add-Member -MemberType NoteProperty -Name Count -Value $this.State.Sheets }
        return $sheets
    }
    return $wb
}

$rejected = New-Object System.Runtime.InteropServices.COMException('Call was rejected by callee.', -2147418111)
$path = 'C:\pccm\build\stage_b\PCCM_stageB.xlsm'

# 1. CLASSIFICATION, on real error records.
$record = $null; try { throw $rejected } catch { $record = $_ }
Emit 'classify.transient' (Get-ComRejectionName $record)
$record = $null; try { throw (New-Object System.InvalidOperationException('the sheet is gone')) } catch { $record = $_ }
Emit 'classify.non-transient' ('<' + (Get-ComRejectionName $record) + '>')

# 2. THE SHARED AUTHORITY, through a throwing METHOD read (terminating on every platform).
function New-StandInItems { param([int]$Rejections, $Failure)
    $items = New-Object PSObject
    $items | Add-Member -MemberType NoteProperty -Name Left -Value $Rejections
    $items | Add-Member -MemberType NoteProperty -Name Failure -Value $Failure
    $items | Add-Member -MemberType ScriptMethod -Name Item -Value { param($Key)
        if ($this.Left -gt 0) { $this.Left = $this.Left - 1; throw $this.Failure }
        return ('sheet:' + [string]$Key) }
    return $items
}
$read = Invoke-ComRetryRead -Target (New-StandInItems -Rejections 2 -Failure $rejected) -Member 'Item' -Key 'Setup' -Description 'a stand-in Item' -FirstDelayMs 10 -MaxDelayMs 20
Emit 'authority.transient-retried' ('value=' + [string]$read.Value + '|attempts=' + [string]$read.Attempts + '|rejections=' + $read.Rejections)
$escaped = ''
try { $null = Invoke-ComRetryRead -Target (New-StandInItems -Rejections 1 -Failure (New-Object System.InvalidOperationException('the sheet is gone'))) -Member 'Item' -Key 'Setup' -Description 'a stand-in Item' }
catch { $escaped = $_.Exception.GetType().Name + ': ' + $_.Exception.Message }
Emit 'authority.non-transient-escapes' $escaped

# 3. THE BARRIER: at once, after silence, never, wrong workbook, not a count.
$ready = Wait-FaWorkbookReady -Workbook (New-StandInWorkbook -Path $path -Sheets 14 -Silent 0) -ExpectedPath $path
Emit 'ready.at-once' ('ready=' + [string]$ready.Ready + '|attempts=' + [string]$ready.Attempts + '|waited=' + [string]$ready.WaitedMs + '|sheets=' + [string]$ready.Sheets)
$ready = Wait-FaWorkbookReady -Workbook (New-StandInWorkbook -Path $path -Sheets 14 -Silent 2) -ExpectedPath $path -FirstDelayMs 10 -MaxDelayMs 20
Emit 'ready.after-silence' ('ready=' + [string]$ready.Ready + '|attempts=' + [string]$ready.Attempts + '|waited=' + [string]$ready.WaitedMs + '|sheets=' + [string]$ready.Sheets)
$never = New-StandInWorkbook -Path $path -Sheets 14 -Silent 1000
$ready = Wait-FaWorkbookReady -Workbook $never -ExpectedPath $path -MaxAttempts 3 -FirstDelayMs 10 -MaxDelayMs 20 -TotalBudgetMs 100
Emit 'never-ready.bounded' ('ready=' + [string]$ready.Ready + '|attempts=' + [string]$ready.Attempts + '|waited=' + [string]$ready.WaitedMs + '|fullname=' + $ready.NameState + '|worksheets=' + $ready.SheetState + '|calls=' + [string]$never.State.Calls)
$wrong = ''
try { $null = Wait-FaWorkbookReady -Workbook (New-StandInWorkbook -Path 'C:\elsewhere\other.xlsm' -Sheets 14 -Silent 0) -ExpectedPath $path }
catch { $wrong = $_.Exception.Message }
Emit 'wrong-workbook.error' $wrong
$notCount = ''
try { $null = Wait-FaWorkbookReady -Workbook (New-StandInWorkbook -Path $path -Sheets 'many' -Silent 0) -ExpectedPath $path }
catch { $notCount = $_.Exception.Message }
Emit 'not-a-count.error' $notCount

# 4. THE FATAL DIAGNOSTICS, over a real error record from a real script line.
try { $nothing = $null; $nothing.Item('Setup') }
catch { foreach ($line in @(Format-FaFatalLines -ErrorRecord $_ -MaxStackLines 3)) { Emit 'fatal' $line } }
