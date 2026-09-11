<#
.SYNOPSIS
    PCCM test-only harness: the fixture window's EXCEPTION CONTROL FLOW, executed.

.DESCRIPTION
    WHY THIS EXISTS. Every other control over the fixture window reads the
    runner's source and asserts what it SAYS. "Exactly one compensating
    P10FW_End is attempted, and only on the failing paths" is not a property of
    text - it is a property of what PowerShell does with a try/catch when a
    `return` is taken, when a post-open assertion throws, and when the rollback
    itself refuses. Asserting that from source is asserting a belief about the
    language.

    So this RUNS the real functions. It parses `phase10_benchmark.ps1`, lifts
    `Get-BenchmarkProtectionState`, `Assert-BenchmarkProtectionApplied`,
    `Open-BenchmarkFixtureWindow`, `Invoke-BenchmarkWindowRollback` and
    `Close-BenchmarkFixtureWindow` OUT OF THE RUNNER BY AST - their real bytes,
    not a copy - and executes them against a fake Excel that records every
    `Application.Run` it is given and answers from a scripted queue.

    THE CALLER'S OWN try/finally IS LIFTED THE SAME WAY, by locating its exact
    lines in the runner, so the scenarios that exercise a fixture success and a
    fixture failure run the shipping text rather than a restatement of it.

    Excel is never started and no workbook is opened. The fake answers strings.

.NOTES
    Prints one line per scenario: NAME | calls | outcome | message.
    Exit 0 always - the pytest control reads the lines and decides.
#>
param(
    [string]$Runner
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Runner)) {
    $Runner = Join-Path (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)) `
        'bootstrap/windows/phase10_benchmark.ps1'
}
$runnerPath = (Resolve-Path -LiteralPath $Runner).Path

# --- THE REAL FUNCTIONS, BY AST ---------------------------------------------
$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $runnerPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the runner')
    exit 0
}
$wanted = @('Get-BenchmarkProtectionState', 'Assert-BenchmarkProtectionApplied',
            'Open-BenchmarkFixtureWindow', 'Invoke-BenchmarkWindowRollback',
            'Close-BenchmarkFixtureWindow')
$found = @{}
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    if ($wanted -contains $fn.Name) { $found[$fn.Name] = $fn.Extent.Text }
}
foreach ($name in $wanted) {
    if (-not $found.ContainsKey($name)) {
        Write-Output ('MISSING|' + $name + ' is not defined in the runner')
        exit 0
    }
    Invoke-Expression $found[$name]
}

# --- THE CALLER'S OWN try/finally, LIFTED FROM THE RUNNER --------------------
# Located by its first and last statements so a reformat of what is between them
# cannot silently turn this into a test of something else.
$source = Get-Content -LiteralPath $runnerPath -Raw
$openMark = '    $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest'
$closeMark = '        $protectionAfter = Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest'
if (($source -notmatch [regex]::Escape($openMark)) -or
    ($source -notmatch [regex]::Escape($closeMark))) {
    Write-Output 'MISSING|the caller region anchors are not in the runner'
    exit 0
}
$from = $source.IndexOf($openMark)
$to = $source.IndexOf($closeMark) + $closeMark.Length
$tail = $source.Substring($to)
$callerRegion = $source.Substring($from, $to - $from) + $tail.Substring(0, $tail.IndexOf('}') + 1)

# --- THE FAKE ----------------------------------------------------------------
$script:Calls = New-Object System.Collections.ArrayList
$script:Replies = @{}

function New-FakeExcel {
    $fake = New-Object psobject
    Add-Member -InputObject $fake -MemberType ScriptMethod -Name Run -Value {
        param([string]$Macro)
        $null = $script:Calls.Add($Macro)
        if (-not $script:Replies.ContainsKey($Macro)) {
            throw ('the fake was asked for ' + $Macro + ', which no scenario scripted')
        }
        $queue = $script:Replies[$Macro]
        if ($queue.Count -eq 0) {
            throw ('the fake ran out of scripted replies for ' + $Macro)
        }
        $next = $queue[0]
        if ($queue.Count -gt 1) { $script:Replies[$Macro] = $queue[1..($queue.Count - 1)] }
        if ($next -eq '<<raise>>') { throw ('the fake was told to raise on ' + $Macro) }
        return $next
    }
    return $fake
}

# Fourteen declared worksheets, the shape `$Manifest.protection.sheets` has.
function New-FakeManifest {
    $names = @()
    for ($i = 1; $i -le 14; $i++) { $names += ('Sheet' + [string]$i) }
    return [pscustomobject]@{ protection = [pscustomobject]@{ sheets = $names } }
}

function New-State {
    param([string]$Applied = 'True', [int]$Depth = 0, [string]$Structure = 'True',
          [int]$Sheets = 14, [int]$Protected = 14)
    return ('OK|applied=' + $Applied + '|depth=' + [string]$Depth +
            '|structure=' + $Structure + '|sheets=' + [string]$Sheets +
            '|protected=' + [string]$Protected)
}

# --- THE SCENARIOS -----------------------------------------------------------
# Each names what the fake answers, and whether the lifted caller region runs.
$scenarios = @(
    @{ Name = 'begin-refuses'
       Begin = @('FAIL|the sheets would not release'); End = @(); State = @((New-State))
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-depth-wrong'
       Begin = @('OK|depth=2'); End = @('OK|depth=0')
       State = @((New-State), (New-State -Depth 2))
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-state-read-refuses'
       Begin = @('OK|depth=1'); End = @('OK|depth=0')
       State = @((New-State), 'FAIL|the protection state could not be read')
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-state-read-raises'
       Begin = @('OK|depth=1'); End = @('OK|depth=0')
       State = @((New-State), '<<raise>>')
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-structure-false'
       Begin = @('OK|depth=1'); End = @('OK|depth=0')
       State = @((New-State), (New-State -Depth 1 -Structure 'False'))
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-fails-and-rollback-refuses'
       Begin = @('OK|depth=1'); End = @('FAIL|protection could not be re-applied to Setup')
       State = @((New-State), (New-State -Depth 2))
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'post-open-fails-and-rollback-raises'
       Begin = @('OK|depth=1'); End = @('<<raise>>')
       State = @((New-State), (New-State -Depth 2))
       RunCaller = $false; FixtureThrows = $false },
    @{ Name = 'open-succeeds-fixture-succeeds'
       Begin = @('OK|depth=1'); End = @('OK|depth=0')
       State = @((New-State), (New-State -Depth 1), (New-State))
       RunCaller = $true; FixtureThrows = $false },
    @{ Name = 'open-succeeds-fixture-throws'
       Begin = @('OK|depth=1'); End = @('OK|depth=0')
       State = @((New-State), (New-State -Depth 1), (New-State))
       RunCaller = $true; FixtureThrows = $true },
    @{ Name = 'open-succeeds-close-refuses'
       Begin = @('OK|depth=1'); End = @('FAIL|protection was not restored')
       State = @((New-State), (New-State -Depth 1))
       RunCaller = $true; FixtureThrows = $false },
    @{ Name = 'open-succeeds-close-leaves-depth-open'
       Begin = @('OK|depth=1'); End = @('OK|depth=1')
       State = @((New-State), (New-State -Depth 1), (New-State -Depth 1))
       RunCaller = $true; FixtureThrows = $false }
)

foreach ($scenario in $scenarios) {
    $script:Calls = New-Object System.Collections.ArrayList
    $script:Replies = @{
        'P10FW_Begin' = @($scenario.Begin)
        'P10FW_End'   = @($scenario.End)
        'P10FW_State' = @($scenario.State)
    }
    $excel = New-FakeExcel
    $manifest = New-FakeManifest
    # The lifted caller region's own variables.
    $wb = 'workbook'; $inspection = 'inspection'; $model = 'model'
    $protectionAfter = $null
    $script:FixtureThrows = [bool]$scenario.FixtureThrows
    $script:Timed = 0

    function Set-Phase5Fixture {
        param($Excel, $Workbook, $Manifest, $Inspection, $Model)
        if ($script:FixtureThrows) { throw 'THE FIXTURE RAISED' }
        return 'OK|fixture'
    }

    $outcome = 'RETURNED'
    $message = ''
    try {
        if ([bool]$scenario.RunCaller) {
            Invoke-Expression $callerRegion
        } else {
            $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest
        }
        # Stands in for the timed run loop, which sits after the close.
        $script:Timed = 1
    } catch {
        $outcome = 'THREW'
        $message = ([string]$_.Exception.Message) -replace '\s+', ' '
    }

    # TAGGED. A damaged runner under mutation may write to the host - `Write-Host`
    # lands on stdout - and an untagged line would then be parsed as a scenario
    # result, turning a caught mutation into an unreadable one.
    Write-Output ('FLOW|' + $scenario.Name + '|' + (($script:Calls) -join ',') + '|' + $outcome +
                  '|timed=' + [string]$script:Timed + '|' + $message)
}
exit 0
