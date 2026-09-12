<#
.SYNOPSIS
    PCCM test-only harness: the equivalence gate's STARTING BUNDLE, executed.

.DESCRIPTION
    WHY THIS EXISTS. Equivalence run 1 failed before Excel was started, in both
    passes, because the gate copied the Stage-A workbook and the generated `vba`
    directory into each disposable workdir and NOT `stage_b_manifest.json` - which
    is the first thing `build_stage_b.ps1` reads from the supplied `-BuildDir`.
    Every source-reading control passed; the defect was file plumbing, and file
    plumbing is behaviour.

    So this RUNS it. It lifts `Get-BundleArtifacts`, `New-EquivalenceBundle`,
    `Copy-EquivalenceWorkbook`, `Test-CopyIdentity` and the read-only readiness
    barrier `Wait-EquivalenceWorkbookReady` out of `tests/phase10_fixture_equivalence.ps1`
    BY AST - their real bytes, not a copy - and drives them against a FAKE
    repository build directory holding exactly the artifacts the contract names, a
    fake canonical workbook, and a fake workbook object whose reads are scripted.
    Excel is never started and no workbook is opened.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      ARTIFACT|<kind>|<name>              one entry of the derived contract
      BUNDLE|Baseline|<relative path>     one file that arrived in the one bundle
      COPY|<mode>|<path>                  one filesystem copy of the canonical workbook
      IDENTITY|<identical|differ>|<why>   canonical digest against both copies
      DAMAGED|<differ|identical>|<why>    the same comparison after one copy is edited
      READY|<case>|<ok|refused>|<detail>  the readiness barrier over a scripted workbook
      READYREADS|<case>|<members touched> what the barrier asked of the workbook
      REFUSE|<case>|<refused|accepted>|<message>
      STALE|<absent|travelled|refused>|<detail>
    Exit 0 always.
#>
param(
    [string]$Gate
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Gate)) {
    $Gate = Join-Path $here 'phase10_fixture_equivalence.ps1'
}
$gatePath = (Resolve-Path -LiteralPath $Gate).Path

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile(
    $gatePath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the gate')
    exit 0
}
# The accepted read envelope is the real one, not a restatement.
. (Join-Path (Split-Path -Parent $here) 'bootstrap/windows/com_lifecycle.ps1')
$wanted = @('Get-BundleArtifacts', 'New-EquivalenceBundle', 'Get-WorkbookDigest',
            'Copy-EquivalenceWorkbook', 'Test-CopyIdentity', 'Get-EquivalenceComparablePath',
            'Wait-EquivalenceWorkbookReady')
foreach ($name in $wanted) {
    $body = $null
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($fn.Name -eq $name) { $body = $fn.Extent.Text }
    }
    if ($null -eq $body) {
        Write-Output ('MISSING|' + $name + ' is not defined in the gate')
        exit 0
    }
    Invoke-Expression $body
}

# --- A FAKE REPOSITORY BUILD DIRECTORY --------------------------------------
# Exactly the artifacts the derived contract names, with recognisable contents.
# Nothing here is a real workbook: the question is which files arrive where.
function New-FakeBuild {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-fakebuild-' +
        [System.Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $root -Force
    Set-Content -LiteralPath (Join-Path $root 'stage_b_manifest.json') `
        -Value '{"fake":"manifest"}' -NoNewline
    Set-Content -LiteralPath (Join-Path $root 'PCCM_stageA.xlsx') `
        -Value 'fake-stage-a-workbook' -NoNewline
    $vba = Join-Path $root 'vba'
    $null = New-Item -ItemType Directory -Path $vba -Force
    foreach ($module in @('modConstants', 'modCalcContract', 'modSimContract')) {
        Set-Content -LiteralPath (Join-Path $vba ($module + '.bas')) `
            -Value ('Attribute VB_Name = "' + $module + '"') -NoNewline
    }
    return $root
}

# The manifest shape the bundle builder reads. The generated directory is declared
# as a repository-relative path whose LEAF is what lands in the bundle, which is
# how build_stage_b.ps1 resolves it.
$manifest = [pscustomobject]@{
    stage_a_filename = 'PCCM_stageA.xlsx'
    stage_b_filename = 'PCCM_stageB.xlsm'
    vba = [pscustomobject]@{ generated_dir = 'build/vba' }
}

foreach ($artifact in @(Get-BundleArtifacts -Manifest $manifest)) {
    Write-Output ('ARTIFACT|' + [string]$artifact.Kind + '|' + [string]$artifact.Name)
}

# --- TWO BUNDLES FROM ONE BUILD ---------------------------------------------
$BuildDir = New-FakeBuild
$WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-fakework-' +
    [System.Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $WorkDir -Force

$bundle = $null
try {
    $bundle = New-EquivalenceBundle -Mode 'Baseline' -Manifest $manifest -Stamp 'flow'
} catch {
    Write-Output ('REFUSE|one-bundle|refused|' + (($_.Exception.Message) -replace '\s+', ' '))
}
if ($null -ne $bundle) {
    foreach ($key in @($bundle.Digests.Keys)) { Write-Output ('BUNDLE|Baseline|' + $key) }
    # A FAKE CANONICAL WORKBOOK where the bootstrap would have left it, copied
    # twice through the real function, and the three digests compared.
    $stageB = [string]$bundle.StageB
    Set-Content -LiteralPath $stageB -Value 'fake-verified-stage-b-workbook' -NoNewline
    $canonical = Get-WorkbookDigest -Path $stageB
    $copies = @()
    try {
        foreach ($mode in @('Endpoints', 'Bulk')) {
            $copy = Copy-EquivalenceWorkbook -Source $stageB -Root ([string]$bundle.Root) -Mode $mode
            $copies += $copy
            Write-Output ('COPY|' + $mode + '|' + [string]$copy.Path)
        }
        $problems = @(Test-CopyIdentity -Canonical $canonical -Copies $copies)
        Write-Output ('IDENTITY|' + $(if ($problems.Count -eq 0) { 'identical' } else { 'differ' }) +
                      '|' + ($problems -join '; '))
        $victim = [string]$copies[1].Path
        Set-Content -LiteralPath $victim -Value 'fake-tampered-copy' -NoNewline
        $copies[1] = [pscustomobject]@{ Mode = 'Bulk'; Path = $victim; Digest = (Get-WorkbookDigest -Path $victim) }
        $damaged = @(Test-CopyIdentity -Canonical $canonical -Copies $copies)
        Write-Output ('DAMAGED|' + $(if ($damaged.Count -eq 0) { 'identical' } else { 'differ' }) +
                      '|' + ($damaged -join '; '))
    } catch {
        Write-Output ('IDENTITY|differ|copying failed: ' + (($_.Exception.Message) -replace '\s+', ' '))
    }
}
function Test-Refusal {
    param([string]$Case, [scriptblock]$Arrange)
    $script:BuildDir = New-FakeBuild
    $script:WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-fakework-' +
        [System.Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:WorkDir -Force
    & $Arrange
    try {
        $null = New-EquivalenceBundle -Mode 'Baseline' -Manifest $manifest -Stamp $Case
        Write-Output ('REFUSE|' + $Case + '|accepted|the bundle was built anyway')
    } catch {
        Write-Output ('REFUSE|' + $Case + '|refused|' + (($_.Exception.Message) -replace '\s+', ' '))
    }
    Remove-Item -LiteralPath $script:BuildDir -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $script:WorkDir -Recurse -Force -ErrorAction SilentlyContinue
}

Test-Refusal -Case 'missing-manifest' -Arrange {
    Remove-Item -LiteralPath (Join-Path $script:BuildDir 'stage_b_manifest.json') -Force
}
Test-Refusal -Case 'missing-stage-a-workbook' -Arrange {
    Remove-Item -LiteralPath (Join-Path $script:BuildDir 'PCCM_stageA.xlsx') -Force
}
Test-Refusal -Case 'missing-generated-vba' -Arrange {
    Remove-Item -LiteralPath (Join-Path $script:BuildDir 'vba') -Recurse -Force
}
# A STALE STAGE-B WORKBOOK IN THE SOURCE MUST NOT TRAVEL, and "the bundle built"
# is the wrong question to ask about it. The contract names the Stage-A workbook
# by name, so a stale output beside it is simply not copied - and what matters is
# that the bundle really does not contain one afterwards. The builder also refuses
# outright if one is present in the bundle before the bootstrap, which is the
# belt-and-braces half of the same property.
$BuildDir = New-FakeBuild
Set-Content -LiteralPath (Join-Path $BuildDir 'PCCM_stageB.xlsm') -Value 'stale-build' -NoNewline
$WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-fakework-' +
    [System.Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $WorkDir -Force
try {
    $stale = New-EquivalenceBundle -Mode 'Baseline' -Manifest $manifest -Stamp 'stale'
    $carried = Test-Path -LiteralPath (Join-Path ([string]$stale.Root) 'PCCM_stageB.xlsm')
    Write-Output ('STALE|' + $(if ($carried) { 'travelled' } else { 'absent' }) +
                  '|the bundle holds ' + [string]@($stale.Digests.Keys).Count + ' artifact(s)')
} catch {
    Write-Output ('STALE|refused|' + (($_.Exception.Message) -replace '\s+', ' '))
}
Remove-Item -LiteralPath $BuildDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue


# =============================================================================
# THE READINESS BARRIER, over a workbook whose answers are scripted. Each read
# the barrier makes is recorded; a scripted `nothing` answers null once. Sleeps
# are counted, not slept.
# =============================================================================
$script:Reads = New-Object System.Collections.ArrayList
$script:Slept = 0
$script:Answers = @{}
$script:ScriptedSheets = $null
function Start-Sleep { param([int]$Milliseconds, [int]$Seconds) $script:Slept = $script:Slept + $Milliseconds + (1000 * $Seconds) }
function Get-Scripted {
    param([string]$Member, $Default)
    $null = $script:Reads.Add($Member)
    if ($script:Answers.ContainsKey($Member) -and ($script:Answers[$Member].Count -gt 0)) {
        $next = $script:Answers[$Member][0]
        $script:Answers[$Member] = @($script:Answers[$Member] | Select-Object -Skip 1)
        if ($next -eq '<nothing>') { return $null }
        return $next
    }
    return $Default
}
function New-ScriptedWorkbook {
    param([string]$Path)
    $sheets = Microsoft.PowerShell.Utility\New-Object PSObject
    $sheets | Add-Member -MemberType ScriptMethod -Name Item -Value { param($Key) return (Get-Scripted -Member 'Item' -Default 'sheet') }
    $wb = Microsoft.PowerShell.Utility\New-Object PSObject
    $wb | Add-Member -MemberType ScriptProperty -Name FullName -Value ([scriptblock]::Create("return (Get-Scripted -Member 'FullName' -Default '" + $Path.Replace("'", "''") + "')"))
    $wb | Add-Member -MemberType ScriptProperty -Name Worksheets -Value { return (Get-Scripted -Member 'Worksheets' -Default $script:ScriptedSheets) }
    $script:ScriptedSheets = $sheets
    return $wb
}
$expected = Join-Path ([System.IO.Path]::GetTempPath()) 'PCCM_equiv_endpoints.xlsm'
foreach ($case in @(
        @{ Name = 'first-answer';           Answers = @{};                                          Path = $expected; Max = 12 },
        @{ Name = 'two-nothings-then-ready'; Answers = @{ FullName = @('<nothing>', '<nothing>') }; Path = $expected; Max = 12 },
        @{ Name = 'worksheet-nothing-once';  Answers = @{ Item = @('<nothing>') };                  Path = $expected; Max = 12 },
        @{ Name = 'wrong-workbook';          Answers = @{ FullName = @('C:\somewhere\else.xlsm') }; Path = $expected; Max = 12 },
        @{ Name = 'never-ready';             Answers = @{ Worksheets = @('<nothing>', '<nothing>', '<nothing>', '<nothing>') }; Path = $expected; Max = 3 })) {
    $script:Reads.Clear(); $script:Slept = 0
    $script:Answers = @{}
    foreach ($k in @($case.Answers.Keys)) { $script:Answers[$k] = @($case.Answers[$k]) }
    $wb = New-ScriptedWorkbook -Path ([string]$case.Path)
    try {
        $ready = Wait-EquivalenceWorkbookReady -Workbook $wb -ExpectedPath $expected -KnownSheet 'Cost Lines' -MaxAttempts ([int]$case.Max)
        Write-Output ('READY|' + $case.Name + '|ok|attempt=' + [string]$ready.Attempt + '|waited=' + [string]$ready.WaitedMs + '|slept=' + [string]$script:Slept)
    } catch {
        Write-Output ('READY|' + $case.Name + '|refused|' + (($_.Exception.Message) -replace '\s+', ' ') + '|slept=' + [string]$script:Slept)
    }
    Write-Output ('READYREADS|' + $case.Name + '|' + ((@($script:Reads) | Select-Object -Unique) -join ','))
}
Remove-Item -LiteralPath $BuildDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
exit 0
