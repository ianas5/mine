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

    So this RUNS it. It lifts `Get-BundleArtifacts`, `New-EquivalenceBundle` and
    `Test-BundleIdentity` out of `tests/phase10_fixture_equivalence.ps1` BY AST -
    their real bytes, not a copy - and drives them against a FAKE repository build
    directory holding exactly the artifacts the contract names. Excel is never
    started and no workbook is opened: this is about which files arrive where.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      ARTIFACT|<kind>|<name>              one entry of the derived contract
      BUNDLE|<mode>|<relative path>       one file that arrived, per bundle
      ROOTS|<isolated|shared>             whether the two bundles are separate
      IDENTITY|<identical|differ>|<why>   the starting-state comparison
      DAMAGED|<differ|identical>|<why>    the same comparison after one edit
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
$wanted = @('Get-BundleArtifacts', 'New-EquivalenceBundle', 'Test-BundleIdentity')
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

$left = $null; $right = $null
try {
    $left = New-EquivalenceBundle -Mode 'Endpoints' -Manifest $manifest -Stamp 'flow'
    $right = New-EquivalenceBundle -Mode 'Bulk' -Manifest $manifest -Stamp 'flow'
} catch {
    Write-Output ('REFUSE|two-bundles|refused|' + (($_.Exception.Message) -replace '\s+', ' '))
}

if (($null -ne $left) -and ($null -ne $right)) {
    foreach ($key in @($left.Digests.Keys)) { Write-Output ('BUNDLE|Endpoints|' + $key) }
    foreach ($key in @($right.Digests.Keys)) { Write-Output ('BUNDLE|Bulk|' + $key) }
    Write-Output ('ROOTS|' + $(if ([string]$left.Root -ne [string]$right.Root) { 'isolated' }
                               else { 'shared' }) + '|' + [string]$left.Root + ' :: ' +
                  [string]$right.Root)
    $problems = @(Test-BundleIdentity -Left $left -Right $right)
    Write-Output ('IDENTITY|' + $(if ($problems.Count -eq 0) { 'identical' } else { 'differ' }) +
                  '|' + ($problems -join '; '))

    # ONE ARTIFACT EDITED, AND THE COMPARISON MUST NOTICE. Without this the
    # identity check could be vacuous and nobody would know.
    $victim = Join-Path ([string]$right.Root) 'stage_b_manifest.json'
    Set-Content -LiteralPath $victim -Value '{"fake":"tampered"}' -NoNewline
    $right.Digests['stage_b_manifest.json'] =
        [string](Get-FileHash -LiteralPath $victim -Algorithm SHA256).Hash
    $damaged = @(Test-BundleIdentity -Left $left -Right $right)
    Write-Output ('DAMAGED|' + $(if ($damaged.Count -eq 0) { 'identical' } else { 'differ' }) +
                  '|' + ($damaged -join '; '))
}

# --- THE REFUSALS, EACH BEFORE EXCEL COULD HAVE STARTED ---------------------
function Test-Refusal {
    param([string]$Case, [scriptblock]$Arrange)
    $script:BuildDir = New-FakeBuild
    $script:WorkDir = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm-fakework-' +
        [System.Guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:WorkDir -Force
    & $Arrange
    try {
        $null = New-EquivalenceBundle -Mode 'Endpoints' -Manifest $manifest -Stamp $Case
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
    $stale = New-EquivalenceBundle -Mode 'Endpoints' -Manifest $manifest -Stamp 'stale'
    $carried = Test-Path -LiteralPath (Join-Path ([string]$stale.Root) 'PCCM_stageB.xlsm')
    Write-Output ('STALE|' + $(if ($carried) { 'travelled' } else { 'absent' }) +
                  '|the bundle holds ' + [string]@($stale.Digests.Keys).Count + ' artifact(s)')
} catch {
    Write-Output ('STALE|refused|' + (($_.Exception.Message) -replace '\s+', ' '))
}
Remove-Item -LiteralPath $BuildDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue

Remove-Item -LiteralPath $BuildDir -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
exit 0
