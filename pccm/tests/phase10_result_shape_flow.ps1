<#
.SYNOPSIS
    PCCM test-only harness: THE PASS RESULT SHAPE THROUGH THE REAL COMPARISON.

.DESCRIPTION
    WHY THIS EXISTS. Equivalence run 12 completed both passes for the first time
    under the single-baseline architecture and then crashed at the comparison:
    "The property 'State' cannot be found on this object". Invoke-EquivalencePass
    had written a READY line to its own output stream ahead of its return record,
    so the caller held a two-element array where it expected the record.

    "What does the pass return", "does it emit anything else", and "does the
    comparison iterate every snapshot field and still reach CALCEQUIV" are
    BEHAVIOUR. So this reads the pass's return record and its emissions out of
    `tests/phase10_fixture_equivalence.ps1` BY AST, lifts the REAL
    Get-EquivalenceSnapshot and drives it over stand-ins to obtain the REAL field
    set, builds two result records the way the pass builds them, and cuts the
    gate's comparison block out of the file by its first and last statements and
    runs it over those records. Excel is never started.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_benchmark_harness.py` decides.

.NOTES
    Prints:
      RETURNS|<keys of the pass's return record, comma-separated>
      EMITS|<count of Write-Output / bare-expression emissions inside the pass>
      FIELDS|<count of snapshot fields>|<keys, comma-separated>
      COMPARE|<case>|equiv=<EQUIV lines>|match=<n>|differ=<n>|calc=<CALC lines>|calcequiv=<line>|calcequiv-after-last-equiv=<bool>|state-read=<bool>
    Exit 0 always.
#>
param([string]$Gate)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Gate)) { $Gate = Join-Path $here 'phase10_fixture_equivalence.ps1' }
$gatePath = (Resolve-Path -LiteralPath $Gate).Path

$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($gatePath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) { Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in the gate'); exit 0 }
$functions = @{}
foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
    $functions[$fn.Name] = $fn
}
foreach ($name in @('Invoke-EquivalencePass', 'Get-EquivalenceSnapshot')) {
    if (-not $functions.ContainsKey($name)) { Write-Output ('MISSING|' + $name + ' is not defined in the gate'); exit 0 }
}

# --- WHAT THE PASS RETURNS, AND WHAT ELSE IT EMITS ------------------------------
$pass = $functions['Invoke-EquivalencePass']
$returnKeys = @()
foreach ($ret in $pass.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.ReturnStatementAst] }, $true)) {
    foreach ($hash in $ret.FindAll({ param($n) $n -is [System.Management.Automation.Language.HashtableAst] }, $true)) {
        foreach ($pair in $hash.KeyValuePairs) { $returnKeys += [string]$pair.Item1.Value }
    }
}
Write-Output ('RETURNS|' + ($returnKeys -join ','))
$emits = 0
foreach ($cmd in $pass.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true)) {
    if (@('Write-Output', 'Write-Host', 'Out-Default', 'Out-String', 'Write-Information') -contains [string]$cmd.GetCommandName()) { $emits++ }
}
# A BARE VALUE AS A STATEMENT emits too: `$x` or `$x.Member` on its own line inside a
# statement block that is not an array, sub-expression or hashtable operand.
foreach ($pipe in $pass.Body.FindAll({ param($n) $n -is [System.Management.Automation.Language.PipelineAst] }, $true)) {
    if (-not ($pipe.Parent -is [System.Management.Automation.Language.StatementBlockAst])) { continue }
    $holder = $pipe.Parent.Parent
    if (($holder -is [System.Management.Automation.Language.ArrayExpressionAst]) -or
        ($holder -is [System.Management.Automation.Language.SubExpressionAst]) -or
        ($holder -is [System.Management.Automation.Language.HashtableAst]) -or
        ($holder -is [System.Management.Automation.Language.ParenExpressionAst])) { continue }
    if ($pipe.PipelineElements.Count -ne 1) { continue }
    $only = $pipe.PipelineElements[0]
    if (-not ($only -is [System.Management.Automation.Language.CommandExpressionAst])) { continue }
    $expr = $only.Expression
    if (($expr -is [System.Management.Automation.Language.VariableExpressionAst]) -or
        (($expr -is [System.Management.Automation.Language.MemberExpressionAst]) -and
         -not ($expr -is [System.Management.Automation.Language.InvokeMemberExpressionAst])) -or
        ($expr -is [System.Management.Automation.Language.StringConstantExpressionAst]) -or
        ($expr -is [System.Management.Automation.Language.BinaryExpressionAst])) { $emits++ }
}
Write-Output ('EMITS|' + [string]$emits)

# --- THE REAL SNAPSHOT, over stand-ins, for the REAL field set ------------------
Invoke-Expression $functions['Get-EquivalenceSnapshot'].Extent.Text
function Get-IdColumnValues { param($Workbook, $Info) return @('CL-001') }
function Get-NamedValue { param($Workbook, [string]$DefinedName) return ('v:' + $DefinedName) }
function Get-TableColumnNames { param($Workbook, [string]$SheetName, [string]$TableName) return @('a', 'b') }
function Get-TableBody { param($Workbook, [string]$SheetName, [string]$TableName) return @(,@('x', 'y')) }
$fakeExcel = New-Object PSObject
$fakeExcel | Add-Member -MemberType ScriptMethod -Name Run -Value { return ('r:' + [string]$args[0]) }
$manifest = [pscustomobject]@{
    registers = @([pscustomobject]@{ key = 'cost_lines'; sheet = 'S'; table_name = 'T1' },
                  [pscustomobject]@{ key = 'risk_register'; sheet = 'S'; table_name = 'T2' })
    counters  = @([pscustomobject]@{ key = 'cost_line'; defined_name = 'nmA' }, [pscustomobject]@{ key = 'risk'; defined_name = 'nmB' })
    grids     = @([pscustomobject]@{ key = 'cost_profiling'; sheet = 'S'; table_name = 'G1' },
                  [pscustomobject]@{ key = 'risk_profiling'; sheet = 'S'; table_name = 'G2' },
                  [pscustomobject]@{ key = 'inflation'; sheet = 'S'; table_name = 'G3' })
}
$inspection = [pscustomobject]@{ input_tables = [pscustomobject]@{
    fx_rates = [pscustomobject]@{ sheet = 'S'; table_name = 'FX' }
    inflation_profiles = [pscustomobject]@{ sheet = 'S'; table_name = 'IP' } } }
function New-FakeState {
    return (Get-EquivalenceSnapshot -Excel $fakeExcel -Workbook 'wb' -Manifest $manifest -Inspection $inspection -SimInspection $null)
}
$fields = New-FakeState
Write-Output ('FIELDS|' + [string]@($fields.Keys).Count + '|' + (@($fields.Keys) -join ','))

# --- THE COMPARISON BLOCK, LIFTED VERBATIM ---------------------------------------
$source = Get-Content -LiteralPath $gatePath -Raw
$startMark = '$completed = @($passes.Keys)'
$endMark = 'if ($null -ne $bundle) {'
$from = $source.IndexOf($startMark); $to = $source.IndexOf($endMark)
if (($from -lt 0) -or ($to -lt $from)) { Write-Output 'MISSING|the comparison anchors are not in the gate'; exit 0 }
$comparison = $source.Substring($from, $to - $from)

function New-PassRecord {
    param([string]$Mode, $State, [string]$Fingerprint)
    # BUILT WITH THE KEYS THE PASS RETURNS, read from its own return statement.
    $record = @{}
    foreach ($key in $returnKeys) { $record[$key] = $null }
    $record['Mode'] = $Mode; $record['State'] = $State
    $record['CalcResult'] = 'OK|Calculation committed.'; $record['CalcStatus'] = 'CURRENT'
    $record['CalcFingerprint'] = $Fingerprint
    if ($record.ContainsKey('Ready')) { $record['Ready'] = [pscustomobject]@{ Attempt = 1; WaitedMs = 0 } }
    if ($record.ContainsKey('Shutdown')) { $record['Shutdown'] = @() }
    return [pscustomobject]$record
}
foreach ($case in @('identical', 'one-field-and-fingerprint-differ')) {
    $left = New-FakeState
    $right = New-FakeState
    $rightPrint = 'ABC'
    if ($case -ne 'identical') { $right['cost_profiling.body'] = 'changed'; $rightPrint = 'XYZ' }
    $passes = @{}
    $passes['Endpoints'] = New-PassRecord -Mode 'Endpoints' -State $left -Fingerprint 'ABC'
    $passes['Bulk'] = New-PassRecord -Mode 'Bulk' -State $right -Fingerprint $rightPrint
    $setupFailed = $false
    $lines = @()
    $stateRead = $true
    try { $lines = @(Invoke-Expression $comparison) } catch { $stateRead = $false; $lines = @('RAISED|' + $_.Exception.Message) }
    $equiv = @($lines | Where-Object { $_ -like 'EQUIV|*' })
    $matches = @($equiv | Where-Object { $_ -like 'EQUIV|*|match|*' }).Count
    $differs = @($equiv | Where-Object { $_ -like 'EQUIV|*|differ|*' }).Count
    $calc = @($lines | Where-Object { $_ -like 'CALC|*' }).Count
    $calcequiv = @($lines | Where-Object { $_ -like 'CALCEQUIV|*' })
    $lastEquivAt = -1; $calcequivAt = -1
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -like 'EQUIV|*') { $lastEquivAt = $i }
        if (($lines[$i] -like 'CALCEQUIV|*') -and ($calcequivAt -lt 0)) { $calcequivAt = $i }
    }
    Write-Output ('COMPARE|' + $case + '|equiv=' + [string]$equiv.Count + '|match=' + [string]$matches +
                  '|differ=' + [string]$differs + '|calc=' + [string]$calc + '|calcequiv=' +
                  $(if ($calcequiv.Count -eq 1) { [string]$calcequiv[0] } else { '<' + [string]$calcequiv.Count + ' lines>' }) +
                  '|calcequiv-after-last-equiv=' + [string](($calcequivAt -gt $lastEquivAt) -and ($lastEquivAt -ge 0)) +
                  '|state-read=' + [string]$stateRead)
}
exit 0
