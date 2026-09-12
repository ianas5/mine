<#
.SYNOPSIS
    PCCM test-only harness: WHICH STAGE-B BUILD COM CALL WAS REFUSED, executed.

.DESCRIPTION
    WHY THIS EXISTS. Two consecutive Windows equivalence runs died in the Stage-B
    build block with RPC_E_CALL_REJECTED, and the transcript could name only the
    region: one try/catch around eleven distinct COM operations reported itself as
    `[FAIL] Stage-B build`. The exact rejected call is STILL UNKNOWN, and only a
    Windows run can name it.

    What CAN be settled here is whether the instrumentation that will name it
    works. "Which label was in flight", "how many attempts did that read take",
    "was the original error preserved", "is this line a COMREJECT or a COMFAIL"
    are BEHAVIOUR, not properties of text. So this lifts the five new functions
    out of `bootstrap/windows/build_stage_b.ps1` BY AST, dot-sources the accepted
    `com_lifecycle.ps1` beside them, and drives them against fakes that raise real
    COMExceptions carrying real HRESULTs.

    Excel is never started, no workbook is opened, and no COM object exists.

    WHY THE RETRY LOOP IS DRIVEN THROUGH A METHOD. This host SWALLOWS an
    exception thrown by a .NET property getter: `$obj.$member` yields $null and
    does not even reach $Error. A method call raises a
    MethodInvocationException whose InnerException is the COMException - which is
    exactly the chain the accepted classifier walks, and why it walks one. So the
    real helper's loop is exercised through `Item`, and the wrapper's own logic is
    exercised separately against a stubbed helper. Neither instrument is asked to
    prove the other's half. The HOST line records this behaviour, because it is
    a fact about the harness and NOT a claim about Excel: COM objects use a
    different PowerShell adapter, and the rejection actually observed on Windows
    arrived as a catchable COMException.

    IT ASSERTS NOTHING. It prints tagged lines and
    `tests/test_phase10_stage_b_verification_retry.py` decides.

.NOTES
    Prints:
      OPS|<n>|<comma-separated vocabulary>
      LABEL|<case>|<asked for>|<in effect>|<outcome>
      HRES|<case>|<extracted>
      LINE|<case>|<the diagnostic line>
      RETRY|<case>|<attempts>|<waited>|<calls>|<outcome>|<type>|<hresult>
      LEDGER|<case>|<ledger line>
      WRAP|<case>|<label at forward>|<attempts>|<lines>|<outcome>|<line>
      HOST|<property-get outcome>|<method-call outcome>
      SAVE|<case>|<SaveAs calls>|<attempts>|<waited>|<format>|<outcome>|<state>
      SAVENOTE|<case>|<diagnostic line>
      SAVESTATE|<case>|<state>|<detail>
      SAVEPATH|<case>|<are the two paths the same>
      READ|<case>|<outcome>|<value or refusal>|<CLR type>|lines=<n>
      TELEMETRY|<case>|...
      SUBOPS|<n>|<comma-separated sub-operation vocabulary>
      STEP|<case>|<label>|<outcome>
      READY|<case>|<outcome>|<attempts>|<waited>|<format>|<reads>|<sleeps>|<detail>
      READYNOTE|<case>|<diagnostic line>
    Exit 0 always.
#>
param(
    [string]$Build,
    [string]$Lifecycle
)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$windows = Join-Path (Split-Path -Parent $here) 'bootstrap/windows'
if ([string]::IsNullOrWhiteSpace($Build))     { $Build     = Join-Path $windows 'build_stage_b.ps1' }
if ([string]::IsNullOrWhiteSpace($Lifecycle)) { $Lifecycle = Join-Path $windows 'com_lifecycle.ps1' }

# The accepted retry helper and its classifier come from the REAL file. No part of
# the retry policy is restated here; the wrapper under test forwards to it.
. (Resolve-Path -LiteralPath $Lifecycle).Path

$buildPath = (Resolve-Path -LiteralPath $Build).Path
$errors = $null; $tokens = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($buildPath, [ref]$tokens, [ref]$errors)
if ($errors -and $errors.Count -gt 0) {
    Write-Output ('PARSE|' + [string]$errors.Count + ' parse error(s) in build_stage_b.ps1')
    exit 0
}

# DECLARED BEFORE IT IS LIFTED. The vocabulary arrives through Invoke-Expression,
# which the uninitialised-scope audit cannot see, so it is given an empty value
# first: a lift that silently produced nothing then prints OPS|0 and fails a
# control rather than reading an undefined script variable.
$script:StageBBuildOps    = @()
$script:StageBBuildSubOps = @()

# Section R's scripted stand-in for the accepted helper records what it was asked
# for. Declared HERE rather than beside the stub, because a script-scope variable
# first assigned after the first helper call is the defect the uninitialised-scope
# audit exists to catch.
$script:ReadyStubScript = @()
$script:ReadyStubIndex  = 0
$script:ReadyStubReads  = 0
$script:ReadyStubPath   = ''

# Section D's stub records what the wrapper handed it. Declared HERE rather than
# beside the stub, because a script-scope variable first assigned after the first
# helper call is the defect the uninitialised-scope audit exists to catch.
$script:StubLabel = ''
$script:StubArgs  = ''

# The vocabulary is an assignment at the top of the script, not a function, so it
# is lifted by its own statement rather than by function name.
$vocabText = ''
$subVocabText = ''
foreach ($assign in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)) {
    if ($assign.Left.Extent.Text -eq '$script:StageBBuildOps')    { $vocabText = $assign.Extent.Text }
    if ($assign.Left.Extent.Text -eq '$script:StageBBuildSubOps') { $subVocabText = $assign.Extent.Text }
}
if ($vocabText -eq '' -or $subVocabText -eq '') {
    Write-Output 'MISSING|the build operation vocabularies are not defined in build_stage_b.ps1'
    exit 0
}
Invoke-Expression $vocabText
Invoke-Expression $subVocabText
$script:StageBBuildOp         = '<before the first labelled operation>'
$script:StageBBuildSubOp      = ''
$script:StageBBuildRejections = New-Object System.Collections.ArrayList

# Add-Note is lifted too, because the SaveAs settlement REPORTS through it and
# the diagnostic lines are part of what has to be observed.
$notes = New-Object System.Collections.ArrayList

foreach ($name in @('Add-Note', 'Set-StageBBuildOp', 'Get-StageBBuildOp',
                    'Set-StageBBuildStep', 'Get-StageBBuildStep', 'Get-StageBBuildLabel',
                    'Get-StageBBuildRejections', 'Add-StageBReadRejection',
                    'Get-StageBScalarInt', 'Get-StageBNonEmptyString',
                    'Get-StageBComHResult', 'New-StageBRejectionLine',
                    'Get-StageBComparablePath', 'Wait-StageBWorkbookReady',
                    'Get-StageBSaveAsPostcondition',
                    'Invoke-StageBSaveAs', 'New-StageBSaveAsResult')) {
    $body = $null
    foreach ($fn in $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) {
        if ($fn.Name -eq $name) { $body = $fn.Extent.Text }
    }
    if ($null -eq $body) {
        Write-Output ('MISSING|' + $name + ' is not defined in build_stage_b.ps1')
        exit 0
    }
    Invoke-Expression $body
}

Write-Output ('OPS|' + [string]@($script:StageBBuildOps).Count + '|' +
              (@($script:StageBBuildOps) -join ','))

# ---------------------------------------------------------------------------
# A. THE CLOSED VOCABULARY
# ---------------------------------------------------------------------------
foreach ($label in @($script:StageBBuildOps)) {
    $outcome = 'accepted'
    try { Set-StageBBuildOp $label } catch { $outcome = 'REFUSED' }
    Write-Output ('LABEL|' + $label + '|' + $label + '|' + (Get-StageBBuildOp) + '|' + $outcome)
}
# A label that is not on the list must not become the label in effect: a
# misspelling has to fail where it is written, not inside a diagnostic.
$before = Get-StageBBuildOp
$outcome = 'accepted'
try { Set-StageBBuildOp 'vbproject.aquire' } catch { $outcome = 'REFUSED' }
Write-Output ('LABEL|typo|vbproject.aquire|' + (Get-StageBBuildOp) + '|' + $outcome)
Write-Output ('LABEL|typo-kept-previous|' + $before + '|' + (Get-StageBBuildOp) + '|unchanged')

# ---------------------------------------------------------------------------
# B. THE HRESULT, AND WHAT IT IS NOT
# ---------------------------------------------------------------------------
function New-ComError {
    param([string]$Message, [int]$Code)
    return (New-Object System.Management.Automation.ErrorRecord `
        (New-Object System.Runtime.InteropServices.COMException $Message, $Code), 'id', 'NotSpecified', $null)
}
function New-PlainError {
    return (New-Object System.Management.Automation.ErrorRecord `
        (New-Object System.InvalidOperationException 'not a COM failure at all'), 'id', 'NotSpecified', $null)
}
function New-WrappedError {
    param([int]$Code)
    $inner = New-Object System.Runtime.InteropServices.COMException 'refused', $Code
    return (New-Object System.Management.Automation.ErrorRecord `
        (New-Object System.InvalidOperationException 'wrapper', $inner), 'id', 'NotSpecified', $null)
}

$cases = @(
    @{ case = 'rejected';   record = (New-ComError 'Call was rejected by callee.' -2147418111) },
    @{ case = 'retrylater'; record = (New-ComError 'Server call retry later.' -2147417846) },
    @{ case = 'accepted';   record = (New-ComError 'Exception from HRESULT: 0x800A03EC' -2146827284) },
    @{ case = 'vba-ignore'; record = (New-ComError 'Ignore.' -2146777998) },
    @{ case = 'not-com';    record = (New-PlainError) },
    @{ case = 'wrapped';    record = (New-WrappedError -2147418111) }
)
foreach ($c in $cases) {
    $hex = Get-StageBComHResult $c.record
    if ($hex -eq '') { $hex = '<empty>' }
    Write-Output ('HRES|' + $c.case + '|' + $hex)
    Write-Output ('LINE|' + $c.case + '|' +
                  (New-StageBRejectionLine -Operation 'vbproject.acquire' -ErrorRecord $c.record))
}
$nullHex = Get-StageBComHResult $null
if ($nullHex -eq '') { $nullHex = '<empty>' }
Write-Output ('HRES|null|' + $nullHex)

# ---------------------------------------------------------------------------
# C. THE ACCEPTED HELPER'S LOOP, DRIVEN THROUGH A METHOD
# ---------------------------------------------------------------------------
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class PccmFakeComTarget {
    public static int Calls = 0;
    public static int Fails = 0;
    public static int Code  = -2147418111;
    public static bool PlainError = false;
    private static void Maybe() {
        Calls++;
        if (Calls > Fails) return;
        if (PlainError) throw new InvalidOperationException("not a COM failure at all");
        throw new COMException("Call was rejected by callee.", Code);
    }
    public object Item(object key) { Maybe(); return "the item"; }
    public object Quiet { get { Maybe(); return "answered"; } }
    // Value shapes a real read can answer with. None of these throws: what is
    // being observed here is whether the VALUE survives the read path, which is
    // what run 5 proved it did not.
    public object Fmt51     { get { return 51; } }
    public object Fmt52     { get { return 52; } }
    public object FmtText   { get { return "52"; } }
    public object FullName  { get { return "/b/PCCM_stageA.xlsx"; } }
    public object Blank     { get { return "   "; } }
    public object Nothing   { get { return null; } }
    public object Pair      { get { return new object[] { 51, 52 }; } }
    public object NotNumber { get { return "xlsm"; } }
    public object Child     { get { return new PccmFakeComTarget(); } }
}
'@
$fake = New-Object PccmFakeComTarget

# HOST BEHAVIOUR, ON THE RECORD. The property form is why section D exists.
[PccmFakeComTarget]::Calls = 0; [PccmFakeComTarget]::Fails = 99
[PccmFakeComTarget]::Code = -2147418111; [PccmFakeComTarget]::PlainError = $false
$propOutcome = 'swallowed'
$member = 'Quiet'
try { $null = $fake.$member } catch { $propOutcome = 'raised ' + $_.Exception.GetType().Name }
$methodOutcome = 'swallowed'
try { $null = $fake.Item(1) } catch { $methodOutcome = 'raised ' + $_.Exception.GetType().Name }
Write-Output ('HOST|' + $propOutcome + '|' + $methodOutcome)

foreach ($c in @(
    @{ case = 'answers-first-call'; fails = 0;  code = -2147418111; plain = $false; attempts = 12; budget = 15000; delay = 1 },
    @{ case = 'refused-twice';      fails = 2;  code = -2147418111; plain = $false; attempts = 12; budget = 15000; delay = 1 },
    @{ case = 'retrylater-once';    fails = 1;  code = -2147417846; plain = $false; attempts = 12; budget = 15000; delay = 1 },
    @{ case = 'always-refused';     fails = 99; code = -2147418111; plain = $false; attempts = 3;  budget = 10;    delay = 1 },
    @{ case = 'budget-bound';       fails = 99; code = -2147418111; plain = $false; attempts = 50; budget = 5;     delay = 2 },
    @{ case = 'accepted-and-failed';fails = 99; code = -2146827284; plain = $false; attempts = 12; budget = 15000; delay = 1 },
    @{ case = 'vba-ignore';         fails = 99; code = -2146777998; plain = $false; attempts = 12; budget = 15000; delay = 1 },
    @{ case = 'not-com';            fails = 99; code = 0;           plain = $true;  attempts = 12; budget = 15000; delay = 1 })) {
    [PccmFakeComTarget]::Calls = 0
    [PccmFakeComTarget]::Fails = [int]$c.fails
    [PccmFakeComTarget]::Code  = [int]$c.code
    [PccmFakeComTarget]::PlainError = [bool]$c.plain
    $ledgerBase = @(Get-ComRetryLedger).Count
    $outcome = 'answered'; $attempts = -1; $waited = -1; $type = '<none>'; $hex = '<none>'
    try {
        $record = Invoke-ComRetryRead -Target $fake -Member 'Item' -Key 1 `
            -Description ('fake Item for ' + [string]$c.case) `
            -MaxAttempts ([int]$c.attempts) -FirstDelayMs ([int]$c.delay) `
            -MaxDelayMs ([int]$c.delay) -TotalBudgetMs ([int]$c.budget)
        $attempts = [int]$record.Attempts
        $waited   = [int]$record.WaitedMs
        if ([string]$record.Value -ne 'the item') { $outcome = 'WRONG-VALUE' }
    } catch {
        $outcome = 'RAISED'
        $type = $_.Exception.GetType().FullName
        $hex = Get-StageBComHResult $_
        if ($hex -eq '') { $hex = '<empty>' }
    }
    Write-Output ('RETRY|' + $c.case + '|' + [string]$attempts + '|' + [string]$waited + '|' +
                  [string][PccmFakeComTarget]::Calls + '|' + $outcome + '|' + $type + '|' + $hex)
    foreach ($line in @(@(Get-ComRetryLedger) | Select-Object -Skip $ledgerBase)) {
        Write-Output ('LEDGER|' + $c.case + '|' + $line)
    }
}

# ---------------------------------------------------------------------------
# CS. SaveAs: SETTLED BY ITS POSTCONDITION
# ---------------------------------------------------------------------------
# WINDOWS NAMED saveas.xlsm AS THE REJECTED CALL, and SaveAs writes a file and
# rebinds the workbook. 'The message filter says the call never ran' is a contract,
# not an observation, so what is driven here is the OBSERVATION: three independent
# facts, and every combination of them that is not all-source or all-target.
#
# The fake's SaveAs is a METHOD, so its exception really propagates, and it can
# apply each partial side effect a real half-completed save could leave behind.
Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Runtime.InteropServices;
public class PccmFakeWorkbook {
    // BOOLEANS, NOT NULLABLE STRINGS. PowerShell coerces $null to "" when it
    // assigns to a string field, so a `!= null` guard here fired on an empty path
    // and File.WriteAllText raised ArgumentException instead of the COMException
    // the case was testing - a fake that failed for the wrong reason.
    public static int    Calls      = 0;
    public static int    Fails      = 0;
    public static int    Code       = -2147418111;
    public static bool   PlainError = false;
    public static string Bound      = "";
    public static string Target     = "";
    public static int    Format     = 51;
    public static bool   Rebind     = false;
    public static int    SetFormat  = -1;
    public static bool   MakeFile   = false;
    public static bool   NullName   = false;
    // A SaveAs that RETURNS and does nothing. Excel returning from a method is not
    // the same as the save having happened, and nothing else here tests that.
    public static bool   Silent     = false;
    public object FullName   { get { if (NullName) return null; return Bound; } }
    public object FileFormat { get { return Format; } }
    public void SaveAs(object path, object format) {
        Calls++;
        if (Calls <= Fails) {
            // Whatever partial state this case says a half-done save would leave.
            if (Rebind)        Bound  = Target;
            if (SetFormat >= 0) Format = SetFormat;
            if (MakeFile)      File.WriteAllText(Target, "partial");
            if (PlainError) throw new InvalidOperationException("Excel accepted this and it failed");
            throw new COMException("Call was rejected by callee.", Code);
        }
        if (Silent) return;
        Bound  = (string)path;
        Format = Convert.ToInt32(format);
        File.WriteAllText((string)path, "stage b");
    }
}
'@

$saveRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm_saveas_' + [System.Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $saveRoot -Force
$srcPath = Join-Path $saveRoot 'PCCM_stageA.xlsx'
$tgtPath = Join-Path $saveRoot 'PCCM_stageB.xlsm'
Set-Content -LiteralPath $srcPath -Value 'stage a' -NoNewline

$saveCases = @(
    # name                fails code          plain    rebind   setfmt file     nullname att delay budget
    @{ n='clean';         f=0;  c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000; s=$false },
    @{ n='silent-success'; f=0; c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000; s=$true },
    @{ n='not-executed';  f=1;  c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='retrylater';    f=1;  c=-2147417846; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='completed';     f=1;  c=-2147418111; p=$false; r=$true;  sf=52; cf=$true;  nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-file-only'; f=99; c=-2147418111; p=$false; r=$false; sf=-1; cf=$true;  nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-path-only'; f=99; c=-2147418111; p=$false; r=$true;  sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-fmt-only';  f=99; c=-2147418111; p=$false; r=$false; sf=52; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-path-file'; f=99; c=-2147418111; p=$false; r=$true;  sf=-1; cf=$true;  nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-path-fmt';  f=99; c=-2147418111; p=$false; r=$true;  sf=52; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='amb-null-name'; f=99; c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$true;  a=12; d=1; b=15000 ; s=$false },
    @{ n='accepted-err';  f=99; c=-2146827284; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='not-com-err';   f=99; c=0;           p=$true;  r=$false; sf=-1; cf=$false; nn=$false; a=12; d=1; b=15000 ; s=$false },
    @{ n='attempt-bound'; f=99; c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=3;  d=1; b=10000 ; s=$false },
    @{ n='budget-bound';  f=99; c=-2147418111; p=$false; r=$false; sf=-1; cf=$false; nn=$false; a=50; d=2; b=5 ; s=$false }
)
foreach ($c in $saveCases) {
    if (Test-Path -LiteralPath $tgtPath) { Remove-Item -LiteralPath $tgtPath -Force }
    $notes.Clear()
    $null = Set-StageBBuildOp 'saveas.xlsm'
    [PccmFakeWorkbook]::Calls      = 0
    [PccmFakeWorkbook]::Fails      = [int]$c.f
    [PccmFakeWorkbook]::Code       = [int]$c.c
    [PccmFakeWorkbook]::PlainError = [bool]$c.p
    [PccmFakeWorkbook]::Bound      = $srcPath
    [PccmFakeWorkbook]::Format     = 51
    [PccmFakeWorkbook]::NullName   = [bool]$c.nn
    [PccmFakeWorkbook]::Target     = $tgtPath
    [PccmFakeWorkbook]::Rebind     = [bool]$c.r
    [PccmFakeWorkbook]::SetFormat  = [int]$c.sf
    [PccmFakeWorkbook]::MakeFile   = [bool]$c.cf
    [PccmFakeWorkbook]::Silent     = [bool]$c.s
    $fake = New-Object PccmFakeWorkbook

    $outcome = 'returned'; $state = '<none>'; $attempts = -1; $waited = -1; $format = -1
    try {
        $result = Invoke-StageBSaveAs -Workbook $fake -SourcePath $srcPath -TargetPath $tgtPath `
            -TargetFormat 52 -SourceFormat 51 `
            -MaxAttempts ([int]$c.a) -FirstDelayMs ([int]$c.d) -MaxDelayMs ([int]$c.d) `
            -TotalBudgetMs ([int]$c.b)
        $state = [string]$result.State
        $attempts = [int]$result.Attempts
        $waited = [int]$result.WaitedMs
        $format = [int]$result.FileFormat
    } catch {
        $outcome = 'RAISED'
        $state = (Get-StageBComHResult $_)
        if ($state -eq '') { $state = 'no-hresult' }
    }
    Write-Output ('SAVE|' + $c.n + '|' + [string][PccmFakeWorkbook]::Calls + '|' +
                  [string]$attempts + '|' + [string]$waited + '|' + [string]$format + '|' +
                  $outcome + '|' + $state)
    foreach ($line in @($notes)) { Write-Output ('SAVENOTE|' + $c.n + '|' + $line) }
}

# A postcondition inspection that cannot read anything is AMBIGUOUS, not evidence.
$unreadable = Get-StageBSaveAsPostcondition -Workbook $null -SourcePath $srcPath `
    -TargetPath $tgtPath -TargetFormat 52 -SourceFormat 51
Write-Output ('SAVESTATE|unreadable|' + $unreadable.State + '|' + $unreadable.Detail)

# Separator and case differences are not a rebind.
Write-Output ('SAVEPATH|same|' + [string]((Get-StageBComparablePath $tgtPath) -eq
              (Get-StageBComparablePath ($tgtPath.Replace([System.IO.Path]::DirectorySeparatorChar, '/')))))
Write-Output ('SAVEPATH|different|' + [string]((Get-StageBComparablePath $tgtPath) -eq
              (Get-StageBComparablePath $srcPath)))

Remove-Item -LiteralPath $saveRoot -Recurse -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------------
# D. THE READ PATH THE BUILD NOW USES, AND THE VALUE SURVIVING IT
# ---------------------------------------------------------------------------
# RUN 5's FAILURE WAS A VALUE THAT DID NOT ARRIVE. So this drives the form the
# build now uses - the accepted helper called DIRECTLY - and checks that the
# answer comes back intact for every shape a real read produces, that the
# out-of-band telemetry can neither pollute nor consume it, and that the
# validators refuse the answers that are not answers.
#
# ITS OWN TARGET. $fake was rebound to the workbook fake by the SaveAs section
# above, so reusing it here read FullName off the wrong type while every other
# property came back "cannot be found" - a section that tested nothing it claimed.
$reader = New-Object PccmFakeComTarget
foreach ($c in @(
    @{ n='int-51';    m='Fmt51';     kind='int' },
    @{ n='int-52';     m='Fmt52';     kind='int' },
    @{ n='int-as-text'; m='FmtText';  kind='int' },
    @{ n='fullname';   m='FullName';  kind='string' },
    @{ n='object';     m='Child';     kind='object' },
    @{ n='nothing';    m='Nothing';   kind='int' },
    @{ n='blank';      m='Blank';     kind='string' },
    @{ n='two-values'; m='Pair';      kind='int' },
    @{ n='not-number'; m='NotNumber'; kind='int' })) {
    [PccmFakeComTarget]::Calls = 0
    [PccmFakeComTarget]::Fails = 0
    $null = Set-StageBBuildOp 'saveas.xlsm'
    Set-StageBBuildStep 'saveas.presave.fileformat'
    $script:StageBBuildRejections.Clear()
    $outcome = 'ok'; $shown = '<none>'; $type = '<none>'
    try {
        $read = Invoke-ComRetryRead -Target $reader -Member ([string]$c.m) `
                    -Description ('the fake ' + [string]$c.m)
        Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $read
        if ($c.kind -eq 'int') {
            $value = Get-StageBScalarInt -Value $read.Value -What ('the fake ' + [string]$c.m)
            $shown = [string]$value
            $type = $value.GetType().Name
        } elseif ($c.kind -eq 'string') {
            $value = Get-StageBNonEmptyString -Value $read.Value -What ('the fake ' + [string]$c.m)
            $shown = [string]$value
            $type = $value.GetType().Name
        } else {
            if ($null -eq $read.Value) { throw 'the object read answered with nothing' }
            $value = $read.Value
            $shown = 'an object'
            $type = $value.GetType().Name
        }
    } catch {
        $outcome = 'REFUSED'
        $shown = [string]$_.Exception.Message
    }
    Write-Output ('READ|' + $c.n + '|' + $outcome + '|' + $shown + '|' + $type + '|lines=' +
                  [string]@(Get-StageBBuildRejections).Count)
}

# THE TELEMETRY RECORDS AND RETURNS NOTHING. Assigning from it must yield $null,
# and calling it must leave the value it was told about untouched.
[PccmFakeComTarget]::Calls = 0; [PccmFakeComTarget]::Fails = 0
$read = Invoke-ComRetryRead -Target $reader -Member 'Fmt51' -Description 'the fake Fmt51'
$before = [string]$read.Value
$probe = Add-StageBReadRejection -Operation 'saveas.presave.fileformat' -Record $read
$after = [string]$read.Value
Write-Output ('TELEMETRY|emits|' + $(if ($null -eq $probe) { 'nothing' } else { 'SOMETHING: ' + [string]$probe }) +
              '|value-before=' + $before + '|value-after=' + $after +
              '|lines=' + [string]@(Get-StageBBuildRejections).Count)

# THE SAME PROBE ON THE BRANCH THAT ACTUALLY RECORDS. The quiet path returns early,
# so a probe that only exercises it never reaches the statement that could emit.
$script:StageBBuildRejections.Clear()
$fabricated = [pscustomobject]@{ Attempts = 3; WaitedMs = 750
                                 Rejections = 'RPC_E_CALL_REJECTED (0x80010001)' }
$probe2 = Add-StageBReadRejection -Operation 'saveas.presave.fileformat' -Record $fabricated
Write-Output ('TELEMETRY|emits-when-recording|' +
              $(if ($null -eq $probe2) { 'nothing' } else { 'SOMETHING: ' + [string]$probe2 }) +
              '|lines=' + [string]@(Get-StageBBuildRejections).Count)

# AND IT DOES RECORD when a read really was reissued - through the METHOD path,
# because that is the one whose exception this host propagates.
$script:StageBBuildRejections.Clear()
[PccmFakeComTarget]::Calls = 0
[PccmFakeComTarget]::Fails = 2
[PccmFakeComTarget]::Code = -2147418111
[PccmFakeComTarget]::PlainError = $false
$null = Set-StageBBuildOp 'saveas.xlsm'
Set-StageBBuildStep 'saveas.presave.fileformat'
$outcome = 'ok'; $shown = '<none>'
try {
    $read = Invoke-ComRetryRead -Target $reader -Member 'Item' -Key 1 `
                -Description 'the fake Item' -MaxAttempts 12 -FirstDelayMs 1 `
                -MaxDelayMs 1 -TotalBudgetMs 15000
    Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $read
    $shown = [string]$read.Value + '|attempts=' + [string]$read.Attempts
} catch { $outcome = 'REFUSED'; $shown = [string]$_.Exception.Message }
Write-Output ('TELEMETRY|retry-then-value|' + $outcome + '|' + $shown + '|' +
              ((@(Get-StageBBuildRejections)) -join ' ;; '))

# AN EXHAUSTED READ STILL THROWS, and the telemetry never sees a value to record.
$script:StageBBuildRejections.Clear()
[PccmFakeComTarget]::Calls = 0
[PccmFakeComTarget]::Fails = 99
$outcome = 'ok'; $hex = '<none>'
try {
    $read = Invoke-ComRetryRead -Target $reader -Member 'Item' -Key 1 `
                -Description 'an always-refused fake Item' -MaxAttempts 3 -FirstDelayMs 1 `
                -MaxDelayMs 1 -TotalBudgetMs 100
    Add-StageBReadRejection -Operation (Get-StageBBuildLabel) -Record $read
} catch { $outcome = 'REFUSED'; $hex = Get-StageBComHResult $_ }
Write-Output ('TELEMETRY|exhausted|' + $outcome + '|' + $hex + '|lines=' +
              [string]@(Get-StageBBuildRejections).Count + '|calls=' +
              [string][PccmFakeComTarget]::Calls)

# --- the sub-operation vocabulary and the label it produces -----------------
Write-Output ('SUBOPS|' + [string]@($script:StageBBuildSubOps).Count + '|' +
              (@($script:StageBBuildSubOps) -join ','))
foreach ($step in @($script:StageBBuildSubOps)) {
    $null = Set-StageBBuildOp 'saveas.xlsm'
    $outcome = 'accepted'
    try { Set-StageBBuildStep $step } catch { $outcome = 'REFUSED' }
    Write-Output ('STEP|' + $step + '|' + (Get-StageBBuildLabel) + '|' + $outcome)
}
$null = Set-StageBBuildOp 'saveas.xlsm'
$outcome = 'accepted'
try { Set-StageBBuildStep 'saveas.presave.format' } catch { $outcome = 'REFUSED' }
Write-Output ('STEP|typo|' + (Get-StageBBuildLabel) + '|' + $outcome)
# A NEW TOP-LEVEL OPERATION CLEARS THE STEP: a stale sub-operation would name a
# step that already finished.
Set-StageBBuildStep 'saveas.call'
$null = Set-StageBBuildOp 'vbcomponents.import'
Write-Output ('STEP|cleared|' + (Get-StageBBuildLabel) + '|' + (Get-StageBBuildStep) + '|ok')

# ---------------------------------------------------------------------------
# R. THE POST-OPEN READINESS GATE
# ---------------------------------------------------------------------------
# TWO WINDOWS RUNS FAILED ON THE FIRST PROPERTY READ AFTER Workbooks.Open -
# FileFormat at cc9cf8d, FullName at c66e752 - while the first Excel session in the
# same run read both and saved. So the fake here answers with NOTHING for a
# configurable number of attempts, then answers properly, and what is counted is:
# how many attempts it took, how long was waited, and whether anything was waited on
# at all when the workbook answered straight away.
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public class PccmFakeOpened {
    public static int  NameSilent   = 0;   // answer FullName with null this many times
    public static int  FormatSilent = 0;   // answer FileFormat with null this many times
    public static int  NameRefusals = 0;   // raise a retryable refusal this many times
    public static int  NameReads    = 0;
    public static int  FormatReads  = 0;
    public static string Path       = "";
    public static object FormatAnswer = null;
    public static bool PlainError   = false;
    public object FullName {
        get {
            NameReads++;
            if (PlainError) throw new InvalidOperationException("Excel accepted this and it failed");
            if (NameReads <= NameRefusals) throw new COMException("Call was rejected by callee.", -2147418111);
            if (NameReads <= NameRefusals + NameSilent) return null;
            return Path;
        }
    }
    public object FileFormat {
        get {
            FormatReads++;
            if (FormatReads <= FormatSilent) return null;
            return FormatAnswer;
        }
    }
}
'@

$readyRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('pccm_ready_' + [System.Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $readyRoot -Force
$expected = Join-Path $readyRoot 'PCCM_stageA.xlsx'
Set-Content -LiteralPath $expected -Value 'stage a' -NoNewline

foreach ($c in @(
    @{ n='ready-first';        ns=0; fs=0; nr=0; fmt=51; path='EXPECTED'; plain=$false; a=12; d=1;  b=15000 },
    @{ n='name-null-then-ok';  ns=2; fs=0; nr=0; fmt=51; path='EXPECTED'; plain=$false; a=12; d=1;  b=15000 },
    @{ n='format-null-then-ok';ns=0; fs=2; nr=0; fmt=51; path='EXPECTED'; plain=$false; a=12; d=1;  b=15000 },
    @{ n='both-null-then-ok';  ns=3; fs=3; nr=0; fmt=52; path='EXPECTED'; plain=$false; a=12; d=1;  b=15000 },
    @{ n='wrong-path';         ns=0; fs=0; nr=0; fmt=51; path='OTHER';    plain=$false; a=12; d=1;  b=15000 },
    @{ n='bad-format';         ns=0; fs=0; nr=0; fmt='xlsm'; path='EXPECTED'; plain=$false; a=12; d=1; b=15000 },
    @{ n='attempt-exhausted';  ns=99;fs=0; nr=0; fmt=51; path='EXPECTED'; plain=$false; a=3;  d=1;  b=15000 },
    @{ n='budget-exhausted';   ns=99;fs=0; nr=0; fmt=51; path='EXPECTED'; plain=$false; a=50; d=2;  b=5 })) {
    $notes.Clear()
    $null = Set-StageBBuildOp 'open.workbook'
    [PccmFakeOpened]::NameSilent   = [int]$c.ns
    [PccmFakeOpened]::FormatSilent = [int]$c.fs
    [PccmFakeOpened]::NameRefusals = [int]$c.nr
    [PccmFakeOpened]::NameReads    = 0
    [PccmFakeOpened]::FormatReads  = 0
    [PccmFakeOpened]::PlainError   = [bool]$c.plain
    [PccmFakeOpened]::FormatAnswer = $c.fmt
    [PccmFakeOpened]::Path = $(if ([string]$c.path -eq 'EXPECTED') { $expected }
                               else { Join-Path $readyRoot 'SOMETHING_ELSE.xlsx' })
    $opened = New-Object PccmFakeOpened

    $outcome = 'READY'; $attempts = -1; $waited = -1; $format = -1; $detail = ''
    try {
        $r = Wait-StageBWorkbookReady -Workbook $opened -ExpectedPath $expected `
                 -MaxAttempts ([int]$c.a) -FirstDelayMs ([int]$c.d) `
                 -MaxDelayMs ([int]$c.d) -TotalBudgetMs ([int]$c.b)
        $attempts = [int]$r.Attempts
        $waited = [int]$r.WaitedMs
        $format = [int]$r.FileFormat
        if ((Get-StageBComparablePath ([string]$r.FullName)) -ne (Get-StageBComparablePath $expected)) {
            $outcome = 'WRONG-PATH-RETURNED'
        }
    } catch {
        $outcome = 'ABORTED'
        $detail = [string]$_.Exception.Message
    }
    Write-Output ('READY|' + $c.n + '|' + $outcome + '|' + [string]$attempts + '|' +
                  [string]$waited + '|' + [string]$format + '|nameReads=' +
                  [string][PccmFakeOpened]::NameReads + '|formatReads=' +
                  [string][PccmFakeOpened]::FormatReads + '|' + $detail)
    foreach ($line in @($notes)) { Write-Output ('READYNOTE|' + $c.n + '|' + $line) }
}

# --- REFUSALS AND REAL ERRORS, THROUGH A SCRIPTED HELPER --------------------
# A PROPERTY GETTER CANNOT RAISE ON THIS HOST - it is swallowed and the value comes
# back $null. A fake that "threw" a rejection from FullName therefore exercised the
# NULL path and proved nothing about the refusal path, so both are driven through a
# scripted stand-in for the accepted helper instead. Its CONTRACT is what the gate
# depends on: a refusal it could not resolve escapes as a retryable COMException,
# and anything Excel accepted escapes as itself.
#
# LAST IN THE FILE, because PowerShell resolves a function name at CALL time and the
# last definition wins - every case above has already run against the real helper.
function Invoke-ComRetryRead {
    param($Target, [string]$Member, $Key, [string]$Description,
          [int]$MaxAttempts = 12, [int]$FirstDelayMs = 250,
          [int]$MaxDelayMs = 2000, [int]$TotalBudgetMs = 15000)
    $script:ReadyStubReads = $script:ReadyStubReads + 1
    if ($Member -eq 'FileFormat') {
        return [pscustomobject]@{ Description = $Description; Value = 51
                                  Attempts = 1; WaitedMs = 0; Rejections = '' }
    }
    $step = 'value'
    if ($script:ReadyStubIndex -lt @($script:ReadyStubScript).Count) {
        $step = [string]$script:ReadyStubScript[$script:ReadyStubIndex]
    }
    $script:ReadyStubIndex = $script:ReadyStubIndex + 1
    if ($step -eq 'refused') {
        throw (New-Object System.Runtime.InteropServices.COMException 'Call was rejected by callee.', -2147418111)
    }
    if ($step -eq 'accepted-error') {
        throw (New-Object System.Runtime.InteropServices.COMException 'Exception from HRESULT: 0x800A03EC', -2146827284)
    }
    if ($step -eq 'null') {
        return [pscustomobject]@{ Description = $Description; Value = $null
                                  Attempts = 1; WaitedMs = 0; Rejections = '' }
    }
    return [pscustomobject]@{ Description = $Description; Value = $script:ReadyStubPath
                              Attempts = 1; WaitedMs = 0; Rejections = '' }
}

$script:ReadyStubPath = $expected
foreach ($c in @(
    @{ n='refused-then-ok';      script=@('refused', 'refused', 'value') },
    @{ n='non-retryable-aborts'; script=@('accepted-error') },
    @{ n='refused-then-null-then-ok'; script=@('refused', 'null', 'value') })) {
    $notes.Clear()
    $null = Set-StageBBuildOp 'open.workbook'
    $script:ReadyStubScript = @($c.script)
    $script:ReadyStubIndex  = 0
    $script:ReadyStubReads  = 0
    $outcome = 'READY'; $attempts = -1; $waited = -1; $format = -1; $detail = ''
    try {
        $r = Wait-StageBWorkbookReady -Workbook $opened -ExpectedPath $expected `
                 -MaxAttempts 12 -FirstDelayMs 1 -MaxDelayMs 1 -TotalBudgetMs 15000
        $attempts = [int]$r.Attempts; $waited = [int]$r.WaitedMs; $format = [int]$r.FileFormat
    } catch {
        $outcome = 'ABORTED'
        $detail = [string]$_.Exception.Message
        $hex = Get-StageBComHResult $_
        if ($hex -ne '') { $detail = $hex + ' ' + $detail }
    }
    Write-Output ('READY|' + $c.n + '|' + $outcome + '|' + [string]$attempts + '|' +
                  [string]$waited + '|' + [string]$format + '|nameReads=' +
                  [string]$script:ReadyStubIndex + '|formatReads=0|' + $detail)
    foreach ($line in @($notes)) { Write-Output ('READYNOTE|' + $c.n + '|' + $line) }
}

Remove-Item -LiteralPath $readyRoot -Recurse -Force -ErrorAction SilentlyContinue

exit 0
