# ===========================================================================
# PCCM - Excel COM lifecycle policy
# ===========================================================================
# Dot-sourced by build_stage_b.ps1 and phase4_functional_test.ps1 so both use
# ONE implementation of the ownership pattern proven by the Phase-1.6 readiness
# smoke test. Duplicating it in two scripts is how the two would drift apart.
#
# POLICY - unchanged from the run that closed the readiness gate:
#
#   * Marshal.FinalReleaseComObject is PROHIBITED. Only Marshal.ReleaseComObject
#     is used. The goal is disciplined ownership, not forcibly zeroing an RCW's
#     managed reference count.
#   * ReleaseComObject's integer return is REPORTED, never interpreted as a
#     failure on its own, and never looped until zero.
#   * There is NO generic COM stack, release plan, object graph or de-duplication
#     framework. The COM graph here is small, fixed and known, so every long-lived
#     object has an explicitly named variable with an explicit release point.
#   * Release order is leaf before parent; Workbook.Close before the workbook is
#     released; Application.Quit before the application is released.
#   * Diagnostic collections hold plain data ONLY - labels, integers, strings and
#     booleans. They never hold an Excel RCW.
#   * The acceptance criterion is ACTUAL clean Excel shutdown. A forced process
#     stop is an emergency path and is NEVER converted into a PASS.
#   * No script here changes macro security, Trusted Locations, the registry or
#     any organisational policy, and none terminates an Excel process it did not
#     create and cannot positively identify.
#   * A REFUSED call may be reissued; a FAILED one may not. Only the two OLE
#     message-filter HRESULTs - the ones whose contract is that the call never
#     ran - are retried, only at the read boundary, and only within bounds that
#     are reported. There is no general "retry COM automation" behaviour here.
# ===========================================================================

Set-StrictMode -Version 2.0

$script:transientFailures = New-Object System.Collections.ArrayList

function Format-Err {
    param($ErrorRecord)
    if ($null -eq $ErrorRecord) { return 'unknown error' }
    $msg = ''
    try { $msg = [string]$ErrorRecord.Exception.Message } catch { }
    if ([string]::IsNullOrWhiteSpace($msg)) { try { $msg = [string]$ErrorRecord } catch { $msg = 'unknown error' } }
    $type = ''
    try { $type = $ErrorRecord.Exception.GetType().FullName } catch { }
    if ([string]::IsNullOrWhiteSpace($type)) { return $msg }
    return ('{0}: {1}' -f $type, $msg)
}

# Records only plain diagnostic data. Never stores the COM object.
# The caller MUST null its own variable afterwards - PowerShell parameter binding
# cannot null a caller's variable.
function Release-ComObjectSafe {
    param($Obj, [string]$Label)
    $rec = [pscustomobject]@{ Label = $Label; Status = 'SKIPPED'; Count = -1; Error = '' }
    if ($null -eq $Obj) { $rec.Error = 'reference was already null'; return $rec }
    $isCom = $false
    try {
        $isCom = [System.Runtime.InteropServices.Marshal]::IsComObject($Obj)
    } catch {
        $rec.Status = 'FAIL'; $rec.Error = 'IsComObject threw: ' + (Format-Err $_); return $rec
    }
    if (-not $isCom) { $rec.Error = 'not a COM object'; return $rec }
    try {
        $n = [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj)
        $rec.Count  = [int]$n
        $rec.Status = 'PASS'
    } catch {
        $rec.Status = 'FAIL'
        $rec.Error  = Format-Err $_
    }
    return $rec
}

# Transient objects (Range, ListObject, ListColumn, VBComponent, CodeModule,
# Shape, Property) follow acquire -> use -> release -> $var = $null inside the
# narrowest possible scope. Failures are recorded, never hidden, and they gate
# the run's outcome.
function Release-Transient {
    param($Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'FAIL') {
        $null = $script:transientFailures.Add(($rec.Label + ' :: ' + $rec.Error))
    }
}

function Get-TransientFailures { return @($script:transientFailures) }

function New-ReleaseLedger {
    param([string]$Name)
    return [pscustomobject]@{
        Name              = $Name
        Lines             = (New-Object System.Collections.ArrayList)
        Failed            = (New-Object System.Collections.ArrayList)
        Attempted         = 0
        Succeeded         = 0
        WorkbookClosed    = $false
        QuitCalled        = $false
        NaturalExit       = $false
        EmergencyRequired = $false
    }
}

# Releases one named object and records a plain-text line. The caller must set its
# variable to $null immediately after calling this.
function Invoke-NamedRelease {
    param($Ledger, $Obj, [string]$Label)
    $rec = Release-ComObjectSafe -Obj $Obj -Label $Label
    if ($rec.Status -eq 'PASS') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $Ledger.Succeeded = $Ledger.Succeeded + 1
        $null = $Ledger.Lines.Add(("      {0,-24} | PASS    | ReleaseComObject returned {1}" -f $rec.Label, $rec.Count))
    } elseif ($rec.Status -eq 'FAIL') {
        $Ledger.Attempted = $Ledger.Attempted + 1
        $null = $Ledger.Failed.Add($rec.Label)
        $null = $Ledger.Lines.Add(("      {0,-24} | FAIL    | {1}" -f $rec.Label, $rec.Error))
    } else {
        $null = $Ledger.Lines.Add(("      {0,-24} | SKIPPED | {1}" -f $rec.Label, $rec.Error))
    }
}

function Format-ReleaseLedger {
    param($Ledger)
    if ($null -eq $Ledger) { return '      (no shutdown was attempted for this instance)' }
    $out = New-Object System.Collections.ArrayList
    foreach ($l in $Ledger.Lines) { $null = $out.Add($l) }
    $failTxt = '(none)'
    if ($Ledger.Failed.Count -gt 0) { $failTxt = ($Ledger.Failed -join ', ') }
    $null = $out.Add(("      releases attempted : {0}" -f $Ledger.Attempted))
    $null = $out.Add(("      releases succeeded : {0}" -f $Ledger.Succeeded))
    $null = $out.Add(("      failed labels      : {0}" -f $failTxt))
    $null = $out.Add(("      Workbook.Close     : {0}" -f $Ledger.WorkbookClosed))
    $null = $out.Add(("      Application.Quit   : {0}" -f $Ledger.QuitCalled))
    $null = $out.Add(("      natural PID exit   : {0}" -f $Ledger.NaturalExit))
    $null = $out.Add(("      emergency required : {0}" -f $Ledger.EmergencyRequired))
    return ($out -join "`r`n")
}

# ---------------------------------------------------------------------------
# Transient COM rejection: a bounded retry at the READ boundary
# ---------------------------------------------------------------------------
# WHY THIS EXISTS. Stage-B verification failed twice, identically, on the FIRST
# per-sheet read of the reopened .xlsm:
#
#     [FAIL] Verify the reopened .xlsm
#            Dashboard: System.Runtime.InteropServices.COMException:
#            Call was rejected by callee. (0x80010001 RPC_E_CALL_REJECTED)
#
# The build had completed, the workbook had been saved, and the reopen itself
# succeeded. 'Dashboard' is sheets[0] of the manifest, so this is the first
# iteration of the CodeName loop - and the 13 sheets after it, all 32 modules
# and all 11 buttons then verified. The workbook was never in question. The
# call was refused.
#
# WHAT IS RETRIED, AND WHY THE SET CANNOT BE WIDER. RPC_E_CALL_REJECTED and
# RPC_E_SERVERCALL_RETRYLATER are the two results an OLE message filter returns
# when the server declines to accept an incoming call. Their contract is that
# the call was NOT delivered: it did not run, nothing executed, no state moved.
# THAT is what makes issuing it again safe, and it is the entire justification.
# A COMException from a call Excel accepted describes something that actually
# happened; repeating it would be guessing, so 'retry a COMException' is exactly
# the generalisation this must not make.
#
# VBA_E_IGNORE (0x800AC472) is deliberately NOT in the set. It comes from a
# different mechanism - Excel sitting in a UI or modal state - and a run parked
# behind a dialog should fail and say so rather than spin out its budget.
#
# WHAT THIS HELPER CANNOT DO, BY CONSTRUCTION. It takes an object and a member
# NAME. It reads a property, or it calls .Item(key). There is no scriptblock
# parameter, so no write, no Open, no SaveAs, no Import, no Protect and no Run
# can be expressed through it at all. The read-only guarantee is the shape of
# the API rather than a rule a future edit has to remember.
#
# AND IT ASSIGNS, NEVER EMITS. The value can be an Excel COLLECTION object
# (Worksheets, Shapes, VBComponents). Writing one of those to a PowerShell
# pipeline makes PowerShell enumerate it into its members - the Phase-10 Run-3
# record collapse in a new costume. So the read is an ASSIGNMENT, and the caller
# gets a record whose .Value is untouched.

# HRESULT -> name. Keyed by the SIGNED int32 the CLR reports as ErrorCode.
$script:ComRetryableHResults = @{
    -2147418111 = 'RPC_E_CALL_REJECTED (0x80010001)'
    -2147417846 = 'RPC_E_SERVERCALL_RETRYLATER (0x8001010A)'
}

# Plain diagnostic data only - labels, integers and strings. Never a COM RCW.
$script:comRetryLedger  = New-Object System.Collections.ArrayList
$script:comRetryWaitMs  = 0

# Returns the retryable rejection's NAME, or '' when the error is anything else.
# '' is the answer that means "do not retry this", so every path that cannot
# positively identify a rejection returns it.
function Get-ComRejectionName {
    param($ErrorRecord)
    if ($null -eq $ErrorRecord) { return '' }
    $ex = $null
    try { $ex = $ErrorRecord.Exception } catch { return '' }
    # PowerShell surfaces a COM failure bare in some paths and wrapped in a
    # MethodInvocationException in others, so the chain is walked - to a BOUNDED
    # depth, because a cyclic InnerException chain would otherwise hang here.
    for ($depth = 0; $depth -lt 5; $depth++) {
        if ($null -eq $ex) { return '' }
        if ($ex -is [System.Runtime.InteropServices.COMException]) {
            $code  = 0
            $known = $false
            try { $code = [int]$ex.ErrorCode; $known = $true } catch { $known = $false }
            if ($known -and $script:ComRetryableHResults.ContainsKey($code)) {
                return [string]$script:ComRetryableHResults[$code]
            }
        }
        $next = $null
        try { $next = $ex.InnerException } catch { $next = $null }
        $ex = $next
    }
    return ''
}

# Reads ONE member of ONE COM object, retrying only a refused call.
#
#   -Member <name>            read that property
#   -Member Item -Key <k>     call .Item(k)
#
# Bounds: at most $MaxAttempts attempts, each delay at most $MaxDelayMs, and the
# TOTAL time spent waiting never exceeds $TotalBudgetMs. When the bounds run out
# the ORIGINAL error is rethrown, unchanged, and the caller fails exactly as it
# would have before this helper existed.
function Invoke-ComRetryRead {
    param(
        $Target,
        [string]$Member,
        $Key,
        [string]$Description,
        [int]$MaxAttempts   = 12,
        [int]$FirstDelayMs  = 250,
        [int]$MaxDelayMs    = 2000,
        [int]$TotalBudgetMs = 15000
    )
    if ([string]::IsNullOrWhiteSpace($Description)) {
        throw 'Invoke-ComRetryRead: every retried read must name the action it is retrying.'
    }
    if ($null -eq $Target) {
        throw ('Invoke-ComRetryRead: no target object for ' + $Description + '.')
    }
    if ([string]::IsNullOrWhiteSpace($Member)) {
        throw ('Invoke-ComRetryRead: no member name for ' + $Description + '.')
    }
    $useItem = $PSBoundParameters.ContainsKey('Key')
    if ($useItem -and $Member -ne 'Item') {
        throw ('Invoke-ComRetryRead: a key may only be passed to Item, not to ' + $Member + '.')
    }
    if ($MaxAttempts -lt 1) { throw 'Invoke-ComRetryRead: MaxAttempts must be at least 1.' }
    if ($TotalBudgetMs -lt 0) { throw 'Invoke-ComRetryRead: TotalBudgetMs may not be negative.' }

    $value      = $null
    $attempts   = 0
    $waitedMs   = 0
    $rejections = New-Object System.Collections.ArrayList
    $delay      = $FirstDelayMs

    # THE BOUND IS IN THE HEAD. A loop whose only exits are a break and a throw
    # is bounded in fact and unbounded to a reader, and on a Windows machine a
    # harness that hangs is worse than one that fails: it holds an Excel process
    # open and produces no transcript at all.
    $answered = $false
    while ($attempts -lt $MaxAttempts) {
        $attempts = $attempts + 1
        try {
            # ASSIGNMENT, not pipeline output. See the note above.
            if ($useItem) { $value = $Target.Item($Key) } else { $value = $Target.$Member }
            $answered = $true
            break
        } catch {
            $name = Get-ComRejectionName $_
            if ([string]::IsNullOrWhiteSpace($name)) {
                # Not a refused call. Excel accepted this one and it failed, so it
                # is a real result and it is rethrown untouched.
                throw
            }
            $null = $rejections.Add($name)
            $reason = ''
            if ($attempts -ge $MaxAttempts) {
                $reason = ('attempt limit of ' + [string]$MaxAttempts + ' reached')
            } elseif (($waitedMs + $delay) -gt $TotalBudgetMs) {
                $reason = ('retry budget of ' + [string]$TotalBudgetMs + ' ms would be exceeded')
            }
            if ($reason -ne '') {
                $null = $script:comRetryLedger.Add(
                    ('EXHAUSTED  ' + $Description + ' :: ' + [string]$attempts + ' attempt(s), ' +
                     [string]$waitedMs + ' ms waited, ' + $reason + '; rejections: ' +
                     (($rejections | Select-Object -Unique) -join ', ')))
                $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs
                throw
            }
            Start-Sleep -Milliseconds $delay
            $waitedMs = $waitedMs + $delay
            $delay = [Math]::Min(($delay + $FirstDelayMs), $MaxDelayMs)
        }
    }

    # UNREACHABLE WHILE THE CATCH THROWS ON EXHAUSTION, and it still throws rather
    # than falling through to a record with a $null .Value. An edit that changes
    # the head bound must not be able to turn "never answered" into an answer.
    if (-not $answered) {
        throw ('Invoke-ComRetryRead: ' + $Description + ' was never answered after ' +
               [string]$attempts + ' attempt(s).')
    }

    if ($attempts -gt 1) {
        $null = $script:comRetryLedger.Add(
            ('RETRIED    ' + $Description + ' :: answered on attempt ' + [string]$attempts +
             ' after ' + [string]$waitedMs + ' ms; rejections: ' +
             (($rejections | Select-Object -Unique) -join ', ')))
        $script:comRetryWaitMs = $script:comRetryWaitMs + $waitedMs
    }

    # ONE pipeline item: a record. The bare value is never emitted, so a COM
    # collection cannot be enumerated on its way back to the caller.
    return [pscustomobject]@{
        Description = $Description
        Value       = $value
        Attempts    = $attempts
        WaitedMs    = $waitedMs
        Rejections  = (($rejections | Select-Object -Unique) -join ', ')
    }
}

function Get-ComRetryLedger    { return @($script:comRetryLedger) }
function Get-ComRetryWaitTotal { return [int]$script:comRetryWaitMs }

# ---------------------------------------------------------------------------
# Process identity
# ---------------------------------------------------------------------------
$script:haveNative = $false
try {
    Add-Type -Namespace PccmStageB -Name Native -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll", SetLastError=true)]
public static extern int GetWindowThreadProcessId(System.IntPtr hWnd, out int lpdwProcessId);
'@ -ErrorAction Stop
    $script:haveNative = $true
} catch { $script:haveNative = $false }

# Captures a strong identity for the Excel instance WE created.
#   Source = 'HWND'     -> pid derived from our own Application.Hwnd. Force-stop permitted.
#   Source = 'FALLBACK' -> diagnostic only. Force-stop NOT permitted.
function Get-ExcelIdentity {
    param($ExcelApp, [int[]]$PreExistingPids)
    $id = [pscustomobject]@{
        ProcessId      = 0
        Source         = 'UNKNOWN'
        HasStartTime   = $false
        StartTimeTicks = [long]0
        StartTimeText  = 'unavailable'
        ProcessName    = ''
    }
    if ($script:haveNative) {
        try {
            $procId = 0
            $hwnd = New-Object System.IntPtr ([int]$ExcelApp.Hwnd)
            $null = [PccmStageB.Native]::GetWindowThreadProcessId($hwnd, [ref]$procId)
            if ($procId -gt 0) { $id.ProcessId = $procId; $id.Source = 'HWND' }
        } catch { }
    }
    if ($id.ProcessId -le 0) {
        try {
            $now = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
            $new = @($now | Where-Object { $PreExistingPids -notcontains $_ })
            if ($new.Count -eq 1) { $id.ProcessId = [int]$new[0]; $id.Source = 'FALLBACK' }
        } catch { }
    }
    if ($id.ProcessId -gt 0) {
        try {
            $p = Get-Process -Id $id.ProcessId -ErrorAction Stop
            $id.ProcessName = [string]$p.ProcessName
            try {
                $st = $p.StartTime
                $id.HasStartTime   = $true
                $id.StartTimeTicks = [long]$st.Ticks
                $id.StartTimeText  = $st.ToString('yyyy-MM-dd HH:mm:ss.fff')
            } catch { }
        } catch { }
    }
    return $id
}

# Force-stop is allowed only when all three identity facts still hold.
function Test-IsOurExcelProcess {
    param($Identity)
    if ($null -eq $Identity) { return $false }
    if ($Identity.ProcessId -le 0) { return $false }
    if ($Identity.Source -ne 'HWND') { return $false }
    if (-not $Identity.HasStartTime) { return $false }
    $p = Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue
    if ($null -eq $p) { return $false }
    if ($p.ProcessName -notmatch '^(?i)excel$') { return $false }
    try { if ([long]$p.StartTime.Ticks -ne [long]$Identity.StartTimeTicks) { return $false } } catch { return $false }
    return $true
}

# Returns $true only when the process is confirmed gone. Never force-stops.
function Wait-ExcelExit {
    param($Identity, [int]$TimeoutSeconds = 25)
    if ($null -eq $Identity -or $Identity.ProcessId -le 0) { return $false }
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue)) { return $true }
        Start-Sleep -Milliseconds 400
    }
    return ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue))
}

# Emergency path only. Never called from a check; never converts a failed graceful
# close into a PASS. It only ever touches a process this script created and can
# still positively identify.
function Invoke-EmergencyExcelCleanup {
    param($Identity, [string]$Label)
    if ($null -eq $Identity -or $Identity.ProcessId -le 0) {
        return ("{0}: no process identity was captured; nothing to clean up." -f $Label)
    }
    if ($null -eq (Get-Process -Id $Identity.ProcessId -ErrorAction SilentlyContinue)) {
        return ("{0}: pid {1} already exited." -f $Label, $Identity.ProcessId)
    }
    if (Test-IsOurExcelProcess $Identity) {
        try {
            Stop-Process -Id $Identity.ProcessId -Force -ErrorAction Stop
            return ("{0}: EMERGENCY - force-stopped PCCM Stage-B Excel pid {1} (identity verified: HWND-derived, name EXCEL, StartTime {2}). This run is NOT a pass." -f $Label, $Identity.ProcessId, $Identity.StartTimeText)
        } catch {
            return ("{0}: force-stop of pid {1} FAILED: {2}. Please close that Excel window manually." -f $Label, $Identity.ProcessId, $_.Exception.Message)
        }
    }
    return ("{0}: pid {1} is still running but its identity could NOT be verified (source={2}, startTimeCaptured={3}). NOT terminated. Please close the leftover Excel window manually." -f $Label, $Identity.ProcessId, $Identity.Source, $Identity.HasStartTime)
}

function Get-PreExistingExcelPids {
    $pids = @()
    try {
        $pids = @(Get-Process -Name 'excel' -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    } catch { $pids = @() }
    return $pids
}

# ---------------------------------------------------------------------------
# Trust Center prerequisite reporting
# ---------------------------------------------------------------------------
# The VBA project object model requires "Trust access to the VBA project object
# model" in the Trust Center. NOTHING here enables it: the script reports the
# prerequisite and stops. Changing that setting programmatically would be exactly
# the security bypass this project has refused from the first readiness run.
function Test-TrustAccessError {
    param($ErrorRecord)
    $text = ''
    try { $text = [string]$ErrorRecord.Exception.Message } catch { }
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return ($text -match '(?i)programmatic access|not trusted|trust access|VBProject')
}

function Get-TrustAccessGuidance {
    return @(
        'Excel refused programmatic access to the VBA project.'
        ''
        'This is a Trust Center setting a person must make, once, on this machine:'
        '  File > Options > Trust Center > Trust Center Settings > Macro Settings'
        '  tick "Trust access to the VBA project object model"'
        ''
        'This script will NOT change that setting, will not lower macro security,'
        'will not edit the registry and will not alter Trusted Locations.'
    ) -join "`r`n"
}
