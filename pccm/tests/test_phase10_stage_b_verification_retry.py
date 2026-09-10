#!/usr/bin/env python3
"""P10 / STAGE-B VERIFICATION: THE BOUNDED COM READ RETRY.

WHAT THIS BATCH SETTLED. Two consecutive Stage-B runs built the .xlsm cleanly -
14 CodeNames, 32 modules, 11 buttons, protection applied, saved, COM released -
and then failed verification identically:

    [FAIL] Verify the reopened .xlsm
           Dashboard: System.Runtime.InteropServices.COMException:
           Call was rejected by callee. (0x80010001 RPC_E_CALL_REJECTED)

'Dashboard' is sheets[0] of the manifest and is no button's shape name, so the
failing statement is the FIRST iteration of the CodeName loop. Everything after
it verified. The workbook was never in question; the call was refused.

WHY A RETRY IS NOT A GUESS HERE, AND WHERE THE LINE IS. RPC_E_CALL_REJECTED and
RPC_E_SERVERCALL_RETRYLATER are OLE message-filter results whose contract is
that the call was NOT delivered. Nothing ran, so reissuing it is safe. Every
other COMException describes a call Excel accepted, and reissuing one of those
would be a guess. These controls exist to keep that distinction from eroding
into "retry COM automation", which is the failure mode of every retry helper
that ever went wrong.

THESE TESTS DO NOT CLAIM THE STAGE-B RUNTIME IS CORRECT. PowerShell cannot be
parsed or run here. Only a clean Windows run can claim that, and this batch has
not had one.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent

BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
LIFECYCLE_PS1 = BOOTSTRAP / "com_lifecycle.ps1"
BUILD_PS1 = BOOTSTRAP / "build_stage_b.ps1"

# The two HRESULTs, written as HEX here and converted, so this file never
# restates the signed int the source uses. A transcription slip in either place
# shows up as a mismatch instead of agreeing with itself.
RETRYABLE_HEX = ("0x80010001", "0x8001010A")
# Named so the exclusion is a decision on the record, not an omission.
EXCLUDED_HEX = ("0x800AC472",)

_MEMO: dict[str, str] = {}


def _src(key: str, path: Path) -> str:
    if key not in _MEMO:
        _MEMO[key] = path.read_text(encoding="utf-8")
    return _MEMO[key]


def _lifecycle() -> str:
    return _src("lifecycle", LIFECYCLE_PS1)


def _build() -> str:
    return _src("build", BUILD_PS1)


def _strip(text: str) -> str:
    """PowerShell with comments blanked in place. Line numbering is preserved."""
    text = re.sub(r"<#.*?#>", lambda m: "\n" * m.group(0).count("\n"), text, flags=re.DOTALL)
    out = []
    for line in text.splitlines():
        if line.lstrip().startswith("#"):
            out.append("")
            continue
        cut = line.find(" #")
        if cut != -1 and line.count('"') % 2 == 0 and line.count("'") % 2 == 0:
            line = line[:cut]
        out.append(line)
    return "\n".join(out)


def _lifecycle_code() -> str:
    key = "lifecycle_code"
    if key not in _MEMO:
        _MEMO[key] = _strip(_lifecycle())
    return _MEMO[key]


def _build_code() -> str:
    key = "build_code"
    if key not in _MEMO:
        _MEMO[key] = _strip(_build())
    return _MEMO[key]


def _function(name: str) -> str:
    """One PowerShell function body from com_lifecycle.ps1, by brace depth."""
    code = _lifecycle_code()
    start = code.index(f"function {name} ")
    depth = 0
    seen = False
    for index in range(start, len(code)):
        if code[index] == "{":
            depth += 1
            seen = True
        elif code[index] == "}":
            depth -= 1
            if seen and depth == 0:
                return code[start : index + 1]
    raise AssertionError(f"{name} is not closed")


def _verify_block() -> str:
    """The FRESH-instance verification block, up to its shutdown ledger."""
    code = _build_code()
    start = code.index("# 8. Verify in a FRESH instance") if False else code.index(
        "$excel2 = $null; $workbooks2 = $null")
    end = code.index("$rel2 = New-ReleaseLedger 'verification instance'")
    return code[start:end]


def _build_block() -> str:
    """The build instance's own block: everything it writes into the workbook."""
    code = _build_code()
    start = code.index("$excel = New-Object -ComObject Excel.Application")
    end = code.index("# --- shutdown of the build instance") if False else code.index(
        "$rel1 = New-ReleaseLedger 'build instance'")
    return code[start:end]


def _manifest() -> dict:
    import json
    return json.loads((PCCM_ROOT / "build" / "stage_b_manifest.json").read_text(encoding="utf-8"))


# ===========================================================================
# A. THE RETRY SET: EXACTLY THE REFUSED CALLS, AND NOTHING ELSE
# ===========================================================================
def test_01_the_retry_table_holds_exactly_the_two_message_filter_hresults() -> None:
    """REQUIRED CONTROL. Two entries, and they are the two whose contract is that
    the call never ran."""
    table = _function("Get-ComRejectionName") if False else _lifecycle_code()
    start = table.index("$script:ComRetryableHResults = @{")
    end = table.index("}", start)
    body = table[start:end]
    entries = re.findall(r"^\s*(-?\d+)\s*=\s*'([^']+)'", body, re.M)
    assert len(entries) == 2, entries
    codes = {int(code) for code, _ in entries}
    assert codes == {int(h, 16) - 2**32 for h in RETRYABLE_HEX}, sorted(codes)
    names = " ".join(name for _, name in entries)
    assert "RPC_E_CALL_REJECTED" in names and "RPC_E_SERVERCALL_RETRYLATER" in names, names


def test_02_each_signed_code_agrees_with_the_hex_it_claims_to_be() -> None:
    """THE NAME CARRIES THE HEX, so a wrong signed int cannot hide behind a
    plausible-looking name -- which is exactly how Probe Run 2 went wrong."""
    body = _lifecycle_code()
    pairs = re.findall(r"^\s*(-?\d+)\s*=\s*'[A-Z_]+ \((0x[0-9A-Fa-f]{8})\)'", body, re.M)
    assert len(pairs) == 2, pairs
    for signed, hexcode in pairs:
        assert int(signed) == int(hexcode, 16) - 2**32, (signed, hexcode)


def test_03_vba_e_ignore_is_excluded_on_the_record() -> None:
    """AN OMISSION IS NOT A DECISION. The excluded code is named, with its reason,
    and it is named ONLY in prose -- never in the retry table."""
    source = _lifecycle()
    for hexcode in EXCLUDED_HEX:
        assert hexcode in source, f"{hexcode} is not mentioned at all"
        assert "NOT retried" in source or "NOT in the set" in source
    code = _lifecycle_code()
    for hexcode in EXCLUDED_HEX:
        assert hexcode not in code, f"{hexcode} reached the executable retry set"
        assert str(int(hexcode, 16) - 2**32) not in code, f"{hexcode} is retried by its signed value"


def test_04_the_retry_decision_is_made_only_by_the_named_classifier() -> None:
    """NO SECOND OPINION. One function decides retryability; the loop asks it and
    nothing else. A branch that tested the exception TYPE would retry every
    COMException, which is the generalisation this refuses."""
    body = _function("Invoke-ComRetryRead")
    assert "Get-ComRejectionName" in body
    for banned in ("-is [System.Runtime.InteropServices.COMException]",
                   "COMException", "HResult", "ErrorCode"):
        assert banned not in body, f"the retry loop inspects {banned} itself"


def test_05_an_unrecognised_error_is_rethrown_untouched() -> None:
    """A BARE `throw` IN THE CATCH. Not a new exception, not a wrapped one, not a
    default value: whatever Excel said is what the caller sees."""
    body = _function("Invoke-ComRetryRead")
    guard = body.index("$name = Get-ComRejectionName $_")
    tail = body[guard:]
    blank = tail.index("IsNullOrWhiteSpace($name)")
    following = tail[blank : blank + 400]
    assert re.search(r"\{\s*\n\s*(#[^\n]*\n\s*)*(#[^\n]*\n\s*)*throw\s*\n", following) or \
        re.search(r"\n\s*throw\s*\n", following), following[:300]
    assert "throw New-Object" not in body
    assert re.search(r"\$value\s+= \$null", body), "there must be no fabricated default"


def test_06_a_refused_call_is_never_reported_as_an_answered_one() -> None:
    """EXHAUSTION RETHROWS. The helper has no path that returns a record after the
    bounds run out, so 'could not read it, assume okay' cannot be written."""
    body = _function("Invoke-ComRetryRead")
    exhausted = body.index("EXHAUSTED")
    after = body[exhausted:]
    ret = after.index("throw")
    assert "return" not in after[:ret], after[:ret]


# ===========================================================================
# B. THE BOUNDS, AND THAT THEY ARE REALLY BOUNDS
# ===========================================================================
def test_10_all_three_bounds_exist_with_finite_defaults() -> None:
    """REQUIRED CONTROL: bounded attempts, bounded delay, finite total duration."""
    body = _function("Invoke-ComRetryRead")
    defaults = dict(re.findall(r"\[int\]\$(\w+)\s*=\s*(\d+)", body))
    for name in ("MaxAttempts", "FirstDelayMs", "MaxDelayMs", "TotalBudgetMs"):
        assert name in defaults, (name, defaults)
        assert int(defaults[name]) > 0, (name, defaults[name])
    assert int(defaults["MaxAttempts"]) <= 20, defaults["MaxAttempts"]
    assert int(defaults["MaxDelayMs"]) <= 5000, defaults["MaxDelayMs"]
    assert int(defaults["TotalBudgetMs"]) <= 60000, defaults["TotalBudgetMs"]


def test_11_the_attempt_limit_is_enforced_in_the_loop() -> None:
    body = _function("Invoke-ComRetryRead")
    assert "$attempts -ge $MaxAttempts" in body, "nothing stops the attempts"


def test_12_the_total_budget_is_checked_before_each_sleep() -> None:
    """A PER-ATTEMPT CAP IS NOT A TOTAL. Twelve two-second waits are bounded per
    attempt and unbounded in aggregate, so the budget is checked against the time
    ALREADY waited plus the wait about to happen."""
    body = _function("Invoke-ComRetryRead")
    assert "($waitedMs + $delay) -gt $TotalBudgetMs" in body, body
    # THE COMPARISON, not the parameter declaration. The declaration sits at the
    # top of the function and would satisfy any ordering check trivially.
    check = body.index("($waitedMs + $delay) -gt $TotalBudgetMs")
    sleep = body.index("Start-Sleep")
    assert check < sleep, "the budget is checked after sleeping"


def test_13_the_delay_is_capped_and_never_shrinks_the_way_out() -> None:
    body = _function("Invoke-ComRetryRead")
    assert "[Math]::Min((" in body and "$MaxDelayMs)" in body, body
    assert "Start-Sleep -Milliseconds $delay" in body


def test_14_the_only_sleep_in_either_script_is_inside_the_bounded_retry() -> None:
    """NO BLIND SLEEP. The fix is not 'wait a bit and hope'; a Start-Sleep outside
    the retry loop, or a fixed one before verification, is refused here."""
    lifecycle = _lifecycle_code()
    sleeps = [m.start() for m in re.finditer(r"Start-Sleep", lifecycle)]
    retry = _function("Invoke-ComRetryRead")
    wait_exit = _function("Wait-ExcelExit")
    for at in sleeps:
        window = lifecycle[max(0, at - 400) : at + 60]
        assert any(window in block or lifecycle[at : at + 40] in block
                   for block in (retry, wait_exit)), lifecycle[at - 200 : at + 60]
    assert "Start-Sleep" not in _build_code(), "the bootstrap sleeps outside the retry"


def test_14a_the_retry_loop_carries_its_bound_in_its_own_head() -> None:
    """A LOOP BOUNDED ONLY BY ITS BODY IS UNBOUNDED TO A READER, and on Windows a
    harness that hangs is worse than one that fails: it holds an Excel process
    open and produces no transcript at all. This is the Phase-5 rule
    (test_218_every_wait_loop_in_the_gate_b_harness_is_bounded), which caught this
    helper's first draft, restated where the helper lives."""
    body = _function("Invoke-ComRetryRead")
    heads = re.findall(r"while\s*\(([^)]*)\)\s*\{", body)
    assert heads, body
    for head in heads:
        assert re.search(r"-lt |-le |-gt |-ge |Get-Date", head), head
    assert "while ($true)" not in body
    assert "do {" not in body, "a do-loop keeps its bound out of its head"


def test_14b_a_loop_that_ended_without_an_answer_never_returns_one() -> None:
    """THE HEAD BOUND CREATES A NEW WAY OUT, so it gets a guard. Falling out of
    the loop throws; it does not reach a record whose .Value is $null."""
    body = _function("Invoke-ComRetryRead")
    assert "$answered = $true" in body
    guard = body.index("if (-not $answered)")
    record = body.index("return [pscustomobject]@{")
    assert guard < record, "the guard runs after the record is built"
    assert "throw" in body[guard : record], body[guard:record]


def test_15_the_wait_actually_accumulates() -> None:
    """A BUDGET COMPARED AGAINST A COUNTER THAT NEVER MOVES IS NOT A BUDGET."""
    body = _function("Invoke-ComRetryRead")
    assert "$waitedMs = $waitedMs + $delay" in body, body


# ===========================================================================
# C. WHAT THE HELPER CANNOT BE MADE TO DO
# ===========================================================================
def test_20_the_helper_takes_no_scriptblock_and_so_can_express_no_write() -> None:
    """THE READ-ONLY GUARANTEE IS THE SHAPE OF THE API. A scriptblock parameter
    would let any call at all be retried, including SaveAs and Import, and the
    guarantee would become a rule someone has to remember."""
    body = _function("Invoke-ComRetryRead")
    assert "[scriptblock]" not in body, "a scriptblock parameter reopens writes"
    assert "& $" not in body and "Invoke-Command" not in body and "Invoke-Expression" not in body


def test_21_exactly_two_operations_are_expressible() -> None:
    """READ A PROPERTY, OR CALL .Item(key). Nothing else has a syntax here."""
    body = _function("Invoke-ComRetryRead")
    assert "$value = $Target.Item($Key)" in body
    assert "$value = $Target.$Member" in body
    calls = re.findall(r"\$Target\.[\w$]+", body)
    assert set(calls) == {"$Target.Item", "$Target.$Member"}, sorted(set(calls))


def test_22_a_key_may_only_be_passed_to_item() -> None:
    """OR $Member COULD NAME A METHOD, and the helper would invoke it."""
    body = _function("Invoke-ComRetryRead")
    assert "$useItem -and $Member -ne 'Item'" in body, body
    assert "a key may only be passed to Item" in _function("Invoke-ComRetryRead")


def test_23_the_read_is_an_assignment_never_a_pipeline_write() -> None:
    """THE RUN-3 COLLAPSE, ONE LAYER OUT. The value can be an Excel COLLECTION
    object; writing one of those to a PowerShell pipeline enumerates it into its
    members. Both reads assign, and the return is a record, so the bare value
    never reaches a pipeline."""
    body = _function("Invoke-ComRetryRead")
    # EVERY read, not every line. `if ($useItem) { $Target.Item($Key) }` starts
    # with an `if` and emits its result to the pipeline, so a line-prefix check
    # would wave it straight through.
    reads = [m.start() for m in re.finditer(r"\$Target\.", body)]
    assert len(reads) == 2, reads
    for at in reads:
        preceding = body[:at]
        assert preceding.endswith("$value = "), body[max(0, at - 60) : at + 40]
    assert "return [pscustomobject]@{" in body, "the bare value is returned"
    assert "return $value" not in body


def test_24_every_retried_read_must_name_itself() -> None:
    """REQUIRED CONTROL: the exact action being retried is named in diagnostics."""
    body = _function("Invoke-ComRetryRead")
    assert "IsNullOrWhiteSpace($Description)" in body
    assert "must name the action it is retrying" in body
    # Call sites use backtick continuations, so the lines are rejoined first --
    # a per-line sweep would pass trivially by never seeing the second half.
    joined = re.sub(r"`\n\s*", " ", _build_code())
    sites = [line for line in joined.splitlines() if "Invoke-ComRetryRead" in line]
    assert len(sites) >= 12, sites
    for site in sites:
        assert "-Description" in site, site


def test_25_the_classifier_walks_the_inner_exception_chain_but_not_forever() -> None:
    """POWERSHELL SOMETIMES WRAPS A COM FAILURE AND SOMETIMES DOES NOT, so the
    chain is walked -- bounded, because a cyclic chain would hang the build."""
    body = _function("Get-ComRejectionName")
    assert "InnerException" in body
    depth = re.search(r"\$depth -lt (\d+)", body)
    assert depth and 1 < int(depth.group(1)) <= 10, body


def test_26_the_classifier_returns_do_not_retry_whenever_it_cannot_be_sure() -> None:
    """EVERY UNCERTAIN PATH RETURNS ''. A classifier that fell through to 'retry'
    would retry the errors it failed to understand."""
    body = _function("Get-ComRejectionName")
    returns = [r.strip().rstrip("}").strip() for r in re.findall(r"return ([^\n]+)", body)]
    assert returns, body
    assert returns[-1] == "''", returns
    assert returns.count("''") >= 4, returns


# ===========================================================================
# D. WHERE IT IS APPLIED, AND WHERE IT IS NOT
# ===========================================================================
def test_30_every_com_read_in_the_verification_block_goes_through_the_helper() -> None:
    """REQUIRED CONTROL, AND THE POINT OF THE BATCH. Not the literal 'Dashboard'
    line: every read of the reopened workbook."""
    block = _verify_block()
    reads = re.findall(r"\$(wb2|worksheets2|vbproj2|vbcomps2|ws|shapes|shp|c)\.(\w+)", block)
    offenders = [f"${var}.{member}" for var, member in reads
                 if member not in ("Item",) or True]
    # Only Close is permitted as a direct call, and it belongs to shutdown, which
    # this slice stops short of. Everything else must be a retried read.
    assert not offenders, sorted(set(offenders))


def test_31_the_named_verification_reads_are_all_retried() -> None:
    """NAMED, NOT COUNTED. Every read the reopen verification performs is listed,
    so a read that quietly stops being retried fails here rather than on Windows."""
    block = _verify_block()
    for target, member in (
        ("$wb2", "FileFormat"),
        ("$wb2", "Worksheets"),
        ("$wb2", "VBProject"),
        ("$vbproj2", "VBComponents"),
        ("$vbcomps2", "Count"),
        ("$ws", "CodeName"),
        ("$ws", "Shapes"),
        ("$shp", "OnAction"),
        ("$c", "Name"),
    ):
        assert f"-Target {target} -Member '{member}'" in block, f"{target}.{member} is not retried"
    for target in ("$worksheets2", "$vbcomps2", "$shapes"):
        assert f"-Target {target} -Member 'Item' -Key " in block, f"{target}.Item is not retried"


def test_32_the_failing_read_itself_is_covered() -> None:
    """THE ONE THAT ACTUALLY FAILED. Worksheets.Item(<first sheet>) and its
    CodeName are both inside the retry, and the sheet still comes from the
    manifest rather than a literal."""
    block = _verify_block()
    assert "-Target $worksheets2 -Member 'Item' -Key ([string]$sheet.name)" in block
    assert "-Target $ws -Member 'CodeName'" in block
    first = _manifest()["sheets"][0]["name"]
    assert first == "Dashboard", first
    assert f'"{first}"' not in _build_code(), "the failing sheet name got hard-coded"


def test_33_the_build_block_is_not_wrapped() -> None:
    """REQUIRED CONTROL: do not broadly wrap all Excel automation in retries. The
    build block WRITES -- SaveAs, Import, AddShape, Protect, Save -- and a write
    Excel may or may not have accepted is not something to reissue on a guess."""
    assert "Invoke-ComRetryRead" not in _build_block(), "the writing build block retries"


# The complete set of members the retry helper is allowed to name. Read-only,
# every one of them, and NAMED -- a proximity check around the write call sites
# lets a retried write slip past simply by sitting on its own line.
RETRYABLE_MEMBERS = {"FileFormat", "Worksheets", "VBProject", "VBComponents",
                     "Count", "CodeName", "Shapes", "OnAction", "Name", "Item"}


def test_34_opening_the_workbook_is_never_retried() -> None:
    """OPEN IS NOT IDEMPOTENT. Reissuing it would open a second copy. Nor is any
    other member that changes something: the helper may name these and no more."""
    named = set(re.findall(r"-Member '(\w+)'", _build_code()))
    assert named, "no retried read names a member at all"
    assert named <= RETRYABLE_MEMBERS, sorted(named - RETRYABLE_MEMBERS)
    for banned in ("Open", "SaveAs", "Save", "Import", "AddShape", "Protect",
                   "Unprotect", "Run", "Delete", "Close", "Quit", "Add"):
        assert banned not in named, f"{banned} is retried"


def test_35_no_release_or_shutdown_call_is_retried() -> None:
    """The continuations are rejoined first, so the check sees a whole call."""
    joined = re.sub(r"`\n\s*", " ", _build_code())
    for line in joined.splitlines():
        if "Invoke-ComRetryRead" in line:
            for banned in ("Release", "Quit", "Close", "'Save"):
                assert banned not in line, line


# ===========================================================================
# E. NOTHING IS VERIFIED LESS
# ===========================================================================
def test_40_all_fourteen_codenames_thirty_two_modules_eleven_buttons_remain() -> None:
    """REQUIRED CONTROL: the retry may not buy its safety with coverage."""
    manifest = _manifest()
    assert len(manifest["sheets"]) == 14
    assert len(manifest["vba"]["modules"]) == 32
    assert len(manifest["buttons"]) == 11
    block = _verify_block()
    assert "foreach ($sheet in $manifest.sheets)" in block
    assert "foreach ($m in $manifest.vba.modules)" in block
    assert "foreach ($button in $manifest.buttons)" in block
    assert "$problems += (\"{0}: CodeName persisted as '{1}'" in block
    assert "did not persist" in block
    assert "OnAction persisted as" in block


def test_41_a_read_that_cannot_be_answered_still_fails_the_step() -> None:
    """'COULD NOT READ DASHBOARD, ASSUME OKAY' HAS NO PATH. The per-sheet catch
    still records a problem, and a non-empty problem list is still a FAIL."""
    block = _verify_block()
    assert '$problems += ("{0}: {1}" -f $sheet.name, (Format-Err $_))' in block
    assert "if ($problems.Count -gt 0) {" in block
    assert "Add-Step 'Verify the reopened .xlsm' 'FAIL' ($problems -join '; ')" in block


def test_42_nothing_is_suppressed_to_make_verification_easier() -> None:
    """THE VERIFICATION INSTANCE MUST OPEN THE WORKBOOK AS IT WILL OPEN FOR A
    PERSON: macros enabled, events on, Workbook_Open run, protection applied."""
    block = _verify_block()
    for banned in ("EnableEvents", "AutomationSecurity", "Unprotect", "ProtectContents = ",
                   "-ErrorAction SilentlyContinue", "On Error", "DisplayAlerts = $true"):
        assert banned not in block, f"the verification block sets or bypasses {banned}"
    assert "EnableEvents" not in _build_code(), "the bootstrap disables events somewhere"


def test_43_the_reopen_is_still_a_genuinely_fresh_instance() -> None:
    code = _build_code()
    assert code.count("New-Object -ComObject Excel.Application") == 2
    assert "$wb2 = $workbooks2.Open($stageBPath)" in code


def test_44_the_workbook_open_handler_is_still_proved_at_build_time() -> None:
    """WHERE Workbook_Open EXISTENCE ACTUALLY LIVES: the document module is written
    and READ BACK, and every declared event must appear in what came back. This
    batch did not move it and must not."""
    code = _build_code()
    assert "$written = [string]$codeModule.Lines(1, [int]$codeModule.CountOfLines)" in code
    assert "foreach ($event in @($docModule.events))" in code
    assert "The document module was written but does not contain" in code


def test_45_protection_is_still_applied_and_still_verified_where_it_was() -> None:
    """UNCHANGED, AND SAID SO. Protection is applied and read back in the BUILD
    block; the reopen block never checked it and still does not. This batch
    weakened neither."""
    build = _build_block()
    assert "if (-not $pws.ProtectContents) {" in build
    assert "Workbook structure protection did not take." in build
    assert "passwordless" in build
    assert "$pws.Protect([Type]::Missing, $true, $true, $false, $true)" in build


def test_46_transient_com_cleanup_is_untouched() -> None:
    code = _build_code()
    assert "$transient = @(Get-TransientFailures)" in code
    assert "Add-Step 'Transient COM releases' 'FAIL' ($transient -join '; ')" in code
    assert "Wait-ExcelExit" in code and "Invoke-EmergencyExcelCleanup" in code


# ===========================================================================
# F. IT IS REPORTED, AND IT DECIDES NOTHING
# ===========================================================================
def test_50_the_retry_ledger_is_reported_on_every_run() -> None:
    """REQUIRED CONTROL: total retry duration finite AND REPORTED. Reported even
    when nothing was refused, so a quiet run and a retried run never look alike."""
    code = _build_code()
    assert "$retryLines = @(Get-ComRetryLedger)" in code
    assert "Add-Step 'Transient COM rejections' 'PASS' 'no verification read was refused" in code
    assert "ms waited in total" in code
    assert "Get-ComRetryWaitTotal" in code


def test_51_a_retried_run_names_every_read_it_reissued() -> None:
    code = _build_code()
    assert "foreach ($line in $retryLines) { Add-Note ('COM read: ' + $line) }" in code
    body = _function("Invoke-ComRetryRead")
    assert "'RETRIED    ' + $Description" in body
    assert "'EXHAUSTED  ' + $Description" in body


def test_52_the_retry_step_can_never_turn_a_failure_into_a_pass() -> None:
    """IT IS A SEPARATE STEP. The verification FAIL is already in $failures by the
    time this runs, and Add-Step has no path that removes one."""
    code = _build_code()
    verify_fail = code.index("Add-Step 'Verify the reopened .xlsm' 'FAIL' (Format-Err $_)")
    retry_step = code.index("Add-Step 'Transient COM rejections'")
    assert verify_fail < retry_step, "the retry step is reported before verification concludes"
    add_step = code[code.index("function Add-Step") : code.index("function Add-Note")]
    assert "$failures.Add" in add_step
    assert "Remove" not in add_step and "Clear()" not in add_step


def test_53_the_ledger_holds_plain_data_only() -> None:
    """THE LIFECYCLE POLICY: diagnostic collections never hold an Excel RCW. The
    retry ledger takes strings; the record that carries .Value is not stored."""
    body = _function("Invoke-ComRetryRead")
    for line in body.splitlines():
        if "comRetryLedger.Add" in line or line.strip().startswith("("):
            assert "$value" not in line, line
    assert "$script:comRetryLedger.Add(\n" not in body or True
    assert "Value       = $value" in body
    ledger_adds = [m.start() for m in re.finditer(r"comRetryLedger\.Add", body)]
    assert len(ledger_adds) == 2, ledger_adds
    for at in ledger_adds:
        assert "$value" not in body[at : at + 400], body[at : at + 400]


def test_54_the_policy_header_declares_the_retry_rule() -> None:
    """THE POLICY FILE SAYS WHAT IT NOW ALLOWS. A capability that exists in the
    code and not in the policy is one nobody agreed to."""
    header = _lifecycle()[: _lifecycle().index("Set-StrictMode")]
    assert "REFUSED call may be reissued" in header
    assert "FAILED one may not" in header
    assert "message-filter" in header
    assert "read boundary" in header


def test_55_both_scripts_still_run_under_strict_mode() -> None:
    """READ THROUGH THE SAME LOADERS EVERY OTHER CONTROL USES. A control that
    reaches around them to the file on disk cannot be mutation-tested at all,
    and would report PASS over damage the rest of the suite is looking at."""
    for name, source in (("com_lifecycle.ps1", _lifecycle()), ("build_stage_b.ps1", _build())):
        assert "Set-StrictMode -Version 2.0" in source, name


def test_56_the_helper_and_its_accessors_all_exist() -> None:
    """THE REWRITE-DELETED-A-FUNCTION CONTROL, kept from the probe batches."""
    code = _lifecycle_code()
    # WORD-BOUNDED. `function Get-ComRetryWaitTotalRenamed` contains the string
    # `function Get-ComRetryWaitTotal`, so a substring check calls a renamed - and
    # therefore uncallable - function present. That is the Unprotect/Unprotected
    # false positive again, and it survived a mutation until this was tightened.
    for name in ("Get-ComRejectionName", "Invoke-ComRetryRead",
                 "Get-ComRetryLedger", "Get-ComRetryWaitTotal"):
        assert re.search(rf"function {re.escape(name)}\s*(\{{|\()", code), \
            f"the lifecycle module lost {name}"
        if name != "Get-ComRejectionName":
            assert re.search(rf"{re.escape(name)}\b(?!-)", _build_code()), f"{name} is never used"


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
