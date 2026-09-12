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
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent

BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
LIFECYCLE_PS1 = BOOTSTRAP / "com_lifecycle.ps1"
BUILD_PS1 = BOOTSTRAP / "build_stage_b.ps1"
GATE_PS1 = PCCM_ROOT / "tests" / "phase10_fixture_equivalence.ps1"
BENCHMARK_PS1 = BOOTSTRAP / "phase10_benchmark.ps1"
FLOW_PS1 = PCCM_ROOT / "tests" / "phase10_stage_b_build_ops_flow.ps1"
EVIDENCE_MD = PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md"
PWSH = "/opt/pwsh/pwsh"

# ---------------------------------------------------------------------------
# THE CLOSED STAGE-B BUILD OPERATION VOCABULARY
# ---------------------------------------------------------------------------
# RESTATED HERE, so a label quietly dropped from the script fails a control
# instead of agreeing with itself. The order is the order the build performs them.
BUILD_OPS = (
    "open.workbook",
    "saveas.xlsm",
    "worksheets.acquire",
    "codename.write",
    "vbproject.acquire",
    "vbcomponents.acquire",
    "vbcomponents.import",
    "thisworkbook.write",
    "button.add",
    "protection.apply",
    "workbook.save",
)

# The ONLY build-path reads this batch opted into the accepted retry, each one a
# plain property get, and each one a member the reopen verification block ALREADY
# reissues on the same class of object. Nothing here is a new judgement about what
# may be retried; it applies one already accepted, at the reads nearest the open.
DECLARED_BUILD_READS = {
    "FileFormat": "saveas.xlsm",
    "Worksheets": "worksheets.acquire",
    "VBProject": "vbproject.acquire",
    "VBComponents": "vbcomponents.acquire",
}

# Every non-idempotent build mutation. A retry around any of these would be a
# guess about whether Excel accepted it, which is the one thing the refusal
# contract does not license.
NEVER_RETRIED = (
    "Open", "SaveAs", "Save", "Import", "Remove", "AddFromString", "DeleteLines",
    "AddShape", "Protect", "Unprotect", "Delete", "Add", "Run", "Quit", "Close",
)

_FLOW: dict = {}

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


def _gate() -> str:
    return _src("gate", GATE_PS1)


def _evidence() -> str:
    return _src("evidence", EVIDENCE_MD)


def _at(commit: str, path: str) -> str:
    """One file as it stood at a commit. Repository-relative path."""
    key = f"git::{commit}::{path}"
    if key not in _MEMO:
        _MEMO[key] = subprocess.run(
            ["git", "show", f"{commit}:{path}"], cwd=REPO_ROOT, check=True,
            stdout=subprocess.PIPE, text=True).stdout
    return _MEMO[key]


def _ps_function(source: str, name: str) -> str:
    """One PowerShell function out of any source, by brace depth."""
    start = source.index(f"function {name} ")
    depth = 0
    seen = False
    for index in range(start, len(source)):
        if source[index] == "{":
            depth += 1
            seen = True
        elif source[index] == "}":
            depth -= 1
            if seen and depth == 0:
                return source[start : index + 1]
    raise AssertionError(f"{name} is not closed")


def _joined(text: str) -> str:
    """Backtick line continuations rejoined, so a check sees a whole call."""
    return re.sub(r"`\n\s*", " ", text)


def _flow() -> dict:
    """RUN the build-operation harness and group its tagged lines.

    "Which label was in flight", "how many attempts", "was the original error
    preserved" are COUNTS AND TYPES, not properties of text. They are observed.
    """
    if "flow" not in _FLOW:
        done = subprocess.run(
            [PWSH, "-NoProfile", "-File", str(FLOW_PS1),
             "-Build", str(BUILD_PS1), "-Lifecycle", str(LIFECYCLE_PS1)],
            capture_output=True, text=True, timeout=300)
        assert done.returncode == 0, done.stdout + done.stderr
        rows: dict = {"ops": [], "label": {}, "hres": {}, "line": {}, "retry": {},
                      "ledger": {}, "wrap": {}, "host": (), "unmet": []}
        for raw in done.stdout.splitlines():
            if raw.startswith(("PARSE|", "MISSING|")):
                rows["unmet"].append(raw)
            elif raw.startswith("OPS|"):
                _count, joined = raw[len("OPS|"):].split("|", 1)
                rows["ops"] = joined.split(",")
                assert int(_count) == len(rows["ops"]), raw
            elif raw.startswith("LABEL|"):
                case, asked, effect, outcome = raw[len("LABEL|"):].split("|", 3)
                rows["label"][case] = (asked, effect, outcome)
            elif raw.startswith("HRES|"):
                case, hexcode = raw[len("HRES|"):].split("|", 1)
                rows["hres"][case] = hexcode
            elif raw.startswith("LINE|"):
                case, line = raw[len("LINE|"):].split("|", 1)
                rows["line"][case] = line
            elif raw.startswith("RETRY|"):
                case, attempts, waited, calls, outcome, kind, hexcode = \
                    raw[len("RETRY|"):].split("|", 6)
                rows["retry"][case] = (int(attempts), int(waited), int(calls),
                                       outcome, kind, hexcode)
            elif raw.startswith("LEDGER|"):
                case, line = raw[len("LEDGER|"):].split("|", 1)
                rows["ledger"].setdefault(case, []).append(line)
            elif raw.startswith("HOST|"):
                prop, method = raw[len("HOST|"):].split("|", 1)
                rows["host"] = (prop, method)
            elif raw.startswith("WRAP|"):
                body = raw[len("WRAP|"):]
                if body.startswith("forwarded-arguments|"):
                    rows["wrap"]["forwarded-arguments"] = body[len("forwarded-arguments|"):]
                    continue
                head, tail = body.split("|label-after=", 1)
                after, value = tail.split("|value=", 1)
                case, label, attempts, lines, outcome, recorded = head.split("|", 5)
                rows["wrap"][case] = {
                    "label_at_forward": label, "attempts": int(attempts),
                    "lines": int(lines), "outcome": outcome, "recorded": recorded,
                    "label_after": after, "value": value,
                }
        _FLOW["flow"] = rows
    return _FLOW["flow"]


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


def test_33_the_build_block_retries_only_its_declared_property_gets() -> None:
    """DECLARED, NOT LOOSENED - AND THE OLD FORM HAD BECOME A FALSE REASSURANCE.

    This control used to read `"Invoke-ComRetryRead" not in _build_block()`. The
    build block now reissues four reads, and that assertion STILL PASSED, because
    the calls go through `Invoke-StageBBuildRead` and the only literal
    `Invoke-ComRetryRead` sits in the wrapper ABOVE where the block slice starts.
    A control that passes over the thing it forbids is worse than no control, so
    it is restated as the property it was always meant to hold:

      * the block may retry ONLY the four DECLARED members, each a plain property
        get, each already reissued by the accepted reopen verification;
      * each declared read must carry the operation label its member belongs to;
      * NO write may be retried by ANY path - both helper names are checked.
    """
    block = _joined(_build_block())
    calls = re.findall(r"Invoke-(?:ComRetryRead|StageBBuildRead)\b[^\n]*", block)
    assert calls, "the build block retries nothing at all, so nothing is declared"
    for call in calls:
        assert "Invoke-StageBBuildRead" in call, (
            f"the build block calls the general helper directly: {call}")
        member = re.search(r"-Member '(\w+)'", call)
        assert member, call
        name = member.group(1)
        assert name in DECLARED_BUILD_READS, (
            f"{name} is retried in the build block but is not a declared read")
        expected = DECLARED_BUILD_READS[name]
        assert f"-Operation '{expected}'" in call, (
            f"the {name} read does not carry the {expected} label: {call}")
        for banned in NEVER_RETRIED:
            assert f"-Member '{banned}'" not in call, f"{banned} is retried: {call}"
    # EVERY DECLARED READ MUST ACTUALLY BE THERE. A declaration for a read that no
    # longer exists is a stale exemption, and it would let a later edit drop the
    # retry without a single control noticing.
    for member, label in DECLARED_BUILD_READS.items():
        assert f"-Member '{member}'" in block, f"the declared read {member} is gone"
        assert f"-Operation '{label}'" in block, f"the declared label {label} is gone"
    # AND THE REOPEN PATH IS WHY EACH ONE IS ALLOWED. If verification stops
    # reissuing a member, the build has no precedent left for reissuing it either.
    verify = _joined(_verify_block())
    for member in DECLARED_BUILD_READS:
        assert f"-Member '{member}'" in verify, (
            f"{member} is retried in the build with no precedent in verification")


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
    when nothing was refused, so a quiet run and a retried run never look alike.

    DECLARED CHANGE. This used to read the whole ledger in one place -
    `$retryLines = @(Get-ComRetryLedger)` - because verification was its only
    reader. The build path now reissues four declared reads, and left alone that
    one statement would have reported the BUILD's refusals inside the step that
    says "no VERIFICATION read was refused". The statement is sliced rather than
    weakened: each phase reports ITS OWN refusals, and the accepted sentence stops
    being able to make a false claim.
    """
    code = _build_code()
    # Both baselines, taken in order: build first, then verification.
    build_base = code.index("$buildRetryBase     = @(Get-ComRetryLedger).Count")
    verify_base = code.index("$verifyRetryBase     = @(Get-ComRetryLedger).Count")
    assert build_base < verify_base, "the verification baseline is taken before the build's"
    assert "$buildRetryWaitBase = Get-ComRetryWaitTotal" in code
    assert "$verifyRetryWaitBase = Get-ComRetryWaitTotal" in code
    # Each phase reads only its own slice.
    assert "$buildLedger     = @(@(Get-ComRetryLedger) | Select-Object -Skip $buildRetryBase)" in code
    assert "$retryLines = @(@(Get-ComRetryLedger) | Select-Object -Skip $verifyRetryBase)" in code
    # And each phase reports on every run, refused or not.
    assert "Add-Step 'Transient COM rejections' 'PASS' 'no verification read was refused" in code
    assert "Add-Step 'Transient COM rejections (build)' 'PASS' 'COMREJECT|build|none|attempts=0|waited=0'" in code
    # EACH STEP CARRIES ITS OWN WAIT TOTAL, NAMED. A bare "ms waited in total"
    # anywhere in the file used to satisfy this, and once the build gained a step
    # of its own that let the VERIFICATION step stop reporting its wait entirely.
    assert ('("{0} build read(s) were refused and reissued; {1} ms waited in total"'
            " -f $buildLedger.Count, $buildRetryWait)") in code
    assert ('("{0} read(s) were refused and reissued; {1} ms waited in total"'
            " -f $retryLines.Count, ((Get-ComRetryWaitTotal) - $verifyRetryWaitBase))") in code


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


# ===========================================================================
# G. WHICH BUILD CALL WAS REFUSED - THE INSTRUMENTATION
# ===========================================================================
# TWO CONSECUTIVE EQUIVALENCE RUNS DIED IN THE BUILD BLOCK AND THE TRANSCRIPT
# COULD NAME ONLY THE REGION. `Add-Step 'Stage-B build' 'FAIL' (Format-Err $_)`
# reported one HRESULT and the name of a try/catch wrapping eleven distinct COM
# operations. The exact rejected call is STILL UNKNOWN: these controls exist so
# the NEXT Windows run names it, not so this batch can claim to have.
def test_60_every_declared_build_operation_maps_to_a_real_call_site() -> None:
    """REQUIRED CONTROL. Eleven labels, and each one is set somewhere the build
    actually calls COM - either explicitly, or as the -Operation of a declared
    read. A label with no call site is a vocabulary entry that can never appear
    in a diagnostic, which is the same as not having it."""
    block = _build_block()
    for label in BUILD_OPS:
        explicit = f"Set-StageBBuildOp '{label}'" in block
        viaread = f"-Operation '{label}'" in block
        assert explicit or viaread, f"the label {label} is never set in the build block"
    # AND NO LABEL OUTSIDE THE VOCABULARY. A twelfth one set somewhere would be
    # refused at runtime by Set-StageBBuildOp, and it must be refused here too so
    # the failure is a control rather than a Windows run.
    used = set(re.findall(r"Set-StageBBuildOp '([^']+)'", block))
    used |= set(re.findall(r"-Operation '([^']+)'", block))
    assert used <= set(BUILD_OPS), sorted(used - set(BUILD_OPS))
    assert used == set(BUILD_OPS), sorted(set(BUILD_OPS) - used)


def test_61_the_vocabulary_is_closed_and_a_typo_cannot_become_a_label() -> None:
    """EXECUTED. A misspelt label must fail where it is written. `vbproject.aquire`
    is refused, and the label in effect is left as it was - because a diagnostic
    that names the PREVIOUS operation is worse than one that names none."""
    flow = _flow()
    assert not flow["unmet"], flow["unmet"]
    assert tuple(flow["ops"]) == BUILD_OPS, flow["ops"]
    for label in BUILD_OPS:
        asked, effect, outcome = flow["label"][label]
        assert outcome == "accepted", (label, outcome)
        assert effect == label, (label, effect)
    _asked, effect, outcome = flow["label"]["typo"]
    assert outcome == "REFUSED", flow["label"]["typo"]
    assert effect == "workbook.save", "a refused label became the label in effect"


def test_62_the_build_failure_names_the_operation_that_was_in_flight() -> None:
    """REQUIRED CONTROL, AND THE WHOLE POINT. The catch reports the label, not just
    the region, and it emits one classified line beside it."""
    code = _build_code()
    at = code.index("Add-Step 'Stage-B build' 'FAIL'")
    region = code[at - 200 : at + 700]
    assert "$failedOp = Get-StageBBuildOp" in region, region
    assert "'operation=' + $failedOp" in region, region
    assert "New-StageBRejectionLine -Operation $failedOp -ErrorRecord $_" in region, region


def test_63_a_refused_call_and_an_accepted_one_are_never_reported_alike() -> None:
    """EXECUTED, AND THE DISTINCTION THE WHOLE ARCHITECTURE RESTS ON. Only the two
    message-filter results may be reported as a refusal; an HRESULT Excel returned
    from a call it ACCEPTED describes something that actually happened, and calling
    that a rejection would licence reissuing it."""
    flow = _flow()
    assert flow["hres"]["rejected"] == "0x80010001"
    assert flow["hres"]["retrylater"] == "0x8001010a"
    assert flow["hres"]["accepted"] == "0x800a03ec"
    assert flow["hres"]["vba-ignore"] == "0x800ac472"
    assert flow["hres"]["not-com"] == "<empty>"
    assert flow["hres"]["null"] == "<empty>"
    # A COMException reached through an InnerException chain is still a COM
    # failure, and still carries its code.
    assert flow["hres"]["wrapped"] == "0x80010001"
    for case in ("rejected", "retrylater", "wrapped"):
        assert flow["line"][case].startswith("COMREJECT|build|"), flow["line"][case]
        assert "refused before it ran" in flow["line"][case]
    for case in ("accepted", "vba-ignore", "not-com"):
        assert flow["line"][case].startswith("COMFAIL|build|"), flow["line"][case]
        assert "NOT refused" in flow["line"][case]
    assert "hresult=none" in flow["line"]["not-com"], flow["line"]["not-com"]


def test_64_an_error_excel_accepted_is_never_reissued() -> None:
    """EXECUTED, NOT ASSERTED. The fake's method is called exactly ONCE for an
    accepted-and-failed HRESULT, for VBA_E_IGNORE and for an error that is not a
    COM failure at all. One call means no retry happened."""
    flow = _flow()
    for case, expected_hex in (("accepted-and-failed", "0x800a03ec"),
                               ("vba-ignore", "0x800ac472"),
                               ("not-com", "<empty>")):
        attempts, waited, calls, outcome, kind, hexcode = flow["retry"][case]
        assert outcome == "RAISED", (case, outcome)
        assert calls == 1, f"{case} was reissued {calls} times"
        assert waited == -1, (case, waited)
        assert hexcode == expected_hex, (case, hexcode)
        assert case not in flow["ledger"], f"{case} was written to the retry ledger"


def test_65_a_refused_call_is_reissued_and_the_count_is_observed() -> None:
    """EXECUTED. Refused twice then answered: THREE calls, three attempts, and one
    RETRIED ledger line naming the rejection. Both message-filter codes count."""
    flow = _flow()
    attempts, waited, calls, outcome, _kind, _hex = flow["retry"]["answers-first-call"]
    assert (attempts, calls, outcome) == (1, 1, "answered"), flow["retry"]["answers-first-call"]
    assert "answers-first-call" not in flow["ledger"], "an unrefused read was logged"
    attempts, waited, calls, outcome, _kind, _hex = flow["retry"]["refused-twice"]
    assert (attempts, calls, outcome) == (3, 3, "answered"), flow["retry"]["refused-twice"]
    assert waited > 0, "the wait did not accumulate"
    assert any("RETRIED" in line and "RPC_E_CALL_REJECTED" in line
               for line in flow["ledger"]["refused-twice"]), flow["ledger"]["refused-twice"]
    attempts, _waited, calls, outcome, _kind, _hex = flow["retry"]["retrylater-once"]
    assert (attempts, calls, outcome) == (2, 2, "answered"), flow["retry"]["retrylater-once"]
    assert any("RPC_E_SERVERCALL_RETRYLATER" in line
               for line in flow["ledger"]["retrylater-once"])


def test_66_the_attempt_limit_actually_stops_the_reissuing() -> None:
    """EXECUTED. With the limit at three, an always-refused read is called exactly
    THREE times. A bound that is compared but never binds is not a bound."""
    flow = _flow()
    _attempts, _waited, calls, outcome, _kind, _hex = flow["retry"]["always-refused"]
    assert outcome == "RAISED", flow["retry"]["always-refused"]
    assert calls == 3, f"the attempt limit of 3 allowed {calls} calls"
    assert any("attempt limit of 3 reached" in line
               for line in flow["ledger"]["always-refused"]), flow["ledger"]["always-refused"]


def test_67_the_total_wait_stops_it_before_the_attempt_limit_does() -> None:
    """EXECUTED, AND A SEPARATE BOUND FROM THE LAST ONE. With fifty attempts
    allowed and a five-millisecond budget, the budget is what ends it - so the
    budget is not decoration behind a small attempt limit."""
    flow = _flow()
    _attempts, _waited, calls, outcome, _kind, _hex = flow["retry"]["budget-bound"]
    assert outcome == "RAISED", flow["retry"]["budget-bound"]
    assert calls < 50, f"the retry budget allowed all {calls} attempts"
    assert any("budget of 5 ms would be exceeded" in line
               for line in flow["ledger"]["budget-bound"]), flow["ledger"]["budget-bound"]


def test_68_an_exhausted_retry_rethrows_the_original_failure() -> None:
    """REQUIRED CONTROL, EXECUTED. The caller must fail exactly as it would have
    before the helper existed: the ORIGINAL rejection, with its HRESULT, and an
    EXHAUSTED line on the record. A retry that swallows the failure it could not
    resolve reports success nobody observed."""
    flow = _flow()
    _attempts, _waited, _calls, outcome, kind, hexcode = flow["retry"]["always-refused"]
    assert outcome == "RAISED"
    assert hexcode == "0x80010001", f"the original rejection was replaced: {hexcode}"
    assert "COMException" in kind or "MethodInvocation" in kind, kind
    assert any(line.startswith("EXHAUSTED") for line in flow["ledger"]["always-refused"])
    # The wrapper does not intercept it either, and the label survives.
    wrap = flow["wrap"]["exhausted"]
    assert wrap["outcome"] == "RAISED", wrap
    assert wrap["lines"] == 0, "an exhausted read was reported as answered"
    assert wrap["label_after"] == "vbproject.acquire", wrap


def test_69_the_operation_label_is_set_before_the_call_is_forwarded() -> None:
    """EXECUTED. If the wrapper set the label AFTER forwarding, a read refused on
    its first attempt would carry the PREVIOUS operation's name - and the one
    thing this batch exists to produce is the RIGHT name."""
    flow = _flow()
    for case in ("answered-first-attempt", "answered-on-third", "exhausted",
                 "answered-with-nothing"):
        wrap = flow["wrap"][case]
        assert wrap["label_at_forward"] == "vbproject.acquire", (case, wrap)


def test_70_a_reissued_read_is_named_and_a_quiet_one_is_not_invented() -> None:
    """EXECUTED. One concise line per read that actually had to be reissued,
    carrying the operation, the attempts and the wait; NO line when the read
    answered first time. A diagnostic that fires either way says nothing."""
    flow = _flow()
    assert flow["wrap"]["answered-first-attempt"]["lines"] == 0
    third = flow["wrap"]["answered-on-third"]
    assert third["lines"] == 1, third
    assert third["recorded"].startswith("COMREJECT|build|vbproject.acquire|"), third
    assert "attempts=3" in third["recorded"] and "waited=750" in third["recorded"], third
    assert "RPC_E_CALL_REJECTED" in third["recorded"], third


def test_71_the_wrapper_can_express_no_write_and_not_even_an_item_lookup() -> None:
    """THE GUARANTEE IS THE SHAPE OF THE API, and it is NARROWER than the helper
    it forwards to: no scriptblock, and no -Key either, so a plain property get is
    the only thing that fits through it."""
    body = _ps_function(_build_code(), "Invoke-StageBBuildRead")
    params = body[body.index("param(") : body.index(")", body.index("param("))]
    for banned in ("ScriptBlock", "scriptblock", "$Action", "$Script", "$Body",
                   "$Key", "$Arguments", "$Value"):
        assert banned not in params, f"the wrapper takes {banned}"
    assert "$Target" in params and "$Member" in params
    assert "$Operation" in params and "$Description" in params
    # It forwards those three and nothing else.
    forward = [line for line in _joined(body).splitlines() if "Invoke-ComRetryRead" in line]
    assert len(forward) == 1, forward
    assert "-Key" not in forward[0], forward[0]
    for banned in NEVER_RETRIED:
        assert f"'{banned}'" not in forward[0], forward[0]
    # EXECUTED: the forwarded call really carried no key.
    assert "key=False" in _flow()["wrap"]["forwarded-arguments"], \
        _flow()["wrap"]["forwarded-arguments"]


def test_72_the_accepted_helper_itself_was_not_broadened() -> None:
    """REQUIRED CONTROL. This batch adds CALL SITES, not capability. The helper and
    its classifier are byte-identical to the revision that closed the reserved-row
    batch, so 'narrow read retry' cannot have quietly become something else."""
    now = _lifecycle()
    then = _at("1e0edb2", "pccm/bootstrap/windows/com_lifecycle.ps1")
    assert now == then, "com_lifecycle.ps1 changed in a batch that may not change it"


def test_73_nothing_that_mutates_the_workbook_is_retried() -> None:
    """REQUIRED CONTROL, NAMED PER OPERATION. SaveAs, the module import, the
    ThisWorkbook write, the button creation, the protection call and the final
    Save are each labelled so a refusal NAMES them - and none of them is inside a
    retry. The label is the deliverable; reissuing them is a separate decision
    nobody has taken."""
    joined = _joined(_build_block())
    mutations = {
        "saveas.xlsm": "$wb.SaveAs($stageBPath",
        "vbcomponents.import": "$vbcomps.Import($file)",
        "thisworkbook.write": "$codeModule.AddFromString($docText)",
        "button.add": "$shapes.AddShape(5,",
        "protection.apply": "$pws.Protect([Type]::Missing",
        "workbook.save": "$wb.Save()",
    }
    for label, call in mutations.items():
        assert call in joined, f"the {label} call site is gone: {call}"
        # The label is set for it...
        assert f"Set-StageBBuildOp '{label}'" in joined, f"{label} is never labelled"
        # ...and the call itself is not routed through either retry helper.
        for line in joined.splitlines():
            if call in line:
                assert "Invoke-ComRetryRead" not in line, line
                assert "Invoke-StageBBuildRead" not in line, line
    # AND THE LABEL PRECEDES THE MUTATION IT DESCRIBES. A label set after the call
    # would name it only for whatever failed next.
    for label, call in mutations.items():
        assert joined.index(f"Set-StageBBuildOp '{label}'") < joined.index(call), label


def test_74_every_other_structural_mutation_is_also_left_alone() -> None:
    """THE LIST IS DERIVED FROM THE BLOCK, NOT REMEMBERED. Every method call in the
    build block whose name is a mutation must sit on a line with no retry on it."""
    joined = _joined(_build_block())
    for line in joined.splitlines():
        if "Invoke-ComRetryRead" not in line and "Invoke-StageBBuildRead" not in line:
            continue
        for banned in NEVER_RETRIED:
            assert f".{banned}(" not in line, line
            assert f"-Member '{banned}'" not in line, line


def test_75_no_blanket_sleep_and_no_readiness_gate_was_added() -> None:
    """NOT AUTHORISED IN THIS BATCH, AND REFUSED HERE RATHER THAN REMEMBERED. We do
    not yet know that workbook readiness is the failing condition, and a poll
    inserted before the evidence exists could make the symptom disappear without
    ever proving its cause."""
    code = _build_code()
    assert "Start-Sleep" not in code, "the bootstrap sleeps outside the retry"
    for banned in ("Wait-ExcelReady", "Test-ExcelReady", "Wait-WorkbookReady",
                   "readiness", "-Member 'Ready'", "Start-Process", "Get-Random"):
        assert banned not in code, f"a readiness mechanism appeared: {banned}"
    # THE STRUCTURAL CHECK, NOT A WORD LIST. A post-Open readiness poll would live
    # between the Open and the SaveAs, and there is nothing between them.
    joined = _joined(_build_block())
    span = joined[joined.index("$wb = $workbooks.Open($stageAPath)")
                  : joined.index("$wb.SaveAs($stageBPath")]
    for banned in ("Invoke-ComRetryRead", "Invoke-StageBBuildRead", "Start-Sleep",
                   "while", "do {", "for ("):
        assert banned not in span, f"something was inserted after the Open: {span!r}"


def test_76_no_inter_pass_drain_was_added_to_the_equivalence_gate() -> None:
    """NOT AUTHORISED EITHER. Both Excel instances shut down naturally in the
    observed runs, so nothing links the second-pass failure to a lifecycle
    overlap. A pause between the passes would be a remedy for a cause nobody has
    demonstrated - and it would hide the one we are trying to find."""
    gate = _gate()
    for banned in ("Start-Sleep", "Get-Process", "Stop-Process", "Wait-Process",
                   "WaitForExit", "Wait-ExcelExit", "drain", "quiet period",
                   "System.Threading.Thread]::Sleep"):
        assert banned not in gate, f"the gate waits between passes: {banned}"


def test_77_the_two_passes_are_still_two_bundles_and_two_excel_sessions() -> None:
    """REQUIRED CONTROL: fixture isolation is not the price of a diagnostic. The
    bundle machinery is byte-identical to the accepted revision, and each pass
    still starts its own Excel."""
    gate = _gate()
    then = _at("1e0edb2", "pccm/tests/phase10_fixture_equivalence.ps1")
    for name in ("Get-BundleArtifacts", "New-EquivalenceBundle", "Test-BundleIdentity"):
        assert _ps_function(gate, name) == _ps_function(then, name), \
            f"{name} changed in a batch that may not change it"
    pass_fn = _ps_function(gate, "Invoke-EquivalencePass")
    assert "New-Object -ComObject Excel.Application" in pass_fn
    assert "$tempRoot = [string]$Bundle.Root" in pass_fn
    assert "& $bootstrap -BuildDir $tempRoot -Force" in pass_fn


def test_78_the_gate_refuses_a_bootstrap_that_did_not_succeed() -> None:
    """THE DEFECT THIS BATCH FOUND WHILE READING THE GATE. The pass checked only
    that a Stage-B workbook EXISTED. A build refused after its SaveAs leaves the
    .xlsm on disk with no modules, no buttons and no protection, and the pass
    would have opened it and reported a fixture result against a half-built
    workbook. The exit code is now checked, in the vocabulary the gate already
    has."""
    pass_fn = _ps_function(_gate(), "Invoke-EquivalencePass")
    assert "$bootstrapExit = $LASTEXITCODE" in pass_fn
    exitcheck = pass_fn.index("if ($bootstrapExit -ne 0) {")
    pathcheck = pass_fn.index("if (-not (Test-Path -LiteralPath $stageB)) {")
    assert exitcheck < pathcheck, "the file check runs before the exit-code check"
    assert "BOOTSTRAP: the Stage-B bootstrap for the ' + $Mode + ' pass exited " in pass_fn
    # THE EXISTING VOCABULARY, NOT A NEW ONE.
    assert pass_fn.count("BOOTSTRAP:") >= 2
    assert "half-built" in pass_fn


def test_79_a_clean_build_still_reports_that_nothing_was_refused() -> None:
    """REQUIRED CONTROL. Silence and success looked identical in all four runs, so
    the no-rejection path prints too - one concise aggregate line."""
    code = _build_code()
    assert "'COMREJECT|build|none|attempts=0|waited=0'" in code
    assert "if ($buildLedger.Count -eq 0 -and $buildRejections.Count -eq 0) {" in code
    # A PASS STEP THAT DECIDES NOTHING. It is reported after the build concluded,
    # and Add-Step has no path that removes a failure.
    build_fail = code.index("Add-Step 'Stage-B build' 'FAIL'")
    aggregate = code.index("Add-Step 'Transient COM rejections (build)'")
    assert build_fail < aggregate, "the aggregate is reported before the build concludes"
    assert "'FAIL'" not in code[aggregate : aggregate + 600]


def test_80_shutdown_still_runs_after_a_refused_build() -> None:
    """A HARNESS THAT LEAVES AN EXCEL PROCESS OPEN IS WORSE THAN ONE THAT FAILS.
    The diagnostic sits between the catch and the shutdown ledger and cannot exit,
    return or throw past it."""
    code = _build_code()
    aggregate = code.index("Add-Step 'Transient COM rejections (build)'")
    shutdown = code.index("$rel1 = New-ReleaseLedger 'build instance'")
    assert aggregate < shutdown, "the diagnostic is reported after shutdown"
    between = code[code.index("Add-Step 'Stage-B build' 'FAIL'") : shutdown]
    for banned in ("exit ", "return", "throw"):
        assert banned not in between, f"{banned} can skip the shutdown: {between!r}"
    assert "Wait-ExcelExit -Identity $buildExcelIdentity" in code


def test_81_a_diagnostic_never_replaces_the_failure_it_describes() -> None:
    """THE FAILURE IS THE FINDING. If classifying the error throws, the build
    failure is already recorded and stays recorded; the classification degrades to
    an unclassified line instead of becoming an unhandled error that loses the
    transcript."""
    code = _build_code()
    at = code.index("Add-Step 'Stage-B build' 'FAIL'")
    region = code[at : at + 700]
    assert "try   { Add-Note (New-StageBRejectionLine" in region, region
    assert "catch { Add-Note ('COMFAIL|build|' + $failedOp + '|hresult=unclassified" in region, region
    assert region.index("Add-Step 'Stage-B build' 'FAIL'") < region.index("New-StageBRejectionLine")


def test_82_nothing_is_answered_with_nothing() -> None:
    """THIS HOST DISCARDS AN EXCEPTION THROWN BY A PROPERTY GETTER - observed, not
    assumed, and printed by the harness as its HOST line. A read that neither
    raised nor answered must still name the operation it happened at, rather than
    surfacing three statements later as a null reference nobody can trace."""
    flow = _flow()
    prop, method = flow["host"]
    assert prop == "swallowed", prop
    assert method.startswith("raised"), method
    body = _ps_function(_build_code(), "Invoke-StageBBuildRead")
    assert "if ($null -eq $record.Value) {" in body
    assert "answered with nothing at " in body
    wrap = flow["wrap"]["answered-with-nothing"]
    assert wrap["outcome"] == "RAISED", wrap
    assert wrap["label_after"] == "vbproject.acquire", wrap


def test_83_the_snapshot_and_the_reserved_row_correction_are_untouched() -> None:
    """SEMANTIC FREEZE. This batch diagnoses a COM rejection. It may not move the
    comparison or the fixture, so both are byte-identical to their authorities."""
    snapshot_now = _ps_function(_gate(), "Get-EquivalenceSnapshot")
    snapshot_then = _ps_function(_at("99cb472", "pccm/tests/phase10_fixture_equivalence.ps1"),
                                 "Get-EquivalenceSnapshot")
    assert snapshot_now == snapshot_then, "Get-EquivalenceSnapshot moved"
    assert BENCHMARK_PS1.read_text(encoding="utf-8") == \
        _at("1e0edb2", "pccm/bootstrap/windows/phase10_benchmark.ps1"), \
        "the benchmark runner changed in a batch that may not change it"


def test_84_production_vba_is_byte_identical() -> None:
    """NO PRODUCTION VBA CHANGE IS AUTHORISED. Asked of git rather than of a list,
    so a file nobody thought to name is covered too."""
    done = subprocess.run(["git", "diff", "--name-only", "1e0edb2", "--", "pccm/src/vba"],
                          cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE, text=True)
    assert done.stdout.strip() == "", done.stdout


def test_85_the_historical_runs_do_not_claim_an_exact_rejected_operation() -> None:
    """THE EVIDENCE RULE. Static work cannot say which call runs 3 and 4 were
    refused on, because the log that would have said it did not exist yet. The
    record states the region, states that the operation was NOT identified, and
    names VBProject only as a candidate."""
    text = _evidence()
    at = text.index("## Equivalence runs 3 and 4")
    after = text.find("\n## ", at + 10)
    section = text[at:] if after == -1 else text[at:after]
    assert "INVALID / NOT EVALUATED" in section
    assert "did not identify the exact rejected COM operation" in section, section[:400]
    assert "candidate" in section
    for overclaim in ("the rejected call was", "the failing call was",
                      "VBProject was the rejected", "identified as vbproject"):
        assert overclaim not in section, f"the record overclaims: {overclaim}"
    # AND IT IS NOT A DIFFER, NOT A BULK FAILURE, NOT EQUIVALENCE.
    assert "must not be read as a fixture DIFFER" in section
    lowered = section.lower()
    assert "no comparison" in lowered or "no semantic comparison" in lowered
    assert "never began" in lowered or "never reached" in lowered


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
