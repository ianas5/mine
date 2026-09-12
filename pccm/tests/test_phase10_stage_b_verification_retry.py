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

# Reads that exist ONLY to settle a refused SaveAs. They have no verification
# precedent - the reopen block never reads FullName - and they do not need one:
# they are pure property gets whose whole purpose is to observe whether a
# non-idempotent call happened. Declared separately so the precedent rule that
# governs the four build reads above is not quietly widened to mean anything.
DECLARED_POSTCONDITION_READS = {
    "FullName": "saveas.presave.fullname",
    "FileFormat": "saveas.presave.fileformat",
}

# The closed SUB-operation vocabulary. Run 5 reported a failed PRE-SAVE read as
# operation=saveas.xlsm, which reads as though the save had been attempted; it had
# not. Restated here so a sub-operation quietly dropped fails a control.
BUILD_SUB_OPS = (
    "saveas.presave.fullname",
    "saveas.presave.fileformat",
    "saveas.presave.target",
    "saveas.call",
    "saveas.post.fullname",
    "saveas.post.fileformat",
    "saveas.post.target",
)

# The build-only functions that carry the SaveAs settlement. Named, so a retry
# hidden in one of them is inside what the controls look at rather than above it.
SAVEAS_HELPERS = (
    "Get-StageBComparablePath",
    "Get-StageBSaveAsPostcondition",
    "Invoke-StageBSaveAs",
    "New-StageBSaveAsResult",
)

# Every non-idempotent build mutation. A retry around any of these would be a
# guess about whether Excel accepted it, which is the one thing the refusal
# contract does not license. SaveAs is the single DECLARED exception, and its
# recovery is gated on an observation rather than on the refusal contract - see
# test_86 onward.
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
                      "ledger": {}, "wrap": {}, "host": (), "unmet": [],
                      "save": {}, "savenote": {}, "savestate": {}, "savepath": {},
                      "read": {}, "telemetry": {}, "subops": [], "step": {}}
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
            elif raw.startswith("READ|"):
                case, outcome, detail, kind, lines = raw[len("READ|"):].split("|", 4)
                rows["read"][case] = {"outcome": outcome, "value": detail,
                                      "detail": detail, "type": kind,
                                      "lines": int(lines.split("=", 1)[1])}
            elif raw.startswith("TELEMETRY|"):
                case, rest = raw[len("TELEMETRY|"):].split("|", 1)
                row = {"detail": rest, "outcome": "", "lines": -1, "calls": -1}
                first = rest.split("|", 1)[0]
                if first in ("ok", "REFUSED"):
                    row["outcome"] = first
                for part in rest.split("|"):
                    if part.startswith("lines="):
                        row["lines"] = int(part.split("=", 1)[1])
                    elif part.startswith("calls="):
                        row["calls"] = int(part.split("=", 1)[1])
                rows["telemetry"][case] = row
            elif raw.startswith("SUBOPS|"):
                count, joined = raw[len("SUBOPS|"):].split("|", 1)
                rows["subops"] = joined.split(",")
                assert int(count) == len(rows["subops"]), raw
            elif raw.startswith("STEP|"):
                parts = raw[len("STEP|"):].split("|")
                rows["step"][parts[0]] = parts[1:]
            elif raw.startswith("SAVE|"):
                case, calls, attempts, waited, fmt, outcome, state = \
                    raw[len("SAVE|"):].split("|", 6)
                rows["save"][case] = {
                    "calls": int(calls), "attempts": int(attempts),
                    "waited": int(waited), "format": int(fmt),
                    "outcome": outcome, "state": state,
                }
            elif raw.startswith("SAVENOTE|"):
                case, line = raw[len("SAVENOTE|"):].split("|", 1)
                rows["savenote"].setdefault(case, []).append(line)
            elif raw.startswith("SAVESTATE|"):
                case, state, detail = raw[len("SAVESTATE|"):].split("|", 2)
                rows["savestate"][case] = (state, detail)
            elif raw.startswith("SAVEPATH|"):
                case, same = raw[len("SAVEPATH|"):].split("|", 1)
                rows["savepath"][case] = (same == "True")
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
    a bounded retry loop, or a fixed one before verification, is refused here.

    DECLARED CHANGE. This used to read `"Start-Sleep" not in _build_code()`, and
    that became false when Windows named SaveAs as the rejected call and the save
    gained a bounded recovery. The rule is unchanged and is now stated where it
    actually lives: EVERY sleep in either script is inside a bounded retry loop,
    and the bootstrap's one sleep is additionally reachable only after a
    postcondition inspection PROVED the save did not happen.
    """
    lifecycle = _lifecycle_code()
    sleeps = [m.start() for m in re.finditer(r"Start-Sleep", lifecycle)]
    retry = _function("Invoke-ComRetryRead")
    wait_exit = _function("Wait-ExcelExit")
    for at in sleeps:
        window = lifecycle[max(0, at - 400) : at + 60]
        assert any(window in block or lifecycle[at : at + 40] in block
                   for block in (retry, wait_exit)), lifecycle[at - 200 : at + 60]
    code = _build_code()
    saveas = _ps_function(code, "Invoke-StageBSaveAs")
    build_sleeps = [m.start() for m in re.finditer(r"Start-Sleep", code)]
    assert len(build_sleeps) == 1, f"the bootstrap sleeps in {len(build_sleeps)} places"
    assert "Start-Sleep" in saveas, "the bootstrap's sleep is not in the SaveAs retry"
    # AND IT IS GATED ON THE OBSERVATION, not merely on the refusal.
    gate = saveas.index("if ($state.State -ne 'not-executed') {")
    sleep_at = saveas.index("Start-Sleep")
    assert gate < sleep_at, "the sleep is reached before the state is established"


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
    """DECLARED, NOT LOOSENED - AND RESTATED TWICE NOW, EACH TIME BY A WINDOWS RUN.

    It began as `"Invoke-ComRetryRead" not in _build_block()`. That became a false
    reassurance when reads went through a wrapper the slice could not see. Run 5
    then proved the wrapper itself was the defect - its one execution on Windows
    produced no answer - so the wrapper is gone and every read is the form the
    reopen verification has executed on Windows: the accepted helper, called
    directly. What this control asserts is unchanged:

      * only DECLARED members are reissued, each a plain property get;
      * each read carries the operation label its member belongs to;
      * no write is reissued by any path;
      * and there is NO read wrapper to hide a read behind again.
    """
    code = _build_code()
    scope = _joined(_build_block()) + "\n" + "\n".join(
        _joined(_ps_function(code, name)) for name in SAVEAS_HELPERS)
    allowed = dict(DECLARED_BUILD_READS)
    allowed.update(DECLARED_POSTCONDITION_READS)
    calls = [line for line in re.findall(r"Invoke-ComRetryRead\b[^\n]*", scope)
             if "-Target " in line]
    assert calls, "the build performs no retried read at all"
    for call in calls:
        member = re.search(r"-Member '(\w+)'", call)
        assert member, call
        name = member.group(1)
        assert name in allowed, f"{name} is reissued but is not a declared read"
        for banned in NEVER_RETRIED:
            assert f"-Member '{banned}'" not in call, f"{banned} is retried: {call}"
    # EVERY DECLARED READ MUST ACTUALLY BE THERE - a declaration for a read that no
    # longer exists is a stale exemption.
    for member in allowed:
        assert f"-Member '{member}'" in scope, f"the declared read {member} is gone"
    # AND THE REOPEN PATH IS WHY EACH BUILD READ IS ALLOWED.
    verify = _joined(_verify_block())
    for member in DECLARED_BUILD_READS:
        assert f"-Member '{member}'" in verify, (
            f"{member} is retried in the build with no precedent in verification")
    # THE WRAPPER IS GONE, AND STAYS GONE. Run 5 is the reason.
    assert "function Invoke-StageBBuildRead" not in code, (
        "the read wrapper is back; run 5 is what it does on Windows")
    assert _strip(_build()).count("Invoke-StageBBuildRead") == 0



# The complete set of members the retry helper is allowed to name. Read-only,
# every one of them, and NAMED -- a proximity check around the write call sites
# lets a retried write slip past simply by sitting on its own line.
#
# FullName was added when SaveAs gained a postcondition: proving whether a refused
# save happened means asking the workbook which file it is now bound to. It is a
# property get like every other name here, and it moves nothing.
RETRYABLE_MEMBERS = {"FileFormat", "Worksheets", "VBProject", "VBComponents",
                     "Count", "CodeName", "Shapes", "OnAction", "Name", "Item",
                     "FullName"}


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
    """REQUIRED CONTROL, AND RUN 5 SHARPENED IT. The catch reports the PRECISE label
    - the sub-operation when one is in flight - because run 5 reported a failed
    pre-save READ as operation=saveas.xlsm, which reads as though the save had been
    attempted. It had not been."""
    code = _build_code()
    at = code.index("Add-Step 'Stage-B build' 'FAIL'")
    region = code[at - 300 : at + 700]
    assert "$failedOp = Get-StageBBuildLabel" in region, region
    assert "'operation=' + $failedOp" in region, region
    assert "New-StageBRejectionLine -Operation $failedOp -ErrorRecord $_" in region, region
    # AND THE LABEL PREFERS THE SUB-OPERATION.
    label = _ps_function(code, "Get-StageBBuildLabel")
    assert "if ($script:StageBBuildSubOp -ne '') { return [string]$script:StageBBuildSubOp }" in label
    assert "return [string]$script:StageBBuildOp" in label



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
    # AND THE BUILD'S OWN READ PATH DOES NOT INTERCEPT IT EITHER: the original
    # HRESULT escapes, and the out-of-band telemetry records nothing, because there
    # was no answer to record.
    row = flow["telemetry"]["exhausted"]
    assert row["outcome"] == "REFUSED", row
    assert "0x80010001" in row["detail"], row
    assert row["lines"] == 0, "an exhausted read was reported as answered"
    assert row["calls"] == 3, row



def test_69_the_operation_label_is_set_before_the_call_is_forwarded() -> None:
    """THE LABEL IS SET BEFORE THE READ, NOT AFTER IT. A read refused on its first
    attempt would otherwise carry the PREVIOUS step's name - a diagnostic that is
    confidently wrong, which is worse than none. Checked in source at every read,
    and observed on the reissued one."""
    code = _build_code()
    scope = _joined(_build_block()) + "\n" + "\n".join(
        _joined(_ps_function(code, name)) for name in SAVEAS_HELPERS)
    # IMMEDIATELY BEFORE, NOT MERELY NEARBY. A six-line window passed when the label
    # was moved to AFTER its read, because the PREVIOUS read's label was still in
    # the window - so the check is now the nearest preceding statement.
    lines = [line for line in scope.splitlines()
             if line.strip() and not line.strip().startswith("#")]
    for index, line in enumerate(lines):
        if "Invoke-ComRetryRead" not in line or "-Target " not in line:
            continue
        assert index > 0, line
        previous = lines[index - 1].strip()
        assert previous.startswith(("Set-StageBBuildStep ", "Set-StageBBuildOp ")), (
            f"this read does not set its label immediately before it:\n"
            f"  {previous}\n  {line.strip()}")
    # EXECUTED: the reissued read's telemetry line carries the SUB-operation.
    row = _flow()["telemetry"]["retry-then-value"]
    assert "saveas.presave.fileformat" in row["detail"], row



def test_70_a_reissued_read_is_named_and_a_quiet_one_is_not_invented() -> None:
    """EXECUTED. One concise line per read that actually had to be reissued,
    carrying the sub-operation, the attempts and the wait; NO line when the read
    answered first time. A diagnostic that fires either way says nothing."""
    flow = _flow()
    for case in ("int-51", "int-52", "fullname", "object"):
        assert flow["read"][case]["lines"] == 0, (case, flow["read"][case])
    row = flow["telemetry"]["retry-then-value"]
    assert row["outcome"] == "ok", row
    assert "the item" in row["detail"] and "attempts=3" in row["detail"], row
    assert "COMREJECT|build|saveas.presave.fileformat|attempts=3" in row["detail"], row
    assert "RPC_E_CALL_REJECTED" in row["detail"], row



def test_71_there_is_no_read_wrapper_left_to_hide_a_read_behind() -> None:
    """DECLARED, AND RUN 5 IS THE DECLARATION. This used to prove the read wrapper
    could express no write. The wrapper is gone - its single Windows execution
    returned no value - so what must now hold is stronger: the build has exactly
    ONE read mechanism, the accepted helper, and the thing that replaced the
    wrapper's telemetry cannot read, write, or return anything at all.
    """
    code = _build_code()
    assert "function Invoke-StageBBuildRead" not in code
    # ONE MECHANISM. No second reader was invented to take its place.
    for banned in ("function Invoke-StageBRead", "function Get-StageBProperty",
                   "function Read-StageB", "function Invoke-StageBComRead"):
        assert banned not in code, f"a second read mechanism appeared: {banned}"
    # THE TELEMETRY RECORDER TOUCHES NO COM AND RETURNS NOTHING.
    body = _ps_function(code, "Add-StageBReadRejection")
    params = body[body.index("param(") : body.index(")", body.index("param("))]
    for banned in ("ScriptBlock", "$Member", "$Target", "$Key", "$Action"):
        assert banned not in params, f"the telemetry recorder takes {banned}"
    assert "Invoke-ComRetryRead" not in body, "the telemetry recorder performs a read"
    # AND IT WRITES TO NO STREAM. A Write-Output inside it lands in the CALLER's
    # output, which is how a recorder becomes part of a value.
    for banned in ("Write-Output", "Write-Host", "Write-Information", "echo "):
        assert banned not in body, f"the telemetry recorder emits through {banned}"
    # ITS OWN SUPPRESSED LEDGER Add IS NOT A COM CALL, so the ban is applied to the
    # body with that one statement removed - naming it rather than exempting the
    # whole check.
    ledger = [line for line in body.splitlines()
              if "$script:StageBBuildRejections.Add(" in line]
    assert len(ledger) == 1, ledger
    rest = body.replace(ledger[0], "")
    for banned in NEVER_RETRIED:
        assert f".{banned}(" not in rest, f"the telemetry recorder calls {banned}"
    # Its only exits are BARE returns and a suppressed Add.
    assert body.count("{ return }") == 2, body
    assert "$null = $script:StageBBuildRejections.Add(" in body
    assert not re.search(r"return\s+[^}\s]", body), "the telemetry recorder returns a value"
    # EXECUTED: assigning from it yields nothing, and the value it was told about
    # is unchanged afterwards.
    flow = _flow()
    row = flow["telemetry"]["emits"]
    assert row["detail"].startswith("nothing"), row
    assert "value-before=51" in row["detail"] and "value-after=51" in row["detail"], row
    # AND ON THE BRANCH THAT ACTUALLY RECORDS. The quiet path returns early, so a
    # probe that only exercises it never reaches the statement that could emit.
    recorded = flow["telemetry"]["emits-when-recording"]
    assert recorded["detail"].startswith("nothing"), recorded
    assert "lines=1" in recorded["detail"], recorded



def test_72_the_accepted_helper_itself_was_not_broadened() -> None:
    """REQUIRED CONTROL. This batch adds CALL SITES, not capability. The helper and
    its classifier are byte-identical to the revision that closed the reserved-row
    batch, so 'narrow read retry' cannot have quietly become something else."""
    now = _lifecycle()
    then = _at("1e0edb2", "pccm/bootstrap/windows/com_lifecycle.ps1")
    assert now == then, "com_lifecycle.ps1 changed in a batch that may not change it"


def test_73_nothing_that_mutates_the_workbook_is_retried() -> None:
    """REQUIRED CONTROL, NAMED PER OPERATION. The module import, the ThisWorkbook
    write, the button creation, the protection call and the final Save are each
    labelled so a refusal NAMES them - and none of them is inside a retry.

    DECLARED EXCEPTION: SaveAs. Windows named it as the rejected call, and it now
    has a recovery - but not on the strength of the refusal contract. Its call
    moved into `Invoke-StageBSaveAs`, where every reissue is gated on an observed
    postcondition, so the anchor here moved with it and the call is checked by
    test_86 onward instead. The other five are still bare, and this proves it.
    """
    joined = _joined(_build_block())
    # THE SaveAs CALL IS NO LONGER IN THE BLOCK AT ALL, and that is the declaration:
    # it is in one named function and nowhere else in the script.
    code = _joined(_build_code())
    assert "$wb.SaveAs(" not in joined, "the build block still calls SaveAs directly"
    assert code.count(".SaveAs(") == 1, "SaveAs is called in more than one place"
    assert ".SaveAs(" in _joined(_ps_function(_build_code(), "Invoke-StageBSaveAs"))
    mutations = {
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
    assert joined.index("Set-StageBBuildOp 'saveas.xlsm'") < joined.index("Invoke-StageBSaveAs")


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
    """NOT AUTHORISED IN ANY BATCH SO FAR, AND REFUSED HERE RATHER THAN REMEMBERED.
    We still do not know WHY Excel refuses the second-session SaveAs, and a poll
    inserted before that evidence exists could make the symptom disappear without
    proving its cause."""
    code = _build_code()
    outside = code.replace(_ps_function(code, "Invoke-StageBSaveAs"), "")
    assert "Start-Sleep" not in outside, "the bootstrap sleeps outside the SaveAs retry"
    for banned in ("Wait-ExcelReady", "Test-ExcelReady", "Wait-WorkbookReady",
                   "readiness", "-Member 'Ready'", "Start-Process", "Get-Random"):
        assert banned not in code, f"a readiness mechanism appeared: {banned}"
    # THE STRUCTURAL CHECK, NOT A WORD LIST. Between the Open and the save there is
    # the pre-save OBSERVATION and nothing else - no wait, no loop, no poll.
    joined = _joined(_build_block())
    span = joined[joined.index("$wb = $workbooks.Open($stageAPath)")
                  : joined.index("$saveAs = Invoke-StageBSaveAs")]
    for banned in ("Start-Sleep", "while", "do {", "for (", "-Member 'Worksheets'",
                   "-Member 'VBProject'", "-Member 'VBComponents'"):
        assert banned not in span, f"something was inserted after the Open: {span!r}"
    # Exactly the two baseline reads, each labelled.
    assert span.count("Invoke-ComRetryRead") == 2, span
    assert "-Member 'FullName'" in span and "-Member 'FileFormat'" in span
    assert "Set-StageBBuildStep 'saveas.presave.fullname'" in span
    assert "Set-StageBBuildStep 'saveas.presave.fileformat'" in span



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
    """RUN 5 IS THIS CONTROL'S SUBJECT NOW. A read that neither raised nor answered
    must fail at the operation it happened at, and the validation is at the
    CONSUMER with the expected type - not a truthiness test, which would reject a
    legitimate zero."""
    flow = _flow()
    # EXECUTED, ONE CASE PER WAY OF NOT ANSWERING.
    assert flow["read"]["nothing"]["outcome"] == "REFUSED"
    assert "answered with nothing" in flow["read"]["nothing"]["detail"]
    assert flow["read"]["blank"]["outcome"] == "REFUSED"
    assert "empty string" in flow["read"]["blank"]["detail"]
    assert flow["read"]["two-values"]["outcome"] == "REFUSED"
    assert "2 values" in flow["read"]["two-values"]["detail"]
    assert flow["read"]["not-number"]["outcome"] == "REFUSED"
    assert "not an integer" in flow["read"]["not-number"]["detail"]
    # AND A REAL ANSWER SURVIVES, EXACTLY, INCLUDING ITS TYPE.
    assert flow["read"]["int-51"]["value"] == "51"
    assert flow["read"]["int-51"]["type"] == "Int32"
    assert flow["read"]["int-52"]["value"] == "52"
    assert flow["read"]["int-as-text"]["value"] == "52"
    assert flow["read"]["fullname"]["type"] == "String"
    assert flow["read"]["object"]["outcome"] == "ok"
    # NO TRUTHINESS. `if (-not $value)` would refuse a legitimate 0.
    for name in ("Get-StageBScalarInt", "Get-StageBNonEmptyString"):
        body = _ps_function(_build_code(), name)
        assert "if (-not $Value)" not in body, f"{name} uses truthiness"
        assert "if ($null -eq $Value) {" in body, body
        assert "$Value -is [System.Array]" in body, body
    assert "^-?[0-9]+$" in _ps_function(_build_code(), "Get-StageBScalarInt")



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


# ===========================================================================
# H. SaveAs: WINDOWS NAMED IT, AND ITS POSTCONDITION SETTLES IT
# ===========================================================================
# THE DIAGNOSTIC WORKED. The labelled build reported, repeatably:
#
#     [FAIL] Stage-B build
#            operation=saveas.xlsm; ... 0x80010001 RPC_E_CALL_REJECTED
#     COMREJECT|build|none|attempts=0|waited=0
#
# The second line matters as much as the first: no read was reissued, so the read
# retry had nothing to do with this. SaveAs writes a file and rebinds the
# workbook, so 'the message filter says the call never ran' is a CONTRACT and not
# an observation - and these controls exist to keep the licence being the
# observation.
def test_86_the_save_is_one_dedicated_function_and_not_a_general_retry() -> None:
    """REQUIRED CONTROL: SaveAs is the ONLY non-read operation with special
    treatment, and the treatment lives in one named function that the accepted
    read helper knows nothing about."""
    code = _build_code()
    saveas = _ps_function(code, "Invoke-StageBSaveAs")
    # It does not use, wrap or re-implement the read helper's loop.
    assert "Invoke-ComRetryRead" not in saveas, "the SaveAs retry goes through the read helper"
    # The classifier IS shared, deliberately: a second copy of "which HRESULTs are
    # refusals" is how two answers to one question appear.
    assert "Get-ComRejectionName" in saveas
    assert "$script:ComRetryableHResults" not in code, "the HRESULT table was copied"
    # And no OTHER mutation gained a function like this one.
    for banned in ("Invoke-StageBImport", "Invoke-StageBSave ", "Invoke-StageBProtect",
                   "Invoke-StageBAddShape", "Invoke-StageBAddFromString"):
        assert f"function {banned}" not in code, f"{banned} exists"


def test_87_the_accepted_read_helper_is_still_byte_identical() -> None:
    """REQUIRED CONTROL. The SaveAs settlement may not have reached into the read
    retry on its way past."""
    assert _lifecycle() == _at("3d34b26", "pccm/bootstrap/windows/com_lifecycle.ps1")


def test_88_only_the_two_rpc_rejections_can_lead_to_a_reissued_save() -> None:
    """EXECUTED. An HRESULT Excel returned from a call it ACCEPTED, and an error
    that is not a COM failure at all, are rethrown after exactly ONE SaveAs - and
    no postcondition is even inspected, because there is nothing to decide."""
    flow = _flow()
    assert not flow["unmet"], flow["unmet"]
    for case, expected in (("accepted-err", "0x800a03ec"), ("not-com-err", "no-hresult")):
        row = flow["save"][case]
        assert row["outcome"] == "RAISED", (case, row)
        assert row["calls"] == 1, f"{case} reissued SaveAs {row['calls']} times"
        assert row["state"] == expected, (case, row["state"])
        notes = flow["savenote"][case]
        assert any(line.startswith("SAVEAS|attempt=1|error|") for line in notes), notes
        assert not any("postcondition" in line for line in notes), (
            f"{case} inspected a postcondition for a call Excel accepted")
    # And both refusals DO reach the settlement.
    for case, code in (("not-executed", "0x80010001"), ("retrylater", "0x8001010a")):
        notes = flow["savenote"][case]
        assert f"SAVEAS|attempt=1|rejected|{code}" in notes, notes


def test_89_a_clean_save_is_still_verified_before_anything_is_built_on_it() -> None:
    """REQUIRED CONTROL, EXECUTED. A COM method that returned is not a save that
    happened. All three facts are checked even on the quiet path."""
    flow = _flow()
    row = flow["save"]["clean"]
    assert (row["calls"], row["attempts"], row["waited"]) == (1, 1, 0), row
    assert row["state"] == "completed" and row["format"] == 52, row
    notes = flow["savenote"]["clean"]
    assert "SAVEAS|attempt=1|success" in notes, notes
    assert "SAVEAS|verified|path=True|format=52|exists=True" in notes, notes
    # EXECUTED, AND THIS IS THE CASE THE MUTATION BATTERY EXPOSED AS UNTESTED. A
    # SaveAs that RETURNS and does nothing must FAIL. Asserting only that the
    # refusal's sentence is present let an `if ($false)` in front of it survive.
    silent = flow["save"]["silent-success"]
    assert silent["outcome"] == "RAISED", f"a save that did nothing was accepted: {silent}"
    assert silent["calls"] == 1, silent
    notes = flow["savenote"]["silent-success"]
    assert "SAVEAS|attempt=1|success" in notes, notes
    assert "SAVEAS|verified|path=False|format=51|exists=False" in notes, notes
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    assert "returned without error but its postconditions do not prove" in saveas
    assert "if ($state.State -ne 'completed') {" in saveas


def test_90_the_three_authoritative_facts_are_all_three_required() -> None:
    """THE CONDITION, NOT THE COMMENT. COMPLETED needs the workbook bound to the
    target AND the target format AND the file present; NOT EXECUTED needs the
    source on all three."""
    checker = _ps_function(_build_code(), "Get-StageBSaveAsPostcondition")
    assert "$boundToTarget -and ($format -eq $TargetFormat) -and $targetExists" in checker
    assert "$boundToSource -and ($format -eq $SourceFormat) -and (-not $targetExists)" in checker
    assert "$state = 'ambiguous'" in checker, "ambiguous is not the default"
    # The default must be set BEFORE the two positive tests, so an unmatched state
    # falls to ambiguous rather than to whatever was there last.
    assert checker.index("$state = 'ambiguous'") < checker.index("$state = 'completed'")
    # AND THE READS ARE READS. Nothing in the checker writes.
    for banned in (".SaveAs(", ".Save()", "Remove-Item", "Set-Content", "New-Item",
                   ".Delete(", "= $TargetPath", "Move-Item", "Copy-Item"):
        assert banned not in checker, f"the postcondition checker mutates: {banned}"


def test_91_a_completed_save_is_never_reissued() -> None:
    """EXECUTED, AND THE MOST DANGEROUS CASE. A SaveAs that completed and then
    reported a rejection must be ACCEPTED: reissuing it would overwrite the
    workbook this build already owns."""
    flow = _flow()
    row = flow["save"]["completed"]
    assert row["calls"] == 1, f"a completed save was reissued ({row['calls']} calls)"
    assert row["outcome"] == "returned" and row["state"] == "completed", row
    notes = flow["savenote"]["completed"]
    assert "SAVEAS|postcondition|completed|" in "|".join(notes) or any(
        line.startswith("SAVEAS|postcondition|completed") for line in notes), notes
    assert any("completed-despite-rejection" in line for line in notes), notes


def test_92_a_save_proved_not_to_have_happened_may_be_reissued_once_more() -> None:
    """EXECUTED. This is the only state that permits a second SaveAs, and the
    second one is what makes the run recoverable rather than merely diagnosed."""
    flow = _flow()
    for case, code in (("not-executed", "0x80010001"), ("retrylater", "0x8001010a")):
        row = flow["save"][case]
        assert row["calls"] == 2, (case, row)
        assert row["attempts"] == 2 and row["state"] == "completed", (case, row)
        assert row["waited"] > 0, f"{case} did not back off at all"
        notes = flow["savenote"][case]
        assert any(line.startswith("SAVEAS|postcondition|not-executed") for line in notes), notes
        assert "SAVEAS|attempt=2|success" in notes, notes


def test_93_an_ambiguous_save_is_never_reissued_and_never_cleaned_up() -> None:
    """EXECUTED, FIVE PARTIAL STATES. Target present but still bound to the source;
    rebound with no file; the format moved and nothing else; rebound with a file
    but the wrong format; rebound and reformatted with no file. Each is exactly
    ONE SaveAs and an abort - and the evidence is left where it is."""
    flow = _flow()
    for case in ("amb-file-only", "amb-path-only", "amb-fmt-only",
                 "amb-path-file", "amb-path-fmt"):
        row = flow["save"][case]
        assert row["outcome"] == "RAISED", (case, row)
        assert row["calls"] == 1, f"{case} reissued an ambiguous save ({row['calls']})"
        notes = flow["savenote"][case]
        assert any(line.startswith("SAVEAS|postcondition|ambiguous") for line in notes), notes
    # NOTHING IS DELETED OR MOVED to tidy an ambiguous save away.
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    for banned in ("Remove-Item", "Move-Item", "Copy-Item", ".Delete(", "Clear-Content"):
        assert banned not in saveas, f"an ambiguous save is cleaned up: {banned}"
    assert "nothing is cleaned up" in saveas


def test_94_no_single_fact_is_enough_on_its_own() -> None:
    """REQUIRED CONTROLS 12-14, EXECUTED AS THREE SEPARATE CASES. The file
    existing, the path having changed, and the format having changed are each, on
    their own, AMBIGUOUS - because each is what a half-completed save looks like."""
    flow = _flow()
    detail = {case: "|".join(flow["savenote"][case]) for case in
              ("amb-file-only", "amb-path-only", "amb-fmt-only")}
    # target exists, nothing else moved
    assert "targetExists=True" in detail["amb-file-only"]
    assert "boundToSource=True" in detail["amb-file-only"]
    assert flow["save"]["amb-file-only"]["outcome"] == "RAISED"
    # rebound, but no file and the old format
    assert "boundToTarget=True" in detail["amb-path-only"]
    assert "targetExists=False" in detail["amb-path-only"]
    assert flow["save"]["amb-path-only"]["outcome"] == "RAISED"
    # the format moved and nothing else
    assert "format=52" in detail["amb-fmt-only"]
    assert "boundToTarget=False" in detail["amb-fmt-only"]
    assert flow["save"]["amb-fmt-only"]["outcome"] == "RAISED"


def test_95_an_unreadable_postcondition_is_ambiguous_and_not_evidence() -> None:
    """EXECUTED. If the inspection itself cannot be answered, that is not proof of
    anything - least of all that the save did not happen."""
    flow = _flow()
    state, detail = flow["savestate"]["unreadable"]
    assert state == "ambiguous", (state, detail)
    assert "readError=" in detail and "readError=none" not in detail
    # A workbook that answers a read with NOTHING is the same finding.
    row = flow["save"]["amb-null-name"]
    assert row["outcome"] == "RAISED" and row["calls"] == 1, row
    assert any(line.startswith("SAVEAS|postcondition|ambiguous")
               for line in flow["savenote"]["amb-null-name"])


def test_96_the_inspection_always_precedes_the_next_save() -> None:
    """REQUIRED CONTROL 8. In source: on a refused call the postcondition is read
    BEFORE any decision, and every path out of that catch either returns, throws,
    or falls through to the bounded backoff."""
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    rejected = saveas.index("Add-Note ('SAVEAS|attempt=' + [string]$attempt + '|rejected|'")
    inspect = saveas.index("$state = Get-StageBSaveAsPostcondition", rejected)
    completed = saveas.index("if ($state.State -eq 'completed') {", inspect)
    notexec = saveas.index("if ($state.State -ne 'not-executed') {", inspect)
    sleep_at = saveas.index("Start-Sleep", inspect)
    assert rejected < inspect < completed < notexec < sleep_at, "the order is wrong"
    # The second SaveAs is only reachable after that sequence.
    assert saveas.count(".SaveAs(") == 1, "there is more than one SaveAs statement"
    assert saveas.index(".SaveAs(") < inspect


def test_97_both_bounds_hold_on_the_save_as_well() -> None:
    """EXECUTED, AND SEPARATELY. Three attempts allowed means three SaveAs calls;
    a five-millisecond budget ends it well before fifty. And the ORIGINAL rejection
    is what escapes - not a substitute."""
    flow = _flow()
    attempt_bound = flow["save"]["attempt-bound"]
    assert attempt_bound["calls"] == 3, attempt_bound
    assert attempt_bound["outcome"] == "RAISED"
    assert attempt_bound["state"] == "0x80010001", attempt_bound
    budget = flow["save"]["budget-bound"]
    assert budget["calls"] < 50, budget
    assert budget["outcome"] == "RAISED" and budget["state"] == "0x80010001", budget
    # The bound is in the loop head, as it is in the read helper.
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    heads = re.findall(r"while\s*\(([^)]*)\)\s*\{", saveas)
    assert heads and all("-lt " in head for head in heads), heads
    assert "while ($true)" not in saveas and "do {" not in saveas
    assert "if ($attempt -ge $MaxAttempts) { throw }" in saveas
    assert "if (($waitedMs + $delay) -gt $TotalBudgetMs) { throw }" in saveas
    assert "$waitedMs = $waitedMs + $delay" in saveas
    # A LOOP THAT ENDED WITHOUT SAVING NEVER RETURNS A SAVE. The head bound creates
    # a way out that no path currently takes, so it gets the same guard the read
    # helper's does - and the mutation battery proved this was missing.
    guard = saveas.index("throw ('Invoke-StageBSaveAs: the save was never completed after '")
    assert saveas.rindex("New-StageBSaveAsResult") < guard, \
        "a result can be returned after the loop ended without a save"


def test_98_a_path_is_compared_as_a_path_and_not_as_a_string() -> None:
    """EXECUTED. A separator or case difference is not a rebind, and reading one as
    a rebind would manufacture an ambiguous state out of nothing."""
    flow = _flow()
    assert flow["savepath"]["same"] is True, "the same path compared unequal"
    assert flow["savepath"]["different"] is False, "two different paths compared equal"
    norm = _ps_function(_build_code(), "Get-StageBComparablePath")
    assert "GetFullPath" in norm and "ToLowerInvariant" in norm
    assert "DirectorySeparatorChar" in norm


def test_99_the_baseline_is_observed_before_the_save_and_proved_consistent() -> None:
    """'NOTHING MOVED' IS ONLY PROVABLE AGAINST WHAT THE WORKBOOK WAS, so the
    baseline is READ - never assumed - and it is checked before anything is saved.
    Run 5 is what an unchecked baseline looks like: the read answered with nothing
    and the run had no business attempting a save at all."""
    joined = _joined(_build_block())
    assert "-Description 'the Stage-A workbook FullName before SaveAs'" in joined
    assert "-Description 'the Stage-A workbook FileFormat before SaveAs'" in joined
    name_at = joined.index("$preName = Invoke-ComRetryRead")
    fmt_at = joined.index("$preFormat = Invoke-ComRetryRead")
    call_at = joined.index("$saveAs = Invoke-StageBSaveAs")
    assert name_at < fmt_at < call_at, "the baseline is read after the save"
    assert "-SourceFormat $sourceFormat" in joined
    assert "-SourceFullName $sourceFullName" in joined
    # AND THE BASELINE MUST BE CONSISTENT OR THE SAVE IS NOT ATTEMPTED.
    assert "SAVEAS BASELINE: the workbook is bound to " in joined
    assert "so NOT EXECUTED could not be recognised. The save was not attempted." in joined
    assert "is present before the save, so " in joined
    # UNIQUELY ANCHORED, AND THE CONDITION RATHER THAN THE MESSAGE. The string
    # `if (Test-Path -LiteralPath $stageBPath) {` also opens the stale-target
    # deletion, so checking it alone passed over a DISABLED baseline refusal - the
    # mutation battery caught this control doing exactly that. Each refusal is now
    # matched as its own condition-plus-throw pair, inside the pre-save region.
    region = joined[joined.index("Set-StageBBuildStep 'saveas.presave.fullname'") : call_at]
    for condition, message in (
            ("if ($preSeen -ne (Get-StageBComparablePath $stageAPath)) {",
             "SAVEAS BASELINE: the workbook is bound to "),
            ("if (Test-Path -LiteralPath $stageBPath) {",
             "SAVEAS BASELINE: ' + $stageBPath + ' is present before the save")):
        assert condition in region, f"the refusal condition is gone: {condition}"
        assert message in region, f"the refusal message is gone: {message}"
        assert region.index(condition) < region.index(message), (condition, message)
    # THE TARGET FORMAT STILL COMES FROM THE MANIFEST, and the checker holds no
    # literal format of its own.
    assert "-TargetFormat ([int]$manifest.xlsm_file_format)" in joined
    assert "51" not in _ps_function(_build_code(), "Get-StageBSaveAsPostcondition")
    assert _manifest()["xlsm_file_format"] == 52



def test_100_the_build_still_deletes_the_target_before_saving() -> None:
    """THIS IS WHAT MAKES THE OBSERVATION CONCLUSIVE. Because the target provably
    does not exist when SaveAs is attempted, 'the file is there' afterwards means
    this call put it there. Leave that deletion alone and the whole settlement
    weakens to a guess."""
    joined = _joined(_build_block())
    delete_at = joined.index("Remove-Item -LiteralPath $stageBPath -Force")
    call_at = joined.index("$saveAs = Invoke-StageBSaveAs")
    assert delete_at < call_at, "the stale target is removed after the save"
    assert "if (Test-Path -LiteralPath $stageBPath) { Remove-Item" in joined


def test_101_the_five_other_mutations_still_have_no_recovery_path() -> None:
    """REQUIRED CONTROLS 20-24. The import, the document module, the buttons, the
    protection and the final Save are untouched by this batch: no retry, no
    postcondition, no second attempt."""
    code = _joined(_build_code())
    for call in ("$vbcomps.Import($file)", "$codeModule.AddFromString($docText)",
                 "$shapes.AddShape(5,", "$pws.Protect([Type]::Missing", "$wb.Save()"):
        assert call in code, f"the call site is gone: {call}"
        assert code.count(call) == 1, f"{call} appears more than once"
        for line in code.splitlines():
            if call in line:
                for banned in ("Invoke-ComRetryRead", "Invoke-StageBBuildRead",
                               "Invoke-StageBSaveAs", "Postcondition"):
                    assert banned not in line, line
    # And none of them sits inside a loop that could reissue it.
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    for call in ("Import(", "AddFromString(", "AddShape(", ".Protect(", "$wb.Save()"):
        assert call not in saveas, f"{call} was pulled into the SaveAs retry"


def test_102_the_gate_still_refuses_a_half_built_workbook() -> None:
    """REQUIRED CONTROL 25, AND MORE IMPORTANT NOW THAN WHEN IT WAS ADDED. An
    ambiguous save can leave a target file on disk while Stage-B exits nonzero, and
    target-file existence must never read as a successful bootstrap."""
    pass_fn = _ps_function(_gate(), "Invoke-EquivalencePass")
    assert "$bootstrapExit = $LASTEXITCODE" in pass_fn
    assert pass_fn.index("if ($bootstrapExit -ne 0) {") < \
        pass_fn.index("if (-not (Test-Path -LiteralPath $stageB)) {")
    assert "half-built" in pass_fn
    assert "BOOTSTRAP:" in pass_fn
    # The bootstrap really does exit nonzero on a build failure.
    code = _build_code()
    assert "Add-Step 'Stage-B build' 'FAIL'" in code
    assert "if ($Status -eq 'FAIL') { $null = $failures.Add($Name) }" in code
    assert "if ($failures.Count -eq 0) {" in code and "exit 1" in code


def test_103_the_freezes_this_batch_may_not_touch() -> None:
    """REQUIRED CONTROLS 26-29. Production VBA, the benchmark runner, the
    reserved-row correction and the equivalence snapshot are all byte-identical to
    the Windows-tested revision."""
    done = subprocess.run(["git", "diff", "--name-only", "3d34b26", "--", "pccm/src/vba"],
                          cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE, text=True)
    assert done.stdout.strip() == "", done.stdout
    assert BENCHMARK_PS1.read_text(encoding="utf-8") == \
        _at("3d34b26", "pccm/bootstrap/windows/phase10_benchmark.ps1"), \
        "the benchmark runner changed"
    for name in ("Set-BenchmarkRegisterRowCount", "New-BenchmarkRegisterBlock",
                 "Get-BenchmarkPermanentId", "Set-BenchmarkBulkFixture"):
        now = _ps_function(BENCHMARK_PS1.read_text(encoding="utf-8"), name)
        then = _ps_function(_at("3d34b26", "pccm/bootstrap/windows/phase10_benchmark.ps1"), name)
        assert now == then, f"{name} changed"
    assert _ps_function(_gate(), "Get-EquivalenceSnapshot") == \
        _ps_function(_at("99cb472", "pccm/tests/phase10_fixture_equivalence.ps1"),
                     "Get-EquivalenceSnapshot"), "Get-EquivalenceSnapshot moved"


def test_104_the_diagnostic_no_longer_asserts_a_contract_as_a_fact() -> None:
    """THE CORRECTED WORDING. 'The call was refused before it ran' is what the
    message-filter contract says. For a property get that is the whole answer; for
    SaveAs it is not, and the record now says which kind of claim it is."""
    code = _build_code()
    line = _ps_function(code, "New-StageBRejectionLine")
    assert "refused before it ran" in line, "the read-side wording disappeared entirely"
    # THE QUALIFICATION IS IN THE SOURCE where the distinction lives.
    # RAW SOURCE, not the comment-stripped view: the distinction between a
    # contract and an observation is exactly the thing that has to be written down
    # for the next person, and _build_code() blanks comments by design.
    raw = _build()
    header = raw[raw.index("# SaveAs: A NON-IDEMPOTENT CALL")
                 : raw.index("function Get-StageBComparablePath")]
    assert "contract" in header.lower(), header[-800:]
    assert "OBSERVED STATE" in header or "observation" in header.lower()
    assert "non-idempotent" in header.lower()
    # And the evidence record carries the correction.
    text = _evidence()
    at = text.index("## The diagnostic run that named the rejected call")
    after = text.find("\n## ", at + 10)
    section = text[at:] if after == -1 else text[at:after]
    assert "saveas.xlsm" in section
    assert "COMREJECT|build|none|attempts=0|waited=0" in section
    assert "not sufficient proof" in section or "is not proof" in section
    assert "non-idempotent" in section
    # THE PROPERTY IS "NOT ASSERTED", NOT "NOT MENTIONED". The record has to name
    # the speculations it is refusing, or a later reader cannot tell a considered
    # refusal from an omission - so the ban applies to everything BEFORE the
    # disclaimer, and the disclaimer itself must name them.
    marker = "### What is still NOT known, and is not claimed"
    assert marker in section, section[-600:]
    asserted, disclaimed = section.split(marker, 1)
    for overclaim in ("lifecycle overlap", "readiness race", "OneDrive",
                      "file locking", "modal state", "message-filter timing"):
        assert overclaim.lower() not in asserted.lower(), \
            f"the record speculates beyond the evidence: {overclaim}"
        assert overclaim.lower() in disclaimed.lower(), \
            f"the record does not say {overclaim} is unproved"
    assert "do not know WHY" in disclaimed or "not know why" in disclaimed.lower()
    assert "not root cause" in disclaimed.lower() or "not speculative root cause" in disclaimed.lower()
    assert "INVALID / NOT EVALUATED" in section
    assert "must not be read as a fixture DIFFER" in section
    assert "why" in section.lower()


# ===========================================================================
# I. RUN 5: THE READ THAT ANSWERED WITH NOTHING
# ===========================================================================
# THE PRE-SAVE OBSERVATION FAILED, AND REPORTED ITSELF AS THE SAVE. Run 5:
#
#     operation=saveas.xlsm; Invoke-StageBBuildRead: the Stage-A workbook
#     FileFormat before SaveAs answered with nothing at saveas.xlsm.
#     COMREJECT|build|none|attempts=0|waited=0
#
# SaveAs was never attempted. Two things had to change: the label had to name the
# sub-operation, and the read had to use the form Windows has actually executed.
def test_105_every_build_read_uses_the_windows_proven_form() -> None:
    """THE COMPARATOR IS THE REOPEN VERIFICATION, which has read FileFormat,
    Worksheets, VBProject, VBComponents, CodeName, Count, Item, Shapes, OnAction
    and Name successfully on Windows across every accepted run - through the
    accepted helper, called DIRECTLY. Every build read is now that same shape, and
    no intermediate reader stands between the COM object and the member access."""
    code = _build_code()
    scope = _joined(_build_block()) + "\n" + "\n".join(
        _joined(_ps_function(code, name)) for name in SAVEAS_HELPERS)
    reads = [line for line in re.findall(r"[^\n]*Invoke-ComRetryRead\b[^\n]*", scope)
             if "-Target " in line]
    assert len(reads) >= 6, reads
    for line in reads:
        # THE PROVEN SHAPE: the record is captured, then its .Value is taken in a
        # separate statement or by a validator - never through another function.
        assert re.search(r"\$\w+ = Invoke-ComRetryRead ", line), line
        assert ".Value" not in line, (
            f"the value is taken in the same expression as the read: {line}")
    # AND THE VERIFICATION BLOCK IS UNTOUCHED - it is the comparator, so it may not
    # be quietly changed to match whatever the build now does.
    assert _verify_block() == _ps_verify_at("cc9cf8d"), (
        "the Windows-proven verification block changed")


def _ps_verify_at(commit: str) -> str:
    """The verification block as it stood at a commit, through the same slicer."""
    code = _strip(_at(commit, "pccm/bootstrap/windows/build_stage_b.ps1"))
    start = code.index("$excel2 = $null; $workbooks2 = $null")
    end = code.index("$rel2 = New-ReleaseLedger 'verification instance'")
    return code[start:end]


def test_106_the_value_is_never_produced_by_an_expression_that_also_reports() -> None:
    """ASSIGN, NEVER EMIT - AND THE INVERSE. The read is one statement; the
    telemetry is another; the validation is a third. A value cannot be polluted by
    a report that is not in its expression, and it cannot be lost to one either."""
    code = _build_code()
    scope = _joined(_build_block()) + "\n" + "\n".join(
        _joined(_ps_function(code, name)) for name in SAVEAS_HELPERS)
    for line in scope.splitlines():
        if "Add-StageBReadRejection" not in line:
            continue
        stripped = line.strip()
        assert stripped.startswith("Add-StageBReadRejection"), (
            f"the telemetry is part of another expression: {line}")
        assert "=" not in stripped.split("-Operation")[0], line
    # EVERY read's record is consumed by a validator or a null check, never
    # discarded: a read whose answer nobody looks at is run 5 waiting to happen.
    for record in re.findall(r"\$(\w*[Rr]ead|preName|preFormat|nameRead|formatRead) = Invoke-ComRetryRead", scope):
        assert f"${record}.Value" in scope, f"${record} is read and never examined"


def test_107_the_sub_operation_vocabulary_is_closed_and_complete() -> None:
    """EXECUTED. Seven sub-operations, each accepted; a misspelling REFUSED and the
    label left at the top-level operation; and a new top-level operation clears the
    step, because a stale one would name a step that already finished."""
    flow = _flow()
    assert not flow["unmet"], flow["unmet"]
    assert tuple(flow["subops"]) == BUILD_SUB_OPS, flow["subops"]
    for step in BUILD_SUB_OPS:
        label, outcome = flow["step"][step]
        assert outcome == "accepted", (step, outcome)
        assert label == step, (step, label)
    label, outcome = flow["step"]["typo"]
    assert outcome == "REFUSED", flow["step"]["typo"]
    assert label == "saveas.xlsm", "a refused sub-operation became the label"
    cleared = flow["step"]["cleared"]
    assert cleared[0] == "vbcomponents.import", cleared
    assert cleared[1] == "", "a stale sub-operation survived a new operation"


def test_108_every_sub_operation_maps_to_a_real_call_site() -> None:
    """A VOCABULARY ENTRY WITH NO CALL SITE CAN NEVER APPEAR IN A DIAGNOSTIC."""
    code = _build_code()
    scope = _joined(_build_block()) + "\n" + "\n".join(
        _joined(_ps_function(code, name)) for name in SAVEAS_HELPERS)
    for step in BUILD_SUB_OPS:
        assert f"Set-StageBBuildStep '{step}'" in scope, f"{step} is never set"
    used = set(re.findall(r"Set-StageBBuildStep '([^']+)'", scope))
    assert used == set(BUILD_SUB_OPS), sorted(set(BUILD_SUB_OPS) ^ used)


def test_109_a_failed_observation_is_not_reported_as_a_failed_save() -> None:
    """RUN 5's ACTUAL DEFECT. The pre-save read reported operation=saveas.xlsm,
    which reads as though the save had been attempted. The label now distinguishes
    the observation from the call."""
    code = _build_code()
    # The presave reads set presave labels; the call sets saveas.call.
    joined = _joined(_build_block())
    fmt_at = joined.index("$preFormat = Invoke-ComRetryRead")
    assert joined.index("Set-StageBBuildStep 'saveas.presave.fileformat'") < fmt_at
    saveas = _joined(_ps_function(code, "Invoke-StageBSaveAs"))
    call_at = saveas.index(".SaveAs(")
    assert saveas.index("Set-StageBBuildStep 'saveas.call'") < call_at
    # And the record says so.
    text = _evidence()
    at = text.index("## Equivalence run 5")
    after = text.find("\n## ", at + 10)
    section = text[at:] if after == -1 else text[at:after]
    # BACKTICKS AND BOLD ARE MARKUP, NOT MEANING. The claim is checked with the
    # markup removed so a record cannot fail a control for emphasising the right
    # sentence - nor pass one by dropping it.
    plain = section.replace("`", "").replace("**", "")
    assert "SaveAs itself was NOT executed" in plain, plain[:400]
    assert "saveas.presave.fileformat" in section
    assert "no rpc rejection occurred" in plain.lower(), plain[:600]
    assert "attempts=0" in plain
    assert "INVALID / NOT EVALUATED" in section
    assert "must not be read as a fixture DIFFER" in section
    # STATED POSITIVELY, because a blanket ban on the words caught the record's own
    # DISCLAIMER - it has to name what it is refusing to be read as. These two
    # sentences together are the property, and the mutation battery removes each.
    assert "not another SaveAs rejection" in plain, plain[:600]
    assert "INVALID / NOT EVALUATED" in plain
    assert "Bulk remains NOT authorised" in plain


def test_110_the_baseline_gates_the_save() -> None:
    """REQUIRED CONTROL 13. Without a trustworthy baseline, NOT EXECUTED cannot be
    recognised - so a save is not attempted at all. Run 5 would have carried
    [int]$null = 0 into that classification and called it the source format."""
    joined = _joined(_build_block())
    call_at = joined.index("$saveAs = Invoke-StageBSaveAs")
    head = joined[:call_at]
    # The two reads, the two validators and the two consistency refusals all
    # precede the save.
    for required in ("Get-StageBNonEmptyString -Value $preName.Value",
                     "Get-StageBScalarInt -Value $preFormat.Value",
                     "SAVEAS BASELINE:",
                     "Add-Note ('SAVEAS|baseline|"):
        assert required in head, f"{required} does not precede the save"
    # A validator that threw means the save is unreachable: they are in the same
    # straight-line region, with no try/catch swallowing them.
    region = head[head.index("Set-StageBBuildStep 'saveas.presave.fullname'"):]
    assert "catch" not in region, f"a baseline failure can be swallowed: {region!r}"


def test_111_the_three_state_settlement_is_unchanged_in_substance() -> None:
    """SEMANTIC FREEZE. This batch fixed a read. The classification it feeds must be
    exactly the accepted one."""
    checker = _ps_function(_build_code(), "Get-StageBSaveAsPostcondition")
    assert "$boundToTarget -and ($format -eq $TargetFormat) -and $targetExists" in checker
    assert "$boundToSource -and ($format -eq $SourceFormat) -and (-not $targetExists)" in checker
    assert "$state = 'ambiguous'" in checker
    assert checker.index("$state = 'ambiguous'") < checker.index("$state = 'completed'")
    saveas = _ps_function(_build_code(), "Invoke-StageBSaveAs")
    assert "if ($state.State -eq 'completed') {" in saveas
    assert "if ($state.State -ne 'not-executed') {" in saveas
    assert "SAVEAS AMBIGUOUS after " in saveas
    assert saveas.count(".SaveAs(") == 1


def test_112_the_freezes_this_batch_may_not_touch() -> None:
    """REQUIRED CONTROLS 17-20, against the Windows-tested revision."""
    done = subprocess.run(["git", "diff", "--name-only", "cc9cf8d", "--", "pccm/src/vba"],
                          cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE, text=True)
    assert done.stdout.strip() == "", done.stdout
    assert BENCHMARK_PS1.read_text(encoding="utf-8") == \
        _at("cc9cf8d", "pccm/bootstrap/windows/phase10_benchmark.ps1")
    assert _lifecycle() == _at("cc9cf8d", "pccm/bootstrap/windows/com_lifecycle.ps1")
    assert _gate() == _at("cc9cf8d", "pccm/tests/phase10_fixture_equivalence.ps1")
    for name in ("Set-BenchmarkRegisterRowCount", "New-BenchmarkRegisterBlock",
                 "Get-BenchmarkPermanentId", "Set-BenchmarkBulkFixture"):
        now = _ps_function(BENCHMARK_PS1.read_text(encoding="utf-8"), name)
        then = _ps_function(_at("cc9cf8d", "pccm/bootstrap/windows/phase10_benchmark.ps1"), name)
        assert now == then, f"{name} changed"
    assert _ps_function(_gate(), "Get-EquivalenceSnapshot") == \
        _ps_function(_at("99cb472", "pccm/tests/phase10_fixture_equivalence.ps1"),
                     "Get-EquivalenceSnapshot")


if __name__ == "__main__":
    import pytest
    raise SystemExit(pytest.main([__file__, "-q"]))
