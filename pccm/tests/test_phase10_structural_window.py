#!/usr/bin/env python3
"""P10-RP / THE STRUCTURAL PROTECTION WINDOW.

WHAT WINDOWS DISPROVED, AND IT IS NOT BEING REWRITTEN. The accepted Phase-10
contract (6ab8f6a) held that UserInterfaceOnly:=True let every accepted command
mutate the workbook without ever unprotecting anything. That was recorded
honestly and it was wrong about one specific capability. It stands as historical
evidence.

The corrected protection probe invoked the real production endpoint on Windows
11 / Excel 64-bit, with 14 of 14 worksheets protected and structure protected:

    PCCM_ApplyTimeline   invoked: TRUE   outcome: REFUSED
    FAIL|Error 1004: Table features aren't available because the sheet is
    protected.|...restored to their state from before this operation...
    protection before 14/14 structure=True; after 14/14 structure=True
    structural effect: NONE

and, separately and on the same run, a code-driven VALUE write to a proved-locked
cell SUCCEEDED and restored exactly. So UserInterfaceOnly separates two
capabilities:

    ordinary code-driven cell VALUE writes ............. PERMITTED
    ListObject STRUCTURAL mutation while protected ..... NOT PERMITTED

THE RECONCILED RULE these controls enforce:

    Only modProtection may temporarily release WORKSHEET protection, inside the
    contracted structural-operation envelope, and it must restore the full
    accepted protection state before the operation returns.

Workbook STRUCTURE protection is not released - the 1004 named the sheet, and no
evidence asks for more than that.

THESE TESTS DO NOT CLAIM THE FIX WORKS ON WINDOWS. VBA cannot be compiled or run
here. Only a clean Windows protection run can claim that, and this batch has not
had one.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys

import pytest
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
SRC = PCCM_ROOT / "src" / "vba"
PROBE = PCCM_ROOT / "bootstrap" / "windows" / "phase10_protection_probe.ps1"
EVIDENCE = PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md"

sys.path.insert(0, str(PCCM_ROOT / "tests"))

# The seven commands whose forward work or rollback can perform a ListObject
# structural operation. Confirmed from source, not assumed from the batch note:
# every one of them reaches ListRows.Add/Delete or ListColumns.Add/Delete, or
# reaches modWorkbook.RestoreTable, which does all four.
STRUCTURAL_COMMANDS = {
    "PCCM_ApplyTimeline": "modTimeline",
    "PCCM_AddCostLine": "modDrivers",
    "PCCM_DeleteCostLine": "modDrivers",
    "PCCM_AddRisk": "modDrivers",
    "PCCM_DeleteRisk": "modDrivers",
    "PCCM_Calculate": "modCalcReport",
    "PCCM_RepairProfiling": "modRepair",
}

# The four that reach no ListObject structural operation at all. Reset clears
# ranges; the three stochastic commands read a prepared package and write values.
NON_STRUCTURAL_COMMANDS = {
    "PCCM_ResetResults": "modReset",
    "PCCM_RunSimulation": "modSimReport",
    "PCCM_RunSensitivity": "modSimPostReport",
    "PCCM_RunAnnualStochastic": "modSimAnnualRun",
}

STRUCTURAL_OWNERS = sorted(set(STRUCTURAL_COMMANDS.values()))
NON_STRUCTURAL_OWNERS = sorted(set(NON_STRUCTURAL_COMMANDS.values()))

_MEMO: dict = {}


def _src(name: str) -> str:
    if name not in _MEMO:
        _MEMO[name] = (SRC / name).read_bytes().decode("utf-8")
    return _MEMO[name]


def _code(name: str) -> str:
    key = name + "::code"
    if key not in _MEMO:
        _MEMO[key] = "\n".join(line for line in _src(name).splitlines()
                               if not line.strip().startswith("'"))
    return _MEMO[key]


def _procedure(name: str, proc: str) -> str:
    """One VBA procedure body, from its declaration to its End."""
    lines = _src(name).splitlines()
    head = re.compile(r"^\s*(?:Public\s+|Private\s+)?(?:Static\s+)?(Sub|Function)\s+"
                      + re.escape(proc) + r"\b")
    starts = [i for i, line in enumerate(lines) if head.match(line)]
    assert len(starts) == 1, f"{name}.{proc} is declared {len(starts)} times"
    begin = starts[0]
    end = next(j for j in range(begin + 1, len(lines))
               if re.match(r"^End\s+(Sub|Function)\b", lines[j]))
    return "\n".join(lines[begin:end + 1])


def _all_modules() -> list[str]:
    return sorted(p.name for p in SRC.glob("*.bas"))


def _evidence_section(heading: str) -> str:
    """One run's section of the evidence record, from its heading to the next.

    The document is append-only history, so a document-wide check drifts into
    finding a LATER run's numbers and calling them an earlier run's.
    """
    text = EVIDENCE.read_text(encoding="utf-8")
    start = text.index(heading)
    tail = text[start + len(heading):]
    cut = tail.find("\n## ")
    return heading + (tail if cut == -1 else tail[:cut])


# ===========================================================================
# A. ONE OWNER, STILL
# ===========================================================================
def test_01_only_the_protection_owner_physically_protects_or_unprotects() -> None:
    """REQUIRED CONTROL 1, and REQUIRED CONTROL 14 with it. The reconciliation
    widened what may happen; it did not widen who may do it."""
    offenders = []
    for name in _all_modules():
        if name == "modProtection.bas":
            continue
        if re.search(r"\.(?:Protect|Unprotect)\b", _code(name)):
            offenders.append(name)
    assert not offenders, f"protection is spelled outside its owner: {offenders}"
    # NOT EVEN THE WORD, outside the owner. The envelope asks modProtection by
    # name; it never reaches an Excel object itself.
    for name in _all_modules():
        if name == "modProtection.bas":
            continue
        assert "Unprotect" not in _code(name), f"{name} unprotects something"


def test_02_the_envelope_asks_the_owner_rather_than_excel() -> None:
    body = _code("modAppState.bas")
    assert "modProtection.ProtectionBeginStructural(detail)" in body
    assert "modProtection.ProtectionEndStructural(protectionProblem)" in body
    # THE EXCEL MEMBER CALLS, not any dotted name containing them.
    # "modProtection.ProtectionBeginStructural" contains ".Protect".
    assert not re.search(r"\.(?:Protect|Unprotect)\b", body), (
        "the envelope reaches an Excel object instead of asking the owner")


def test_03_the_window_is_depth_counted_and_the_count_lives_in_the_owner() -> None:
    """REQUIRED CONTROL 2."""
    owner = _code("modProtection.bas")
    assert "Private mStructuralDepth As Long" in owner
    begin = _procedure("modProtection.bas", "ProtectionBeginStructural")
    end = _procedure("modProtection.bas", "ProtectionEndStructural")
    # NESTED BEGIN COUNTS AND RETURNS; it does not release again.
    assert "If mStructuralDepth > 0 Then" in begin
    assert "mStructuralDepth = mStructuralDepth + 1" in begin
    nested = begin[begin.index("If mStructuralDepth > 0 Then"):]
    nested = nested[: nested.index("End If")]
    assert "Unprotect" not in nested, "a nested open releases protection again"
    # NESTED END DECREMENTS AND RETURNS; it does not restore.
    assert "mStructuralDepth = mStructuralDepth - 1" in end
    assert "If mStructuralDepth > 0 Then" in end
    # THE NESTED BRANCH ITSELF, not "everything up to the first ProtectionApply"
    # - that slice ends where the thing it looks for begins, so it could never
    # find it, and a mutation proved the control was decoration.
    inner = end[end.index("If mStructuralDepth > 0 Then", end.index("- 1")):]
    inner = inner[: inner.index("End If")]
    assert "ProtectionApply" not in inner, "an inner close restores protection"
    assert "ProtectionEndStructural = True" in inner
    assert "Exit Function" in inner
    # AND NOBODY ELSE HOLDS A DEPTH.
    for name in _all_modules():
        if name == "modProtection.bas":
            continue
        assert "mStructuralDepth" not in _code(name), f"{name} tracks the depth too"


def test_04_closing_more_often_than_opening_is_refused() -> None:
    end = _procedure("modProtection.bas", "ProtectionEndStructural")
    assert "If mStructuralDepth <= 0 Then" in end
    assert "closed more often than it was opened" in end
    guard = end[end.index("If mStructuralDepth <= 0 Then"):]
    assert "mStructuralDepth = 0" in guard[: guard.index("End If")]


# ===========================================================================
# B. WHAT COMES BACK
# ===========================================================================
def test_10_the_outermost_close_restores_through_the_one_apply() -> None:
    """REQUIRED CONTROLS 3, 4 and 5. There is no second policy: the same
    ProtectionApply that establishes the resting state on open restores it."""
    end = _procedure("modProtection.bas", "ProtectionEndStructural")
    assert "If Not ProtectionApply(detail) Then Exit Function" in end
    apply_body = _procedure("modProtection.bas", "ProtectionApply")
    assert "UserInterfaceOnly:=True" in apply_body
    assert "ThisWorkbook.Protect Structure:=True, Windows:=False" in apply_body


def test_11_the_close_proves_it_took() -> None:
    """REQUIRED CONTROL 4. ProtectionApply returning True is what it TRIED."""
    end = _procedure("modProtection.bas", "ProtectionEndStructural")
    assert "If Not ProtectionIsApplied() Then" in end
    assert "does not report itself protected" in end
    applied = end.index("If Not ProtectionApply(detail) Then Exit Function")
    verified = end.index("If Not ProtectionIsApplied() Then")
    ok = end.rindex("ProtectionEndStructural = True")
    assert applied < verified < ok, "the close reports success before verifying it"


def test_12_workbook_structure_protection_is_never_released() -> None:
    """REQUIRED CONTROL 5. The 1004 named the SHEET. Releasing workbook structure
    would be wider than any evidence asks for."""
    begin = _procedure("modProtection.bas", "ProtectionBeginStructural")
    assert "ThisWorkbook.Unprotect" not in begin, (
        "the structural window releases workbook structure protection")
    assert "ProtectStructure" not in begin
    # The only ThisWorkbook.Unprotect in the module is the maintenance path.
    release = _procedure("modProtection.bas", "ProtectionRelease")
    assert "ThisWorkbook.Unprotect" in release
    assert _code("modProtection.bas").count("ThisWorkbook.Unprotect") == 1


def test_13_the_open_proves_the_release_before_work_begins() -> None:
    """REQUIRED CONTROL 8. Beginning structural work on a still-protected sheet
    would produce the very 1004 this removes, and would look like the defect."""
    begin = _procedure("modProtection.bas", "ProtectionBeginStructural")
    assert "worksheet protection was not released from" in begin
    assert begin.count("For Each sheet In ThisWorkbook.Worksheets") == 2, (
        "the release is not verified by a second pass")


def test_14_a_half_open_window_re_protects_immediately() -> None:
    """THE WORST OUTCOME THERE IS: some sheets released, no depth recorded, and
    therefore nobody who will ever close it."""
    begin = _procedure("modProtection.bas", "ProtectionBeginStructural")
    failed = begin[begin.index("Failed:"):]
    assert "mStructuralDepth = 0" in failed
    assert "If Not ProtectionApply(reapply) Then" in failed
    assert "could not be re-applied either" in failed


def test_15_a_failed_open_stops_the_operation() -> None:
    body = _procedure("modAppState.bas", "BeginStructuralOperation")
    assert "If Not modProtection.ProtectionBeginStructural(detail) Then" in body
    assert "Err.Raise" in body
    assert "structural work was attempted" in body
    assert "Snapshot.Structural = True" in body
    assert body.index("Err.Raise") < body.index("Snapshot.Structural = True"), (
        "the snapshot is marked structural even when the window did not open")


# ===========================================================================
# C. WHO ASKS, AND WHO DOES NOT
# ===========================================================================
def test_20_every_structural_command_owner_requests_the_window() -> None:
    """REQUIRED CONTROL 6."""
    for owner in STRUCTURAL_OWNERS:
        body = _code(f"{owner}.bas")
        assert "modAppState.BeginStructuralOperation" in body, (
            f"{owner} performs structural work without requesting the window")
        assert "modAppState.BeginOperation" not in body, (
            f"{owner} still opens a non-structural envelope")


def test_21_no_non_structural_command_requests_it() -> None:
    """REQUIRED CONTROLS 7 and 21. A stochastic run holds the workbook for
    minutes; leaving every sheet unprotected for that is not a side effect
    anybody asked for."""
    for owner in NON_STRUCTURAL_OWNERS:
        body = _code(f"{owner}.bas")
        assert "BeginStructuralOperation" not in body, (
            f"{owner} opens a structural window it has no structural work for")
        assert "modAppState.BeginOperation" in body


def test_22_the_structural_set_is_exactly_the_commands_that_need_it() -> None:
    """CONFIRMED FROM SOURCE, not hard-coded on the batch note. A command is
    structural when it, or the rollback it owns, can reach a ListObject
    structural operation."""
    structural_sites = re.compile(
        r"\.ListRows\.Add|\.ListRows\([^)]*\)\.Delete|"
        r"\.ListColumns\.Add|\.ListColumns\([^)]*\)\.Delete")
    # The helpers that perform them, and the module each belongs to.
    performers = {name for name in _all_modules() if structural_sites.search(_code(name))}
    assert performers == {"modCalcReport.bas", "modDrivers.bas", "modInflation.bas",
                          "modProfiling.bas", "modWorkbook.bas"}, sorted(performers)
    # modWorkbook.RestoreTable performs all four and IS the rollback, so every
    # command that calls it is structural by its rollback alone.
    rollback_callers = {name for name in _all_modules()
                        if name != "modWorkbook.bas"
                        and "modWorkbook.RestoreTable" in _code(name)}
    assert rollback_callers == {f"{owner}.bas" for owner in STRUCTURAL_OWNERS}, (
        sorted(rollback_callers))
    # AND NO NON-STRUCTURAL OWNER REACHES ONE, directly or through a helper.
    for owner in NON_STRUCTURAL_OWNERS:
        body = _code(f"{owner}.bas")
        assert not structural_sites.search(body), f"{owner} mutates a ListObject"
        assert "modWorkbook.RestoreTable" not in body, f"{owner} owns a table rollback"
        for helper in ("modProfiling.SetYearColumns", "modProfiling.SyncRows",
                       "modProfiling.RemoveRow", "modInflation.SetYearColumns",
                       "modInflation.SyncProfileRows"):
            assert helper not in body, f"{owner} calls the structural helper {helper}"


def test_23_every_named_endpoint_lives_where_this_suite_says() -> None:
    for endpoint, owner in {**STRUCTURAL_COMMANDS, **NON_STRUCTURAL_COMMANDS}.items():
        assert re.search(rf"^Public Sub {endpoint}\(\)", _src(f"{owner}.bas"), re.M), (
            f"{endpoint} is not a public Sub of {owner}")


# ===========================================================================
# D. CLOSING, ON EVERY PATH
# ===========================================================================
def test_30_the_close_is_in_the_shared_cleanup_every_path_already_uses() -> None:
    """REQUIRED CONTROLS 10, 11 and 12 at once, and that is the point of putting
    it here: success, refusal, validation error, runtime error and injected
    failure all route through FinishOperation already."""
    body = _procedure("modAppState.bas", "FinishOperation")
    assert "If Snapshot.Structural Then" in body
    assert "modProtection.ProtectionEndStructural(protectionProblem)" in body
    # AND THE FLAG IS CLEARED FIRST, so a path that reaches cleanup twice cannot
    # decrement the owner's depth twice.
    cleared = body.index("Snapshot.Structural = False")
    closed = body.index("ProtectionEndStructural")
    assert cleared < closed, "the flag is cleared after the close, not before"


# The procedure that holds the envelope for each structural owner, the call
# inside it that can perform or reach a rollback, and the cleanup that closes the
# window. NAMED, because the two shapes differ and a generic scan would prove
# nothing: modTimeline and modDrivers roll back inline in their Failure handler,
# while modCalcReport and modRepair delegate to a worker that rolls back and
# returns an OperationResult.
# `section` is the label the check starts from, because the SUCCESS path also
# calls the cleanup and calls it earlier in the text. modTimeline and modDrivers
# roll back inside their Failure handler, so that handler is the section; the
# other two delegate to a worker that rolls back on the main path, so the whole
# body is.
ROLLBACK_BEFORE_CLEANUP = (
    ("modTimeline", "PCCM_ApplyTimeline", "\nFailure:",
     "TryRestoreTimeline(", "FinishIfCaptured("),
    ("modDrivers", "RunDriverOperation", "\nFailure:",
     "TryRestoreDriver(", "FinishIfCaptured("),
    ("modCalcReport", "PCCM_Calculate", None,
     "result = RunCalculation(", "modAppState.FinishOperation("),
    ("modRepair", "PCCM_RepairProfiling", None,
     "result = RepairProfiling()", "modAppState.FinishOperation("),
)


@pytest.mark.parametrize("owner,entry,section,rollback,cleanup", ROLLBACK_BEFORE_CLEANUP)
def test_31_the_window_covers_the_rollback(owner: str, entry: str, section: str,
                                           rollback: str, cleanup: str) -> None:
    """REQUIRED CONTROL 9, and the one ordering that must not be got wrong.

    A rollback that rebuilds a table needs the window the forward mutation used.
    Because the close lives in FinishOperation, the requirement reduces to this:
    inside the procedure that owns the envelope, the rollback-capable call comes
    before every cleanup call. File order elsewhere proves nothing - the actual
    RestoreTable statements sit in private helpers hundreds of lines further
    down, and an earlier draft of this control was fooled by exactly that.
    """
    body = _procedure(f"{owner}.bas", entry)
    # THE WINDOW IS OPENED FIRST, on the whole body.
    assert "modAppState.BeginStructuralOperation" in body
    opened = body.index("modAppState.BeginStructuralOperation")

    scope = body if section is None else body[body.index(section):]
    assert rollback in scope, f"{owner}.{entry} no longer reaches a rollback"
    assert cleanup in scope, f"{owner}.{entry} no longer routes through the cleanup"
    assert scope.index(rollback) < scope.index(cleanup), (
        f"{owner}.{entry} closes the protection window before its rollback runs")
    if section is None:
        assert opened < body.index(rollback)


def test_32_every_owner_still_routes_all_its_cleanup_through_the_envelope() -> None:
    for owner in STRUCTURAL_OWNERS + NON_STRUCTURAL_OWNERS:
        body = _code(f"{owner}.bas")
        assert "FinishOperation" in body, f"{owner} has no shared cleanup"
        # NO LOCAL PROTECTION HANDLING ANYWHERE. The envelope owns it.
        assert "ProtectionEndStructural" not in body, (
            f"{owner} closes the window itself instead of through the envelope")
        assert "ProtectionBeginStructural" not in body, (
            f"{owner} opens the window itself instead of through the envelope")


def test_33_a_restoration_failure_cannot_be_an_ordinary_success() -> None:
    """REQUIRED CONTROL 13. The business mutation may well have committed; that
    does not make an unprotected workbook a success."""
    body = _procedure("modAppState.bas", "FinishOperation")
    assert 'problems = problems & "  PROTECTION WAS NOT RESTORED: "' in body
    assert "FinishOperation = problems" in body
    # AND EVERY OWNER ALREADY TREATS A NON-EMPTY CLEANUP AS A FAILURE.
    for owner in STRUCTURAL_OWNERS:
        assert re.search(r"If Len\(cleanup\) > 0 Then", _code(f"{owner}.bas")), owner


# ===========================================================================
# E. WHAT MUST NOT HAVE MOVED
# ===========================================================================
def test_40_workbook_open_still_establishes_the_protected_resting_state() -> None:
    """REQUIRED CONTROL 15. Untouched, and a control says so."""
    current = (SRC / "ThisWorkbook.vba").read_bytes()
    accepted = subprocess.run(["git", "show", "58b2394:pccm/src/vba/ThisWorkbook.vba"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout
    assert current == accepted, "the open handler moved"
    text = current.decode("utf-8")
    assert "modProtection.ProtectionApply(detail)" in text
    assert "BeginStructuralOperation" not in text
    assert "ProtectionBeginStructural" not in text


def test_41_command_feedback_is_unchanged() -> None:
    """REQUIRED CONTROL 16. OperationResult, Announce, ReportResult, the MsgBox
    behaviour and the automation hooks are byte-identical after the reversal."""
    from vba_structural_window import strip_structural_window

    current = strip_structural_window("modAppState.bas", _src("modAppState.bas"))
    accepted = subprocess.run(["git", "show", "58b2394:pccm/src/vba/modAppState.bas"],
                              cwd=REPO_ROOT, check=True,
                              stdout=subprocess.PIPE).stdout.decode("utf-8")
    assert current == accepted, "modAppState moved outside the declared window"


def test_42_the_reconciliation_is_the_only_production_change() -> None:
    """AND IT IS PROVED BY REVERSAL, not asserted."""
    from vba_structural_window import (ACCEPTED_BEFORE_RECONCILIATION,
                                       DECLARED_STRUCTURAL_WINDOW_CHANGES,
                                       strip_structural_window)

    changed = [line for line in subprocess.run(
        ["git", "diff", "--name-only", ACCEPTED_BEFORE_RECONCILIATION, "--",
         "pccm/src", "pccm/spec"], cwd=REPO_ROOT, check=True,
        stdout=subprocess.PIPE, text=True).stdout.splitlines() if line.strip()]
    declared = {f"pccm/src/vba/{name}" for name in DECLARED_STRUCTURAL_WINDOW_CHANGES}
    assert set(changed) <= declared, sorted(set(changed) - declared)
    for name in DECLARED_STRUCTURAL_WINDOW_CHANGES:
        accepted = subprocess.run(
            ["git", "show", f"{ACCEPTED_BEFORE_RECONCILIATION}:pccm/src/vba/{name}"],
            cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout.decode("utf-8")
        assert strip_structural_window(name, _src(name)) == accepted, name


@pytest.mark.parametrize("owner,markers", [
    ("modTimeline", ("TryRestoreTimeline", "FailPointCheck", "captured = True",
                     "modAppState.AskConfirm")),
    ("modDrivers", ("modWorkbook.RestoreTable", "FailPointCheck", "modAppState.AskConfirm")),
    ("modCalcReport", ("modWorkbook.RestoreTable", "PrepareCurrentCalculation")),
    ("modRepair", ("modWorkbook.RestoreTable", "modProfiling.SetYearColumns")),
])
def test_43_the_transactional_semantics_are_untouched(owner: str, markers: tuple) -> None:
    """REQUIRED CONTROLS 17, 18, 19 and 20. The window is a protection change; it
    is not a change to what any command does, refuses, or rolls back."""
    body = _code(f"{owner}.bas")
    for marker in markers:
        # WORD-BOUNDED. `TryRestoreTimelineDisabled` contains
        # `TryRestoreTimeline`, so a substring check calls a rollback present
        # that no longer runs - a mutation walked straight through it.
        assert re.search(re.escape(marker) + r"(?!\w)", body), f"{owner} lost {marker}"


def test_44_no_password_appeared_anywhere() -> None:
    owner = _code("modProtection.bas")
    assert "Password" not in owner, "a password argument appeared"
    for name in _all_modules():
        assert "Password:=" not in _code(name), f"{name} passes a password"


def test_45_no_user_facing_unprotected_mode_exists() -> None:
    """REQUIRED: the workbook must not expose an ordinary interactive
    unprotected mode. The window is code-only and no button opens it."""
    structure = (PCCM_ROOT / "spec" / "structure_contract.yaml").read_text(encoding="utf-8")
    for banned in ("ProtectionRelease", "ProtectionBeginStructural", "PCCM_Unprotect"):
        assert banned not in structure, f"{banned} is bound to a button"
    for name in _all_modules():
        assert not re.search(r"^Public Sub PCCM_\w*[Pp]rotect", _src(name), re.M), name


# ===========================================================================
# F. THE RECORD
# ===========================================================================
def test_50_the_windows_evidence_is_recorded_accurately() -> None:
    """REQUIRED CONTROL 24. Both halves of what the run established, and the
    original assumption kept rather than rewritten."""
    text = EVIDENCE.read_text(encoding="utf-8")
    assert "Table features aren't available because the sheet is protected" in text
    assert "invoked" in text and "REFUSED" in text
    # THE VALUE-WRITE CAPABILITY IS STILL RECORDED AS CONFIRMED.
    assert "UserInterfaceOnly" in text
    assert "CONFIRMED" in text
    # AND THE ORIGINAL DECISION IS NOT REWRITTEN.
    assert "6ab8f6a" in text
    assert "Original assumption" in text or "original assumption" in text


def test_52_the_run_6_windows_evidence_is_recorded_accurately() -> None:
    """THE RUN THAT EXERCISED THE RECONCILIATION, recorded with both halves.

    ApplyTimeline SUCCEEDED and did real structural work with protection intact,
    and Calculate refused for a NON-protection reason. Recording only the first
    would be an acceptance claim this run does not support.
    """
    # SCOPED TO ITS OWN SECTION. The record now carries a later run that observed
    # the same shapes, so a document-wide count would be satisfied by the wrong
    # run's numbers.
    text = _evidence_section("## Protection probe Run 6")
    # THE SUCCESS, AND THE ACTUAL SHAPE CHANGE - ALL THREE GRIDS. Two of them
    # moved 25×2 → 25×5, so requiring the string once is satisfied by either one
    # alone; a mutation that deleted one walked through this until it was counted.
    assert text.count("25×2 → 25×5") == 2, (
        "both profiling grids' observed structural effect must be recorded")
    assert "10×1 → 10×4" in text, "the inflation grid's structural effect is missing"
    assert "invoked and SUCCEEDED" in text
    # THE STRUCTURE FLAG STAYED TRUE - the open question this run settled.
    assert "ProtectStructure = True` throughout" in text
    assert "structure = **True**" in text
    # A-E PROVEN, AND THE VERDICT STILL INCONCLUSIVE.
    assert "PROVEN" in text and "INCONCLUSIVE" in text
    # THE REFUSAL, QUOTED, AND NAMED AS NOT ABOUT PROTECTION.
    assert "Discount Rate: the value is blank. A blank is not zero." in text
    assert "non-protection" in text
    # AND THE ORDERING CAVEAT IS NOT QUIETLY DROPPED.
    assert "python3" in text and "Stage A" in text
    assert "not ideal acceptance evidence" in text or \
        "not** ideal acceptance evidence" in text
    # WHAT IT DID NOT ESTABLISH IS STATED AS PLAINLY AS WHAT IT DID.
    assert "capacity expansion was not" in text.lower()


def test_52a_run_7_is_recorded_with_its_coverage_gap_stated() -> None:
    """RUN 7 REPORTED "PRODUCTION IS FINE UNDER PROTECTION", and that was honest
    against the criteria of the day. It proved the ADD direction only.

    Benchmark Run 3 died on a ListRow.Delete(), so the broad "HARNESS defect
    only" conclusion is PENDING delete-path evidence - and the record has to say
    so, or the next reader takes Run 7 for a closure it is not.
    """
    text = _evidence_section("## Protection probe Run 7")
    # WHAT IT PROVED.
    assert "1×3 → 3×3" in text and "1×8 → 3×8" in text
    assert "351 passed, 0 failed" in text
    assert "PRODUCTION IS FINE UNDER PROTECTION" in text
    # AND WHAT IT DID NOT, said as plainly.
    assert "Not proved" in text
    assert "ListRows.Delete" in text
    assert "ListColumns.Delete" in text
    assert "remains PENDING exact delete-path runtime evidence" in text
    assert '"HARNESS' in text and "defect only" in text
    # THE GAP IS A GAP IN THE CRITERIA, NOT AN ERROR IN THE RUN.
    # THE SENTENCE SPANS A LINE BREAK IN THE RECORD, so it is matched in pieces.
    assert "coverage gap in the criteria" in text and "not as an error in" in text
    # AND THIS ROUND CLAIMS NO RUNTIME EVIDENCE OF ITS OWN.
    assert "The delete path is unproved on Windows" in text


def test_52b_the_delete_coverage_is_not_claimed_before_windows_proves_it() -> None:
    """THE RECORD MUST NOT SAY THE DELETE PATH IS SETTLED. Nothing in this round
    is runtime evidence; it is a probe that can now ask the question."""
    text = EVIDENCE.read_text(encoding="utf-8")
    for overclaim in ("ListRows.Delete: OBSERVED on Windows",
                      "the delete path is proved",
                      "delete-path runtime evidence obtained",
                      "Benchmark Run 3 was a HARNESS defect only"):
        assert overclaim not in text, f"the record claims delete coverage it does not have: {overclaim}"


def test_52c_run_8_is_recorded_as_a_probe_regression_not_a_production_failure() -> None:
    """RUN 8 ENDED BEFORE ANY ENDPOINT RAN. Recording it as anything else would
    put a production defect on the record that no run established."""
    text = _evidence_section("## Protection probe Run 8")
    # WHAT SUCCEEDED, INCLUDING THE ORDERING THAT WAS RIGHT THIS TIME.
    assert "351 passed, 0 failed" in text
    assert "rebuilt BEFORE the probe" in text
    assert "Stage-B bootstrap completed" in text
    # THE EXACT FAILURE.
    assert "PropertyNotFoundException" in text
    assert "The property 'ProtectContents' cannot be found on this object." in text
    assert "NOT ATTEMPTED" in text and "INCONCLUSIVE" in text
    assert "Shutdown was clean." in text
    # AND WHAT IT DOES NOT ESTABLISH.
    assert "not a production failure" in text
    assert "No production defect is established" in text
    assert "No delete-path evidence was obtained" in text
    assert "remains **OPEN**" in text
    # RUN 7 IS NOT REWRITTEN.
    assert "Run 7 remains valid" in text


def test_52d_the_run_8_root_cause_separates_proof_from_inference() -> None:
    """THE BYTE-IDENTICAL FINDING IS PROOF; THE COM MECHANISM IS NOT. Labelling
    one as the other is how a plausible story becomes a settled cause - which is
    exactly what went wrong at probe Run 2."""
    text = _evidence_section("## Protection probe Run 8")
    assert "PROVED FROM SOURCE" in text
    assert "byte-identical" in text
    assert "INFERRED, AND LABELLED AS SUCH" in text
    assert "cannot be proved from Linux" in text
    assert "The delete-path work did not introduce this" in text


def test_52e_run_8_claims_no_delete_evidence() -> None:
    """REQUIRED: no claiming delete coverage from a run that reached no
    endpoint."""
    text = _evidence_section("## Protection probe Run 8")
    for overclaim in ("LISTROW DELETE    : OBSERVED", "delete path is proved",
                      "ListRows.Delete: OBSERVED"):
        assert overclaim not in text, overclaim
    assert "Nothing here bears on" in text


def test_52f_run_9_separates_valid_evidence_from_a_defective_predicate() -> None:
    """TWO DIFFERENT THINGS, RECORDED AS TWO DIFFERENT THINGS. The four classes
    were genuinely OBSERVED; the acceptance predicate was wrong. Collapsing them
    either way would misstate the run."""
    text = _evidence_section("## Protection probe Run 9")
    assert "Runtime production evidence — VALID" in text
    assert "Acceptance / reporting predicate — DEFECTIVE" in text
    # THE FOUR CLASSES, WITH THE OBSERVED SHAPES.
    for shape in ("25×2 → 25×5", "10×1 → 10×4", "25×5 → 25×3", "10×4 → 10×2",
                  "1×3 → 3×3", "1×8 → 3×8", "3×3 → 1×3", "3×8 → 1×8"):
        assert shape in text, shape
    # THE DELETE CLASS IS NAMED AS THE ONE BENCHMARK RUN 3 FAILED ON.
    assert "same `ListRows.Delete` class that Benchmark Run 3" in text
    # PROTECTION HELD THROUGHOUT.
    assert "before and after every invoked endpoint" in " ".join(text.split())
    assert "No protection 1004 occurred" in text


def test_52g_the_run_9_record_does_not_pretend_c_printed_met() -> None:
    """REQUIRED: do not rewrite the historical output."""
    text = _evidence_section("## Protection probe Run 9")
    assert "C. NOT MET" in text
    assert "STRUCTURAL INITIALISATION: NOT PROVEN - C not met." in text
    assert "did print NOT MET, and this record does not pretend otherwise" in text


def test_52h_the_run_9_root_cause_is_the_selector_not_mutable_state() -> None:
    """AND THE RECORD SAYS WHICH. "It read the workbook after the shrink" is a
    plausible story that happens to be wrong, and a wrong root cause is how the
    same defect comes back somewhere else."""
    text = _evidence_section("## Protection probe Run 9")
    assert "selected its round by **endpoint name**" in text
    assert "never reading post-shrink workbook state" in text
    assert "append-only" in text and "immutable" in text
    assert "computed **after**" in text


def test_52i_run_9_does_not_reopen_the_protection_architecture() -> None:
    """REQUIRED: a reporting defect is not a reason to widen the envelope."""
    text = _evidence_section("## Protection probe Run 9")
    assert "protection architecture is not reopened" in text.lower()
    assert "no privilege-envelope expansion" in text.lower()
    assert "structure protection is still never released" in text
    # AND PRODUCTION REALLY IS UNCHANGED.
    changed = subprocess.run(
        ["git", "diff", "--name-only", "04fcf82", "--", "pccm/src", "pccm/spec"],
        cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE, text=True).stdout
    assert not changed.strip(), f"production changed: {changed}"


def test_53_the_runtime_evidence_for_keeping_structure_protected_is_recorded() -> None:
    """RUN 6 TURNED AN INFERENCE INTO EVIDENCE. The reconciliation batch left
    workbook-structure protection applied on a reading of Excel's 1004 wording;
    this run proved ListColumns.Add does not need it released. The envelope is
    not widened, and the reason is on the record rather than in a commit
    message."""
    text = EVIDENCE.read_text(encoding="utf-8")
    assert "privilege envelope is not widened" in text.lower() or \
        "not widened" in text.lower()
    assert "ListColumns.Add" in text
    # AND THE CODE STILL REFUSES TO RELEASE IT.
    begin = _procedure("modProtection.bas", "ProtectionBeginStructural")
    assert "ThisWorkbook.Unprotect" not in begin


def test_54_production_vba_is_byte_identical_to_the_reconciliation() -> None:
    """THIS ROUND IS SOURCE/STATIC AND PROBE-ONLY. Windows exercised 0946cf6, so
    a production edit here would mean the evidence above describes code that no
    longer exists."""
    changed = [line for line in subprocess.run(
        ["git", "diff", "--name-only", "0946cf6", "--", "pccm/src", "pccm/spec"],
        cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE, text=True).stdout.splitlines()
        if line.strip()]
    assert changed == [], f"production changed after the run that exercised it: {changed}"


def test_51_the_owner_records_what_was_disproved_and_what_was_not() -> None:
    text = _src("modProtection.bas")
    assert "6ab8f6a" in text, "the historical authority is not cited"
    assert "Error 1004" in text
    assert "PERMITTED" in text and "NOT PERMITTED" in text
    assert "historical evidence" in text
    # THE SENTENCE THAT REFUSES TO REWRITE THE ORIGINAL DECISION. Losing it is
    # how a reconciliation quietly becomes "it was always like this".
    assert "AND IT IS NOT BEING REWRITTEN" in text
    assert "it was WRONG about one" in text


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
