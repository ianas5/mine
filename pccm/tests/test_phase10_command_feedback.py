#!/usr/bin/env python3
"""P10-2A.1: the four operational buttons tell a human what happened.

WHAT THIS AUDIT FOUND, AND IT IS NOT WHAT THE PREVIOUS RETURN SAID. The P10-2A
return listed "Calculate and Run Simulation tell the user nothing" as an open
item. That was WRONG, and it was wrong for a specific and instructive reason: it
came from grepping each module for MsgBox and finding none. The dialog is not in
those modules. Every one of the four endpoints publishes its outcome through
`modAppState.Announce`, which records the result for automation ALWAYS and calls
ReportResult ONLY when automation is inactive - which is the human-feedback
contract, implemented since Phase 5 and proved on Windows by every Gate-B run.

SO THIS BATCH CHANGES NO PRODUCTION BYTE. What was actually missing is the
control: nothing asserted that the four endpoints announce, that they announce
exactly once, that the wording is the endpoint's own, or that a dialog can never
reach a harness. Those are the assertions here.

WHY THAT MATTERS MORE THAN A CHANGE WOULD HAVE. An unproved behaviour that
happens to be right is one edit away from being wrong, and the edit that broke
it would have looked like an improvement.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
SRC = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

from vba_reset_plumbing import strip_reset_addition  # noqa: E402

# THE COMMIT THIS BATCH MAY NOT MOVE. P10-2A is accepted; a feedback batch that
# edited an endpoint's logic or a button's binding would be a different batch.
ACCEPTED = "21a2774"

# (endpoint, owner, the operation label its failures carry)
ENDPOINTS = (
    ("PCCM_Calculate", "modCalcReport", '"Calculate"'),
    ("PCCM_RunSimulation", "modSimReport", '"Run Simulation"'),
    ("PCCM_RunSensitivity", "modSimPostReport", '"Run Sensitivity"'),
    ("PCCM_RunAnnualStochastic", "modSimAnnualRun", "ANNUAL_OPERATION"),
)

OWNERS = tuple(owner for _e, owner, _l in ENDPOINTS)


def _source(module: str) -> str:
    return (SRC / f"{module}.bas").read_text(encoding="utf-8")


def _code(module: str) -> str:
    """Executable VBA only - a claim must not be satisfied by the comment
    explaining it."""
    return "\n".join(line for line in _source(module).splitlines()
                     if not line.strip().startswith("'"))


def _endpoint_body(endpoint: str, module: str) -> str:
    code = _code(module)
    start = code.index(f"Public Sub {endpoint}()")
    return code[start:code.index("\nEnd Sub", start)]


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                          stdout=subprocess.PIPE, text=True).stdout


# ===========================================================================
# A. AUTOMATION IS UNTOUCHED
# ===========================================================================
def test_01_announce_is_the_one_gate_between_a_result_and_a_dialog() -> None:
    """A. The whole contract, in one accepted owner, in five lines.

    RecordResult runs unconditionally so a harness always has the outcome;
    ReportResult runs only when automation is inactive so a harness never meets
    a modal dialog nobody is there to dismiss.
    """
    body = _code("modAppState")
    announce = body[body.index("Public Sub Announce("):]
    announce = announce[:announce.index("\nEnd Sub")]
    assert "RecordResult" in announce, "the outcome is not recorded for automation"
    assert "If Not gAutomationActive Then ReportResult Result" in announce, (
        "a dialog is not gated on automation")
    # THE GUARD COMES AFTER THE RECORD, so an automated run is never left with
    # no recorded outcome because a dialog decision short-circuited first.
    assert announce.index("RecordResult") < announce.index("gAutomationActive")
    # AND THE RECORDED STRING DISTINGUISHES THE TWO OUTCOMES, which is what a
    # harness reads instead of a dialog.
    assert '"OK|"' in announce and '"FAIL|"' in announce


@pytest.mark.parametrize("endpoint,module,_label", ENDPOINTS)
def test_02_every_endpoint_publishes_only_through_announce(endpoint: str, module: str,
                                                           _label: str) -> None:
    """A. No endpoint reaches a dialog by any other route."""
    body = _endpoint_body(endpoint, module)
    assert "modAppState.Announce" in body, f"{endpoint} announces nothing"
    for direct in ("MsgBox", "ReportResult", "ReportFailure"):
        assert direct not in body, (
            f"{endpoint} reaches {direct} directly, bypassing the automation gate")


@pytest.mark.parametrize("module", OWNERS)
def test_03_no_owner_touches_the_automation_flag(module: str) -> None:
    """A. An endpoint that cleared or set the flag could show a dialog to a
    harness, or hide one from a human."""
    code = _code(module)
    assert "gAutomationActive =" not in code, f"{module} assigns the automation flag"
    assert "ClearAutomation" not in code, f"{module} clears automation state"


# ===========================================================================
# B. HUMAN SUCCESS - ONE MESSAGE, AND IT IS THE ENDPOINT'S OWN
# ===========================================================================
@pytest.mark.parametrize("module,fragment", [
    ("modCalcReport", "Calculation committed."),
    ("modSimReport", "Simulation complete. "),
    ("modSimPostReport", "Sensitivity complete: "),
    ("modSimAnnualRun", "Annual stochastic complete: "),
])
def test_04_each_success_message_has_exactly_one_owner(module: str,
                                                       fragment: str) -> None:
    """B and section 4. The wording lives with the operation that earned it, and
    is spelled ONCE - a second copy anywhere is a second owner of the sentence a
    user reads."""
    assert _code(module).count(fragment) == 1, (
        f"{fragment!r} is spelled {_code(module).count(fragment)} times in {module}")
    elsewhere = [other for other in _all_modules() if other != module
                 and fragment in _code(other)]
    assert not elsewhere, f"{fragment!r} is duplicated into {elsewhere}"


def _all_modules() -> list[str]:
    return sorted(path.stem for path in SRC.glob("*.bas"))


@pytest.mark.parametrize("endpoint,module,_label", ENDPOINTS)
def test_05_the_success_result_reaches_announce_unaltered(endpoint: str, module: str,
                                                          _label: str) -> None:
    """B. The endpoint hands its own result object to Announce. Nothing rebuilds
    a message, and nothing recalculates to describe what just happened."""
    body = _endpoint_body(endpoint, module)
    assert re.search(r"^\s*modAppState\.Announce result$", body, re.M), (
        f"{endpoint} does not announce its own result on the success path")
    # NO RECALCULATION TO BUILD A SENTENCE.
    for banned in ("PrepareCurrentCalculation", "DeriveStatus", "DeriveSimStatus"):
        assert banned not in body, (
            f"{endpoint} re-derives state while reporting: {banned}")


# ===========================================================================
# C-D. HUMAN FAILURE - ONE MESSAGE, THE AUTHORITATIVE ONE, NEVER BOTH
# ===========================================================================
def test_06_the_failure_branch_carries_the_reason_and_the_detail() -> None:
    """C. ReportResult is one procedure with two branches, and the failure
    branch shows the owner's reason AND its detail. A user told only that
    something failed has not been told anything."""
    body = _code("modAppState")
    report = body[body.index("Public Sub ReportResult("):]
    report = report[:report.index("\nEnd Sub")]
    assert "If Result.Ok Then" in report
    assert "vbInformation" in report and "vbExclamation" in report
    assert "Result.Detail" in report, "a refusal's detail never reaches the user"
    assert report.index("vbInformation") < report.index("vbExclamation")


@pytest.mark.parametrize("endpoint,module,label", ENDPOINTS)
def test_07_every_failure_path_announces_a_failed_result(endpoint: str, module: str,
                                                         label: str) -> None:
    """C. Each handler builds a Failed result with the operation's own label,
    rather than a bare string or a status word parsed back out of a cell."""
    body = _endpoint_body(endpoint, module)
    handlers = [line for line in body.splitlines()
                if line.strip().endswith(":") and not line.startswith(" ")]
    assert handlers, f"{endpoint} has no error handler"
    assert f"modAppState.Failed({label}" in body, (
        f"{endpoint} does not label its failures with {label}")
    # AND NO STATUS STRING IS PARSED to decide that something failed.
    for parsed in ('"INVALID"', '"STALE"', '"NOT CALCULATED"'):
        assert parsed not in body, f"{endpoint} infers failure from a status word"


@pytest.mark.parametrize("endpoint,module,_label", ENDPOINTS)
def test_08_exactly_one_announcement_can_run_per_invocation(endpoint: str, module: str,
                                                            _label: str) -> None:
    """D. THE ONE THING THAT WOULD SHOW A USER TWO DIALOGS. Every announcement
    is immediately followed by an exit, so no path can fall through into the
    handler beneath it - and a success followed by a failure message is exactly
    the shape that would confuse a user about whether their work was saved.
    """
    body = _endpoint_body(endpoint, module)
    lines = body.splitlines()
    announces = [i for i, line in enumerate(lines) if "modAppState.Announce" in line]
    assert len(announces) >= 2, f"{endpoint} has no error announcements"
    for index in announces:
        # THE STATEMENT MAY BE CONTINUED, so walk to the end of it first.
        end = index
        while lines[end].rstrip().endswith("_"):
            end += 1
        following = [line.strip() for line in lines[end + 1:] if line.strip()]
        assert following[:1] == ["Exit Sub"] or end == len(lines) - 1, (
            f"{endpoint}: an announcement at line {index} is not the last act of "
            f"its path; it falls through to {following[:1]}")


# ===========================================================================
# E-F. NOTHING ELSE MOVED
# ===========================================================================
@pytest.mark.parametrize("module", OWNERS)
def test_09_no_endpoint_owner_changed_since_the_accepted_batch(module: str) -> None:
    """E. THE STRONGEST FORM THIS CLAIM CAN TAKE. This batch reports outcomes; it
    does not touch calculation, simulation, sensitivity, annual replay,
    fingerprints, RNG, publication, state derivation or attempt history. Rather
    than list those and hope, the four owners are required to be byte-identical
    to the accepted P10-2A tree.

    CORRECTED AT P10-2B, AND NOT LOOSENED. Three of these four owners have since
    gained a narrow Reset Results clear/restore pair, appended as one declared
    block each. The requirement is no weaker: the block is taken back OUT
    mechanically and the accepted bytes must come back exactly, so anything that
    rode along above it still fails here. modSimAnnualRun gained nothing, and the
    same reversal leaves it untouched - which is why one control still covers all
    four rather than three plus an exception.
    """
    current = (SRC / f"{module}.bas").read_bytes().decode("utf-8")
    accepted = subprocess.run(
        ["git", "show", f"{ACCEPTED}:pccm/src/vba/{module}.bas"],
        cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout.decode("utf-8")
    assert strip_reset_addition(current) == accepted, (
        f"{module}.bas moved outside the declared P10-2B reset block")


def test_10_the_reporting_owner_did_not_change_either() -> None:
    """E. modAppState is shared by every Phase-4 structural command. Changing
    Announce or ReportResult to suit four endpoints would change all of them."""
    current = (SRC / "modAppState.bas").read_bytes()
    accepted = subprocess.run(["git", "show", f"{ACCEPTED}:pccm/src/vba/modAppState.bas"],
                              cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout
    assert current == accepted, "modAppState.bas moved"


def test_11_msgbox_lives_in_exactly_one_module() -> None:
    """F. Any new dialog outside the accepted owner would be a second reporting
    mechanism, and the second one is always the one that forgets the guard."""
    owners = [name for name in _all_modules() if "MsgBox" in _code(name)]
    assert owners == ["modAppState"], f"MsgBox appears outside its owner: {owners}"
    # AND THE DOCUMENT MODULE HAS NONE EITHER.
    event = (SRC / "ThisWorkbook.vba").read_text(encoding="utf-8")
    event = "\n".join(line for line in event.splitlines()
                      if not line.strip().startswith("'"))
    assert "MsgBox" not in event, "the open handler shows its own dialog"


# ===========================================================================
# G. THE BUTTONS ARE THE ACCEPTED ONES
# ===========================================================================
def test_12_no_button_binding_moved_in_this_batch() -> None:
    """G. A feedback batch that quietly rebound a button would be changing what
    the user runs, not how they are told about it.

    CORRECTED AT P10-2B. The contract has since gained the Reset Results button,
    its entry point and its module, so it is no longer byte-identical to the
    accepted tree. The claim is unchanged and is now made where it can still be
    made exactly: every accepted button, entry point and module is compared
    field for field against P10-2A, and the DELTA is named. A rebinding, a moved
    anchor, a renamed shape or a second undeclared addition all still fail.
    """
    import yaml

    current = yaml.safe_load((SPEC / "structure_contract.yaml").read_text(encoding="utf-8"))
    accepted = yaml.safe_load(subprocess.run(
        ["git", "show", f"{ACCEPTED}:pccm/spec/structure_contract.yaml"],
        cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout.decode("utf-8"))

    def by_key(doc, path, key):
        node = doc
        for step in path:
            node = node[step]
        return {item[key]: item for item in node}

    # CORRECTED AT P10-UX, AND THE BINDING HALF IS UNWEAKENED. A human opened
    # the workbook and found Apply / Update Timeline drawn across the paragraph
    # explaining the Applied Timeline block; it now leads the command block, so
    # every command anchor moved down one pitch. What a rebinding batch must not
    # do is change what a button RUNS, and that is what is compared field for
    # field below. The anchors are compared as a SHAPE - the declared column, the
    # declared pitch, no gaps - which is a stronger statement than the row
    # numbers were, because it also refuses a collision and a stray button.
    was = by_key(accepted, ("buttons", "definitions"), "shape_name")
    now = by_key(current, ("buttons", "definitions"), "shape_name")
    for shape, definition in was.items():
        after = dict(now.get(shape) or {})
        assert after, f"{shape} was removed"
        for field in ("key", "sheet", "shape_name", "caption", "entry_point"):
            assert after[field] == definition[field], f"{shape} was rebound"
    assert set(now) - set(was) == {"btnPCCMResetResults"}, sorted(set(now) - set(was))
    assert {k: v for k, v in now["btnPCCMResetResults"].items()
            if k != "anchor_cell"} == {
        "key": "reset_results", "sheet": "Setup",
        "shape_name": "btnPCCMResetResults", "caption": "Reset Results",
        "entry_point": "PCCM_ResetResults",
    }
    block = current["commands"]["block"]
    column, first = str(block["button_column"]), int(block["first_button_row"])
    pitch = int(block["button_row_pitch"])
    commands = [b for b in current["buttons"]["definitions"]
                if b["sheet"] == block["sheet"]]
    assert [b["anchor_cell"] for b in commands] == [
        f"{column}{first + pitch * i}" for i in range(len(commands))]
    assert commands[0]["entry_point"] == "PCCM_ApplyTimeline"
    assert commands[-1]["entry_point"] == "PCCM_ResetResults"

    assert (set(current["vba"]["entry_points"]) - set(accepted["vba"]["entry_points"])
            == {"PCCM_ResetResults"})
    assert not set(accepted["vba"]["entry_points"]) - set(current["vba"]["entry_points"])

    was_modules = by_key(accepted, ("vba", "modules"), "name")
    now_modules = by_key(current, ("vba", "modules"), "name")
    for name, module in was_modules.items():
        assert now_modules.get(name) == module, f"the {name} declaration moved"
    assert set(now_modules) - set(was_modules) == {"modReset"}

    # AND THE PROTECTION POLICY IS UNTOUCHED BY ALL OF IT.
    assert current["protection"] == accepted["protection"]

    # THE COMMAND BLOCK KEEPS ITS GEOMETRY. Only the user-facing note changed,
    # and it changed to describe the button that was added to the block.
    block_now = dict(current["commands"]["block"])
    block_was = dict(accepted["commands"]["block"])
    assert block_now.pop("note") != block_was.pop("note")
    assert block_now == block_was, "the command block moved"
    # THE NOTE NAMES BOTH ADDITIONS: P10-2B's Reset Results, and P10-UX's
    # ordering - the timeline is applied first, which is why it leads the block.
    note = current["commands"]["block"]["note"]
    assert "Reset Results" in note
    assert "Apply / Update Timeline" in note and "first" in note


def test_13_protection_and_the_open_handler_are_untouched() -> None:
    """AND THE REST OF P10-2A WITH THEM."""
    for path in ("src/vba/modProtection.bas", "src/vba/ThisWorkbook.vba",
                 "builder/pccm_builder/protection.py"):
        current = (PCCM_ROOT / path).read_bytes()
        accepted = subprocess.run(["git", "show", f"{ACCEPTED}:pccm/{path}"],
                                  cwd=REPO_ROOT, check=True, stdout=subprocess.PIPE).stdout
        assert current == accepted, f"{path} moved; protection is accepted at {ACCEPTED}"


def test_14_repair_profiling_has_not_started() -> None:
    """THE SCOPE FENCE, checked rather than promised.

    CORRECTED AT P10-2B. Reset Results has landed under its own authorisation
    and is no longer behind this fence; Repair Profiling is, and the fence is
    kept rather than deleted so that the NEXT step cannot start early either.
    """
    assert not (SRC / "modRepair.bas").exists(), "modRepair belongs to a later step"
    structure = (SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    assert "PCCM_RepairProfiling" not in structure, "Repair belongs to a later step"


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(pytest.main([__file__, "-q"]))
