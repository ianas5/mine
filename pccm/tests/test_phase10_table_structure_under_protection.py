#!/usr/bin/env python3
"""P10-4A W3: table structure under protection - the audit, pinned.

WHAT WINDOWS RUN 3 SHOWED. The benchmark died building PERF-SMALL, on a
PowerShell `ListRow.Delete()`, with:

    Table features aren't available because the sheet is protected.

The immediate call was the harness's own. But the audit below is why that is not
the interesting half: SIX OF THE SEVEN USER COMMANDS make the same class of call
from VBA, and two of them make it UNCONDITIONALLY.

WHAT RUN 3 ALSO SHOWED, AND IT MATTERS. The fixture wrote four Setup scalars to
locked cells on protected sheets BEFORE it died. So `UserInterfaceOnly:=True` is
honoured for code that writes VALUES, and the thing Excel refused was
specifically a LISTOBJECT STRUCTURAL operation. The two capabilities are
different, and the workbook needs both.

WHAT THIS FILE DOES, AND DELIBERATELY DOES NOT DO.

  IT DOES NOT ASSERT THE DEFECT IS FIXED. It is not, and a control that claimed
  otherwise would be worse than none.

  IT PINS THE BLAST RADIUS. Every ListObject structural operation in production
  is named here with the command that reaches it and whether it fires
  unconditionally. A new one cannot appear undeclared, and the correction batch
  has an inventory to work from rather than a grep.

  IT PROVES THE LIFECYCLE IS NOT THE CAUSE. `Workbook_Open` calls the sole
  protection owner, which applies `UserInterfaceOnly:=True` to every sheet. That
  is classification C ruled out at the source.

  IT HOLDS THE PROBE HONEST. The probe that settles the remaining question -
  whether the VBA caller is refused exactly as the COM caller was - runs real
  endpoints, saves nothing, measures nothing and decides nothing.

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
BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
SPEC = PCCM_ROOT / "spec"
BUILD = PCCM_ROOT / "build"

import pytest  # noqa: E402
import yaml  # noqa: E402

PROBE = BOOTSTRAP / "phase10_protection_probe.ps1"
BENCHMARK = BOOTSTRAP / "phase10_benchmark.ps1"
ACCEPTED = "a0a4dc5"

# ---------------------------------------------------------------------------
# THE INVENTORY
# ---------------------------------------------------------------------------
# EVERY LISTOBJECT STRUCTURAL OPERATION IN PRODUCTION, the command that reaches
# it, and whether it fires unconditionally or only past a reserved capacity.
#
# "unconditional" is the load-bearing column. A capacity-dependent site is a
# defect a small model might not meet; an unconditional one is met by the first
# button a user presses on a fresh delivered workbook.
STRUCTURAL_SITES = (
    ("modProfiling", "target.ListColumns(target.ListColumns.Count).Delete",
     "PCCM_ApplyTimeline, PCCM_RepairProfiling", False,
     "shrinking the project-year band"),
    ("modProfiling", "Set added = target.ListColumns.Add",
     "PCCM_ApplyTimeline, PCCM_RepairProfiling", True,
     "a fresh workbook has NO year columns, so any timeline adds them"),
    ("modProfiling", "target.ListRows.Add",
     "SyncRows: every Add driver, PCCM_ApplyTimeline, PCCM_RepairProfiling", False,
     "past the reserved profiling rows"),
    ("modProfiling", "target.ListRows(r).Delete",
     "RemoveRow: PCCM_DeleteCostLine, PCCM_DeleteRisk", False,
     "when the grid holds more than one body row"),
    ("modInflation", "target.ListColumns(target.ListColumns.Count).Delete",
     "PCCM_ApplyTimeline", False, "shrinking the calendar band"),
    ("modInflation", "Set added = target.ListColumns.Add",
     "PCCM_ApplyTimeline", True,
     "a fresh workbook has NO inflation year columns"),
    ("modInflation", "target.ListRows.Add",
     "SyncProfileRows: PCCM_ApplyTimeline", False,
     "past the reserved inflation profile rows"),
    ("modDrivers", "register.ListRows.Add",
     "PCCM_AddCostLine, PCCM_AddRisk", False,
     "past the reserved register rows"),
    ("modDrivers", "register.ListRows(rowIndex).Delete",
     "PCCM_DeleteCostLine, PCCM_DeleteRisk", False,
     "when the register holds more than one body row"),
    ("modCalcReport", "target.ListRows(modWorkbook.BodyRowCount(target)).Delete",
     "PCCM_Calculate", True,
     "the _Calc tables are resized to the model on every Calculate"),
    ("modCalcReport", "target.ListRows.Add",
     "PCCM_Calculate", True,
     "the _Calc tables are resized to the model on every Calculate"),
    # AND THE ROLLBACK PATH ITSELF. `modWorkbook.RestoreTable` puts a snapshotted
    # table back to its recorded shape, and it is what PCCM_Calculate's
    # transactional rollback calls for all five _Calc tables. A structural
    # command that failed halfway under protection could therefore not be undone
    # either, which is a worse outcome than a clean refusal.
    ("modWorkbook", "Target.ListColumns(Target.ListColumns.Count).Delete",
     "RestoreTable: PCCM_Calculate rollback", False, "restoring a narrower table"),
    ("modWorkbook", "Target.ListColumns.Add",
     "RestoreTable: PCCM_Calculate rollback", False, "restoring a wider table"),
    ("modWorkbook", "Target.ListRows(BodyRowCount(Target)).Delete",
     "RestoreTable: PCCM_Calculate rollback", False, "restoring a shorter table"),
    ("modWorkbook", "Target.ListRows.Add",
     "RestoreTable: PCCM_Calculate rollback", False, "restoring a taller table"),
)

# The commands a user can press that reach at least one site above.
BLOCKED_COMMANDS = (
    "PCCM_ApplyTimeline", "PCCM_AddCostLine", "PCCM_AddRisk",
    "PCCM_DeleteCostLine", "PCCM_DeleteRisk", "PCCM_Calculate",
    "PCCM_RepairProfiling",
)

_MEMO: dict = {}


def _code(module: str) -> str:
    """Executable VBA only. A structural call must not be found in the comment
    that explains it, and a ban must not be satisfied by one."""
    key = ("code", module)
    if key not in _MEMO:
        _MEMO[key] = "\n".join(
            line for line in (SRC / f"{module}.bas").read_text(encoding="utf-8").splitlines()
            if not line.strip().startswith("'"))
    return _MEMO[key]


def _all_modules() -> list[str]:
    return sorted(path.stem for path in SRC.glob("*.bas"))


def _probe() -> str:
    if "probe" not in _MEMO:
        _MEMO["probe"] = PROBE.read_text(encoding="utf-8")
    return _MEMO["probe"]


def _probe_code() -> str:
    if "probe_code" not in _MEMO:
        text = re.sub(r"<#.*?#>", "", _probe(), flags=re.S)
        _MEMO["probe_code"] = "\n".join(
            line for line in text.splitlines() if not line.strip().startswith("#"))
    return _MEMO["probe_code"]


def _git(*args: str) -> str:
    if args not in _MEMO:
        _MEMO[args] = subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                                     stdout=subprocess.PIPE, text=True).stdout
    return _MEMO[args]


# ===========================================================================
# A. THE INVENTORY IS COMPLETE AND CURRENT
# ===========================================================================
def test_01_every_declared_structural_site_is_still_in_its_module() -> None:
    """The inventory describes the tree as it is, not as it was."""
    for module, statement, commands, _unconditional, _why in STRUCTURAL_SITES:
        assert statement in _code(module), f"{module}: {statement}"
        assert commands.strip()


def test_02_no_undeclared_structural_site_exists_anywhere_in_production() -> None:
    """NAMING OVER COUNTING. A structural operation that appeared in a module the
    inventory does not cover would be a seventh broken command nobody had
    audited, so the SET of modules is fixed and every occurrence inside them is
    accounted for."""
    pattern = re.compile(r"\.(ListRows|ListColumns)\b[^\n]*?\.(Add|Delete)\b"
                         r"|\.(ListRows|ListColumns)\.Add\b")
    found: dict[str, int] = {}
    for module in _all_modules():
        hits = [line.strip() for line in _code(module).splitlines()
                if pattern.search(line)]
        if hits:
            found[module] = len(hits)

    declared_modules = {module for module, *_rest in STRUCTURAL_SITES}
    assert set(found) == declared_modules, (
        f"the set of modules performing table structure moved: {sorted(found)} "
        f"vs the declared {sorted(declared_modules)}")
    for module, count in found.items():
        declared = len([1 for name, *_rest in STRUCTURAL_SITES if name == module])
        assert count == declared, (
            f"{module} performs {count} table-structure operations and the "
            f"inventory declares {declared}")


def test_03_two_commands_are_structural_unconditionally() -> None:
    """THIS IS WHY THE FINDING IS NOT ACADEMIC. Apply Timeline and Calculate are
    the first two buttons a user presses on a fresh delivered workbook, and
    neither can avoid the operation Excel refuses."""
    unconditional = {module for module, _s, _c, is_unconditional, _w in STRUCTURAL_SITES
                     if is_unconditional}
    assert unconditional == {"modProfiling", "modInflation", "modCalcReport"}, unconditional
    reached = {command for _m, _s, commands, is_unconditional, _w in STRUCTURAL_SITES
               if is_unconditional for command in re.findall(r"PCCM_\w+", commands)}
    assert "PCCM_ApplyTimeline" in reached
    assert "PCCM_Calculate" in reached


def test_04_the_affected_commands_are_declared_entry_points() -> None:
    """Every one of them is a real button, not an internal helper."""
    declared = set(yaml.safe_load(
        (SPEC / "structure_contract.yaml").read_text(encoding="utf-8"))["vba"]["entry_points"])
    for command in BLOCKED_COMMANDS:
        assert command in declared, command
    reached = set()
    for _m, _s, commands, _u, _w in STRUCTURAL_SITES:
        reached.update(re.findall(r"PCCM_\w+", commands))
    assert reached <= set(BLOCKED_COMMANDS), sorted(reached - set(BLOCKED_COMMANDS))


def test_05_the_rollback_path_is_structural_too() -> None:
    """THE FINDING WITH THE SHARPEST EDGE. `modWorkbook.RestoreTable` is what
    PCCM_Calculate's transactional rollback calls, and it resizes tables. A
    command that failed part-way through under protection could not be undone,
    which turns a clean refusal into a half-mutated workbook."""
    rollback = [module for module, _s, commands, _u, _w in STRUCTURAL_SITES
                if "RestoreTable" in commands]
    assert set(rollback) == {"modWorkbook"}, rollback
    calc = _code("modCalcReport")
    assert "modWorkbook.RestoreTable modWorkbook.Lo(CALC_SHEET, TBL_CALC_DRIVERS)" in calc
    assert "modWorkbook.SnapshotTable(modWorkbook.Lo(CALC_SHEET, TBL_CALC_DRIVERS))" in calc


# ===========================================================================
# B. THE LIFECYCLE IS NOT THE CAUSE
# ===========================================================================
def test_10_workbook_open_calls_the_sole_protection_owner() -> None:
    """CLASSIFICATION C, RULED OUT AT THE SOURCE. The open handler does exactly
    what the protection contract says it should."""
    event = "\n".join(
        line for line in (SRC / "ThisWorkbook.vba").read_text(encoding="utf-8").splitlines()
        if not line.strip().startswith("'"))
    assert "Private Sub Workbook_Open()" in event
    assert "modProtection.ProtectionApply(detail)" in event
    assert "Unprotect" not in event, "the open handler releases protection itself"


def test_11_protection_is_applied_with_user_interface_only() -> None:
    code = _code("modProtection")
    assert "sheet.Protect UserInterfaceOnly:=True," in code
    assert "If sheet.ProtectContents Then sheet.Unprotect" in code, (
        "the flag is assumed rather than re-established")
    assert "ThisWorkbook.Protect Structure:=True" in code


def test_12_there_is_exactly_one_protection_owner() -> None:
    """AND IT IS STILL THE ONLY ONE. Whatever the correction turns out to be, it
    may not put an Unprotect into a structural module."""
    owners = [module for module in _all_modules()
              if ".Protect " in _code(module) or ".Unprotect" in _code(module)]
    assert owners == ["modProtection"], f"protection is handled outside its owner: {owners}"
    event = (SRC / "ThisWorkbook.vba").read_text(encoding="utf-8")
    assert ".Unprotect" not in "\n".join(
        line for line in event.splitlines() if not line.strip().startswith("'"))


def test_13_no_structural_module_touches_protection_today() -> None:
    for module, *_rest in STRUCTURAL_SITES:
        code = _code(module)
        assert "Unprotect" not in code, f"{module} releases protection"
        assert ".Protect " not in code, f"{module} applies protection"


# ===========================================================================
# C. THE PROBE
# ===========================================================================
def test_20_the_probe_runs_the_real_entry_points() -> None:
    """IT PRESSES BUTTONS. A probe that reached around the commands would answer
    a question nobody asked."""
    code = _probe_code()
    # THE INVOCATION, NOT THE LABEL. `-Endpoint 'PCCM_Calculate'` also appears on
    # the Set-ProbeStage lines that bracket the _Calc shape reads, so the bare
    # substring is satisfied by a Calculate that is only being NAMED. A mutation
    # that swapped the real call for PCCM_RunSimulation walked straight through
    # this until it was anchored on the Invoke-ProbeEndpoint argument list.
    # TWICE FOR THE TWO THAT RUN TWICE. The delete path needs a second
    # ApplyTimeline (shrink the applied duration) and a second Calculate (let
    # ResizeBody delete the _Calc rows), so the growth round alone can no longer
    # satisfy this.
    # APPLYTIMELINE TWICE, AND CALCULATE THROUGH ITS ROUND HELPER TWICE. The
    # delete path needs a second ApplyTimeline (shrink the applied duration) and
    # a second Calculate (let ResizeBody delete the _Calc rows), so the growth
    # round alone can no longer satisfy this.
    assert code.count("-Endpoint 'PCCM_ApplyTimeline' -Resolution $resolution") == 2, (
        "ApplyTimeline is not invoked for both the growth and the shrink round")
    for endpoint in ("PCCM_AddCostLine", "PCCM_AddRisk"):
        assert code.count(f"-Endpoint '{endpoint}' -Resolution $resolution") == 1, endpoint
    # Calculate runs through ONE helper so the two rounds cannot drift apart,
    # and that helper is invoked exactly twice with the two round labels.
    assert code.count("-Endpoint 'PCCM_Calculate' -Resolution $Resolution") == 1, (
        "Calculate is invoked outside its round helper")
    assert code.count("Invoke-ProbeCalculateRound -Excel $excel") == 2, (
        "Calculate does not run for both the growth and the shrink round")
    assert "-Label 'growth'" in code and "-Label 'shrink'" in code
    # AND NO OTHER ENDPOINT IS INVOKED. Simulation, sensitivity, the annual step
    # and Reset are not structural and are not this probe's question.
    invoked = set(re.findall(r"-Endpoint '(\w+)' -Resolution \$[Rr]esolution", code))
    assert invoked == {"PCCM_ApplyTimeline", "PCCM_Calculate", "PCCM_AddCostLine",
                       "PCCM_AddRisk"}, sorted(invoked)
    assert "$Excel.Run($Endpoint)" in code
    assert "PCCM_AutomationResult" in code


def test_21_the_probe_separates_the_two_capabilities() -> None:
    """A VALUE WRITE AND A TABLE STRUCTURAL OPERATION ARE DIFFERENT THINGS, and
    Windows Run 3 showed one working and the other refused. The probe tests the
    control case deliberately rather than as a side effect."""
    code = _probe_code()
    assert "CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?" in code
    assert "function Invoke-ProbeLockedCellControl" in code
    assert "so UserInterfaceOnly IS honoured for code that writes VALUES to a" in code
    assert "It is a SEPARATE capability" in code or \
        "That is a SEPARATE capability from permission to perform" in code


def test_22_the_probe_checks_protection_is_actually_in_force() -> None:
    """OTHERWISE IT PROVES NOTHING. A workbook that came back unprotected would
    make every command below succeed for the wrong reason."""
    code = _probe_code()
    assert "function Get-ProbeProtectionState" in code
    # ON OPENING, AND AGAIN ON BOTH SIDES OF EVERY ENDPOINT. Probe Run 1 checked
    # only the opening and the very end; a command that released protection and
    # put it back would have looked identical to one that never touched it.
    assert "$opened = Get-ProbeProtectionState -Workbook $wb -Where 'as the workbook opened'" in code
    assert code.count("Get-ProbeProtectionState -Workbook $Workbook -Where") == 2
    assert "THE PROBE STOPS HERE, INCONCLUSIVE: the workbook did not come back" in code
    assert "ProtectionIsApplied" in code
    assert "$protectionInForce = ([bool](([int]$opened.Protected -eq [int]$opened.Total) -and" in code


def test_23_the_probe_measures_nothing_and_decides_nothing() -> None:
    """IT IS NOT A BENCHMARK AND NOT AN ACCEPTANCE RUN. Nothing it produces may
    be read as either."""
    code = _probe_code()
    assert "Stopwatch" not in code, "the probe is timing something"
    assert "warm" not in code.lower()
    assert "baseline" not in code.lower() or "NOT A BASELINE" in _probe()
    assert "THIS IS NOT A BASELINE, NOT A GATE-B RESULT AND NOT AN ACCEPTANCE RUN." in code
    assert "Invoke-Phase5GateBScenarios" not in code
    assert "Invoke-Phase6GateBScenarios" not in code


def test_24_the_probe_never_saves_and_shuts_down_cleanly() -> None:
    code = _probe_code()
    assert ".Save(" not in code and "SaveAs" not in code
    assert "$wb.Close($false)" in code
    for step in ("New-ReleaseLedger", "Invoke-NamedRelease $rel $excel 'Application'",
                 "$excel.Quit()", "Wait-ExcelExit -Identity $excelIdentity",
                 "Invoke-EmergencyExcelCleanup -Identity $excelIdentity",
                 "Format-ReleaseLedger $rel"):
        assert step in code, step
    assert "Stop-Process" not in code
    assert "} finally {" in code


def test_25_the_probe_reads_no_property_off_a_bare_command() -> None:
    """THE W1 LESSON, APPLIED FROM THE START rather than after another aborted
    run."""
    code = _probe_code()
    assert "Set-StrictMode -Version 2.0" in code
    assert "function Get-ProbeProperty" in code
    assert "-ErrorAction SilentlyContinue).Value" not in code
    assert "catch { }" not in code


def test_26_the_probe_reports_a_verdict_in_both_directions() -> None:
    """IT CAN EXONERATE PRODUCTION. A probe that could only confirm the
    hypothesis would not be evidence."""
    code = _probe_code()
    assert "PRODUCTION IS FINE UNDER PROTECTION" in code
    assert "PRODUCTION IS BLOCKED BY PROTECTION" in code
    assert "was a HARNESS defect only" in code


# ===========================================================================
# C2. PROBE RUN 1 - THE `.Count` DEFECT
# ===========================================================================
# WHAT HAPPENED. Probe Run 1 raised PropertyNotFoundStrict for 'Count' and died
# before it asked its question. Six statements in that draft read `.Count`; five
# were wrapped in `@()`, which guarantees an array before anything is read from
# it, and one was not:
#
#     for ($index = 1; $index -le $sheets.Count; $index++)
#
# `$sheets` was whatever `$Workbook.Worksheets` handed back, and under
# `Set-StrictMode -Version 2.0` a member that is not there is terminating.
#
# THE SHAPE IS GONE, NOT GUARDED. Collections are enumerated; no COM collection
# is indexed by position or asked for its `.Count` anywhere in the file.

# `.Count` read off something that is NOT the result of an `@(...)` - the shape
# that ended Probe Run 1. An ArrayList always has Count, but the probe does not
# rely on that either: the rule is uniform so a reader never has to judge.
_BARE_COUNT = re.compile(r"(?<!\))\s*\.Count\b")


def _bare_count_reads(code: str) -> list[str]:
    found: list[str] = []
    for match in re.finditer(r"\.Count\b", code):
        before = code[max(0, match.start() - 80):match.start()].rstrip()
        if before.endswith(")"):
            continue
        found.append(code[max(0, match.start() - 60):match.end()].splitlines()[-1].strip())
    return found


def test_40_no_count_is_read_off_an_unnormalised_operand() -> None:
    """PROBE RUN 1's DEFECT, AS A SHAPE RATHER THAN A LINE NUMBER."""
    offenders = _bare_count_reads(_probe_code())
    assert offenders == [], f"a .Count is read off an unnormalised operand: {offenders}"
    assert "$sheets.Count" not in _probe_code(), "the Probe Run 1 defect is back verbatim"
    # AND THE DETECTOR IS NOT VACUOUS: the exact statement that failed is fed to
    # it, and the normalised form it was replaced with is not flagged.
    assert _bare_count_reads("for ($index = 1; $index -le $sheets.Count; $index++) {"), (
        "the detector no longer detects the Probe Run 1 defect")
    assert _bare_count_reads("if (@($names).Count -gt 0) {") == []


def test_41_collections_are_enumerated_rather_than_indexed() -> None:
    """NO POSITION, NO LENGTH. Both are properties of a shape the probe cannot
    guarantee, and it needs neither."""
    code = _probe_code()
    assert "foreach ($sheet in @($sheets))" in code
    assert "function Measure-ProbeCollection" in code
    assert "foreach ($item in @($Collection))" in code
    assert ".Item($index)" not in code, "a COM collection is still indexed by position"
    assert "for ($index" not in code, "a positional loop over a COM collection remains"


def test_42_an_absent_collection_refuses_rather_than_counting_zero() -> None:
    """A FAKE ZERO WOULD BE A SHAPE CHANGE THE PROBE INVENTED, and shape change
    is exactly what its verdict turns on."""
    code = _probe_code()
    assert "throw ($Where + ': expected a ' + $Label + ' collection and got nothing')" in code
    assert "function Get-ProbeRequiredProperty" in code
    assert "It is not defaulted." in code
    assert "the workbook enumerated no worksheets" in code


def test_43_strict_mode_stays_on_and_nothing_is_swallowed() -> None:
    code = _probe_code()
    assert "Set-StrictMode -Version 2.0" in code
    assert code.count("Set-StrictMode") == 1
    assert "Set-StrictMode -Off" not in code
    assert "catch { }" not in code
    assert "-ErrorAction Ignore" not in code
    assert "-ErrorAction SilentlyContinue).Value" not in code


def test_44_the_failure_diagnostics_name_the_stage_and_the_endpoint() -> None:
    """PROBE RUN 1 REPORTED ONLY THE EXCEPTION. It could not say which of six
    `.Count` reads it meant."""
    code = _probe_code()
    assert "$script:ProbeCursor" in code
    assert "function Set-ProbeStage" in code
    assert "function Format-ProbeFailure" in code
    for field in ("stage", "doing", "endpoint", "detail", "exception", "message",
                  "at line", "statement", "command"):
        assert f"'  {field}" in code or f"  {field}" in code, field
    stages = set(re.findall(r"Set-ProbeStage -Stage '(\w+)'", code))
    assert stages == {"preflight", "setup", "protection", "resolve", "control",
                      "endpoint", "verdict"}, stages
    # THE ACTIONS TOO, NOT ONLY THE STAGE WORDS. Two boundaries now carry the
    # 'verdict' stage - weighing and reporting - so a set of stage names is
    # satisfied by either one alone, and dropping the other went unnoticed until
    # a mutation said so. Each boundary is named.
    actions = set(re.findall(r"Set-ProbeStage -Stage '\w+' `?\s*-Action '([^']+)'", code))
    for required in ("weighing the outcomes",
                     "reporting the structural-initialisation criteria"):
        assert required in actions, f"the '{required}' stage boundary is gone"


def test_45_diagnostics_never_touch_the_workbook_or_the_verdict() -> None:
    """THE CURSOR IS WRITTEN AND READ, and does nothing else. A diagnostic that
    changed the workbook would change the answer."""
    cursor = _probe_code().split("function Set-ProbeStage")[1]
    cursor = cursor[:cursor.index("\n}")]
    for banned in ("$Workbook", "$Excel", "$wb", "$verdict"):
        assert banned not in cursor, f"the cursor touches {banned}"
    failure = _probe_code().split("function Format-ProbeFailure")[1]
    failure = failure[:failure.index("\n}")]
    for banned in ("$Workbook", "$Excel", "$wb", "$verdict ="):
        assert banned not in failure, f"the failure formatter touches {banned}"


# ===========================================================================
# C3. VERDICT DISCIPLINE
# ===========================================================================
def test_50_the_verdict_starts_inconclusive_and_only_evidence_moves_it() -> None:
    """A PROBE-INTERNAL ERROR IS NEVER A STATEMENT ABOUT PRODUCTION."""
    code = _probe_code()
    assert "$verdict = 'INCONCLUSIVE'" in code
    assignments = re.findall(r"\$verdict = '([^']+)'", code)
    assert assignments == ["INCONCLUSIVE", "PRODUCTION IS BLOCKED BY PROTECTION",
                           "PRODUCTION IS FINE UNDER PROTECTION"], assignments
    # THE CATCH LEAVES IT WHERE IT WAS: it writes a REASON, never the verdict.
    assert "$verdictReason = ('the probe itself failed in stage " in code
    assert "the probe did not reach its conclusion" in code
    assert "A PROBE FAILURE IS NEVER A STATEMENT ABOUT PRODUCTION" in _probe()


def test_51_blocked_requires_protection_specific_evidence() -> None:
    """CORRECTED AFTER RUN 5, AND NARROWED WHILE IT MOVED.

    Run 5 produced the real answer - PCCM_ApplyTimeline invoked and refused with
    "Error 1004: Table features aren't available because the sheet is protected"
    - and then summarised it as "2 of 4 production endpoints did not succeed".
    The second one was PCCM_Calculate refusing because the applied timeline was
    pending: a business prerequisite that would refuse on a wholly unprotected
    workbook too. Counting it made the verdict look better-evidenced than it was.

    So BLOCKED is now gated on evidence attributable to the protected structural
    operation, and a bare refusal cannot reach it.
    """
    code = _probe_code()
    assert ("$protectionBlocked = @(@($outcomes) | Where-Object "
            "{ Test-ProbeProtectionBlocked -Outcome $_ })") in code
    # THE GATE ITSELF, anchored at the start of its line. `elseif (...)` contains
    # `if (...)` as a substring, and pinning the bare substring is how this
    # control went vacuous the moment the branch moved.
    gate = [line.strip() for line in code.splitlines()
            if "$protectionBlocked).Count -gt 0" in line]
    assert gate == ["if (@($protectionBlocked).Count -gt 0) {"], gate
    blocked_at = code.index("$verdict = 'PRODUCTION IS BLOCKED BY PROTECTION'")
    gate_at = code.index("if (@($protectionBlocked).Count -gt 0) {")
    assert gate_at < blocked_at, "BLOCKED is set outside the protection-specific gate"
    assert code.count("$verdict = 'PRODUCTION IS BLOCKED BY PROTECTION'") == 1


def test_51a_a_generic_refusal_cannot_reach_blocked() -> None:
    """REQUIRED CONTROL 23. A refusal that is not about protection is reported as
    a real result and is explicitly NOT counted."""
    code = _probe_code()
    assert "} elseif (@($notSucceeded).Count -gt 0) {" in code
    section = code[code.index("} elseif (@($notSucceeded).Count -gt 0) {"):]
    # FROM PAST ITS OWN OPENING LINE, or the slice ends where it starts and the
    # whole control becomes a check on an empty string.
    section = section[: section.index("} elseif", 1)]
    # THE VERDICT, NOT THE REASON. `$verdictReason` contains `$verdict`, and the
    # branch is SUPPOSED to write a reason - it is the verdict itself that must
    # stay where it started.
    assert "$verdict = " not in section, (
        "the not-a-protection-refusal branch sets a verdict; it must leave it INCONCLUSIVE")
    assert "$verdictReason = " in section, "the branch reports nothing at all"
    assert "none of them was" in section and "protection reason" in section
    # AND THE ONES THAT DID NOT COUNT ARE NAMED, not silently dropped.
    assert "$refusedForOtherReasons" in code
    assert "NOT counted here" in _probe()


def test_51b_the_protection_signature_is_excel_s_own_sentence() -> None:
    """REQUIRED CONTROL 22. Not the error number alone - 1004 is the most common
    Excel error there is - and not the word 'protect' alone, which the probe and
    the workbook both use in prose about protection."""
    body = _probe_code().split("function Test-ProbeProtectionBlocked")[1]
    body = body[: body.index("\n}\n")]
    assert "if (-not [bool]$Outcome.Invoked) { return $false }" in body, (
        "an endpoint that was never invoked could be called protection-blocked")
    assert "if ([string]$Outcome.Outcome -eq 'SUCCEEDED') { return $false }" in body
    assert "1004" in body
    assert "(?i)protect" in body
    assert "table features aren't available because the sheet is protected" in body.lower()
    # THE VALIDATION-REFUSAL FAMILIES ARE NOT IN THE SIGNATURE.
    for banned in ("REFUSED", "pending", "stale", "prerequisite"):
        assert banned not in body, f"the signature matches on {banned}"


def test_52_fine_requires_success_a_structural_effect_and_protection_throughout() -> None:
    """AN ANNOUNCEMENT IS NOT ENOUGH. A command that said OK and reshaped nothing
    has not shown that the structural operation is permitted."""
    code = _probe_code()
    assert "$structural = @(@($outcomes) | Where-Object { [bool]$_.StructuralEffect })" in code
    assert "elseif (@($structural).Count -lt 1) {" in code
    assert "no watched table ever " in code
    assert "$lostProtection = @(@($outcomes) | Where-Object {" in code
    assert "elseif (@($lostProtection).Count -gt 0) {" in code
    # FINE is the LAST branch: it is reached only when every other reason to
    # doubt has been ruled out.
    assert code.index("$verdict = 'PRODUCTION IS FINE UNDER PROTECTION'") >         code.index("elseif (@($lostProtection).Count -gt 0) {")


def test_53_every_endpoint_records_protection_and_shape_on_both_sides() -> None:
    """REQUIRED EVIDENCE. Announcement, outcome class, protection before and
    after, and what actually changed."""
    code = _probe_code()
    for field in ("Endpoint", "Outcome", "Result", "Raised", "Changes",
                  "ProtectedBefore", "ProtectedAfter", "TotalSheets",
                  "ProtectionBefore", "ProtectionAfter", "StructuralEffect"):
        assert f"{field}" in code, field
    assert "$protectionBefore = Get-ProbeProtectionState -Workbook $Workbook -Where ('before ' + $Endpoint)" in code
    assert "$protectionAfter = Get-ProbeProtectionState -Workbook $Workbook -Where ('after ' + $Endpoint)" in code
    assert "$shapesBefore = Get-ProbeAllShapes" in code
    assert "$shapesAfter = Get-ProbeAllShapes" in code


def test_54_the_three_failure_kinds_are_distinguished() -> None:
    """A PRODUCTION REFUSAL, AN EXCEL RUNTIME FAILURE AND A PROBE FAILURE are
    three different facts."""
    code = _probe_code()
    assert "$outcome = 'REFUSED'" in code
    assert "if (-not $invoked)    { $outcome = 'NOT INVOKED' }" in code
    assert "elseif (-not $announced) { $outcome = 'RAISED' }" in code
    assert "elseif ($succeeded)   { $outcome = 'SUCCEEDED' }" in code
    # AND AN ENDPOINT THAT WAS NEVER ENTERED IS NOT A PRODUCTION RESULT AT ALL.
    assert "$invoked = $false" in code
    assert "$invoked = $true" in code
    assert code.count("$invoked = $true") == 1, (
        "the invoked flag is set in more than one place")


def test_55_the_watched_tables_come_from_the_manifest() -> None:
    """NOT A LIST TYPED IN THE PROBE. A grid the manifest gained and the probe
    did not would be a structural effect nobody looked for."""
    code = _probe_code()
    assert "foreach ($register in @($Manifest.registers))" in code
    assert "foreach ($grid in @($Manifest.grids))" in code
    for typed in ("tblCostLines", "tblRiskRegister", "tblCostProfiling", "tblInflation",
                  "'Cost Lines'", "'Risk Register'", "'Cost Profiling'"):
        assert typed not in code, f"the probe types the identifier {typed}"
    # AND THE TWO IDENTIFIERS ARE INDEPENDENT. Neither is derived from the other.
    assert "-InputObject $entry -Name 'Sheet' -Where $where" in code
    assert "-InputObject $entry -Name 'Table' -Where $where" in code


def test_55b_the_probe_never_reaches_around_the_commands() -> None:
    """NO BYPASS, AND NOT EVEN THE WORDS. A probe that released protection, or
    performed the structural operation itself, would answer a question nobody
    asked - and would answer it FINE every time."""
    code = _probe_code()
    # THE CALL FORMS, NOT THE LETTERS. "Unprotected" is a state the probe REPORTS
    # and contains the word it must not CALL, so the ban is on the invocation.
    # THE CALL FORM CARRIES A LEADING DOT. The probe NAMES these operations in
    # the sentence that explains what each command does - "ListColumns.Add on
    # three grids" - and naming one is the opposite of performing it. What is
    # banned is `<object>.ListColumns.Add`, which is a call.
    for banned in (".Unprotect(", ".Unprotect ", "'ProtectionRelease'",
                   '"ProtectionRelease"', "'ProtectionApply'", ".Protect(",
                   ".ListRows.Add", ".ListColumns.Add", "$victim.Delete()",
                   ".Rows(1).Delete", ".EntireRow.Delete"):
        assert banned not in code, f"the probe performs or releases: {banned}"
    # AND IT READS THE TWO COLLECTIONS IT MEASURES, which is not a mutation.
    assert "$lo.ListColumns" in code and "$lo.ListRows" in code
    # THE ONE `.Delete()` IN THE FILE IS THE SENTENCE QUOTING RUN 3's FAILURE.
    # A probe that quoted the defect and also performed it would be found here.
    # NAMING THE DELETE IS NOT PERFORMING ONE. The probe quotes the exact call
    # Benchmark Run 3 died on, and the shrink round's expectations say which
    # delete each endpoint must be seen doing - all of it prose. What is banned
    # is a delete this script CALLS, so the rule is stated over the syntax: a
    # `.Delete(` may only appear inside a string.
    for line in code.splitlines():
        if ".Delete(" not in line:
            continue
        stripped = line.strip()
        assert (stripped.startswith("Write-ProbeLine") or
                ("'" in stripped.split(".Delete(")[0]) or
                ('"' in stripped.split(".Delete(")[0])), (
            f"the probe performs a delete rather than naming one: {stripped}")
    # EVERY WRITE GOES THROUGH ONE OF TWO AUDITED HELPERS, and the raw COM
    # assignment sites exist ONLY inside them. Run 4 replaced the probe's own
    # reimplemented setter with the accepted Set-NamedValue, so this states the
    # rule where it now lives rather than counting one call form.
    assert code.count("Set-NamedValue -Workbook $Workbook") == 1, (
        "the named-value write happens somewhere other than Set-ProbeDeclaredInputs")
    assert code.count("Set-ProbeCellExact -Cell $cell") == 2, (
        "the control writes to its cell more than twice")
    # THE THREE TIMELINE INPUTS ARE DECLARED, ONCE, AND DRIVE THAT ONE CALL.
    # FOUR NOW, AND THE FOURTH IS DECLARED. Windows Run 6 invoked PCCM_Calculate
    # and it refused for a NON-protection reason - "Discount Rate: the value is
    # blank. A blank is not zero." - which is production validating correctly.
    # The discount rate is an ordinary Setup input that modCalcResolve requires,
    # so the probe supplies it the way a user does. The set is still exactly the
    # minimum: nothing here is needed by a command that does not read it.
    # FIVE DECLARATIONS OVER FOUR INPUTS: duration_years appears twice because
    # the shrink round re-applies it at a smaller value, which is the whole
    # mechanism that drives ListColumns.Delete and ListRows.Delete.
    declared = re.findall(r"@\{ Key = '(\w+)';\s+Value = \[double\]", code)
    assert declared == ["base_year", "project_start_year", "duration_years",
                        "discount_rate", "duration_years"], declared
    assert declared[-1] == "duration_years", "the shrink round changes something else"
    # AND NO RAW Value2 ASSIGNMENT SURVIVES OUTSIDE THE TWO HELPERS. This is the
    # control that would catch a third write added straight to the COM object.
    setter = code.split("function Set-NamedValue")[1]
    setter = setter[: setter.index("\n}\n")]
    exact = code.split("function Set-ProbeCellExact")[1]
    exact = exact[: exact.index("\n}\n")]
    inside = setter.count(".Value2 = ") + exact.count(".Value2 = ")
    assert code.count(".Value2 = ") == inside == 4, (
        f"a Value2 assignment lives outside the two helpers: {code.count('.Value2 = ')} vs {inside}")


def test_56_an_inconclusive_run_exits_non_zero() -> None:
    code = _probe_code()
    assert "if ($verdict -eq 'INCONCLUSIVE') { exit 2 }" in code
    assert "exit 0" in code


# THE REGION THAT DECIDES THE VERDICT, as opposed to the one that reports the
# structural-initialisation criteria afterwards. The two are deliberately
# separate: criterion B names the locked-cell control's result because reporting
# it is exactly what that block is for, while the deciding branch must never see
# it. A slice that ran to the end of the run would conflate them.
_CRITERIA_STAGE = "-Action 'reporting the structural-initialisation criteria'"


def _verdict_decision_block(code: str) -> str:
    start = code.index("Set-ProbeStage -Stage 'verdict'")
    end = code.index(_CRITERIA_STAGE)
    assert start < end, "the criteria are reported before the verdict is weighed"
    return code[start:end]


def test_57_the_locked_cell_control_is_kept_and_kept_separate() -> None:
    """ITS PURPOSE IS TO PROVE THE OTHER CAPABILITY, and to be unable to be
    mistaken for this one."""
    code = _probe_code()
    assert "CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?" in code
    assert "so UserInterfaceOnly IS honoured for code that writes VALUES to a" in code
    assert "a ListObject structural operation, and it settles nothing about one." in code
    assert "$controlResult = 'NOT ATTEMPTED'" in code
    # THE CONTROL NEVER MOVES THE VERDICT. Bounded to the branch that decides
    # it, because the control's own summary is printed further down and finding
    # the variable there would prove nothing.
    # THE DECIDING REGION ONLY. There are now two 'verdict' stages: the one that
    # weighs the outcomes, and the one that REPORTS the structural-initialisation
    # criteria. Criterion B legitimately names the control's result - reporting
    # it is the point - so the slice stops where deciding stops.
    verdict_block = _verdict_decision_block(code)
    assert "$controlWorked" not in verdict_block, "the control decides the question"
    assert "$controlDetail" not in verdict_block


# ===========================================================================
# C4. PROBE RUN 2 - THE LOOKUP AND THE CONTRADICTORY CONTROL
# ===========================================================================
# TWO DEFECTS, ONE RUN.
#
#   DISP_E_BADINDEX at `$ws = $sheets.Item($SheetName)`. Not a missing sheet:
#   all five watched tab names and all five table names exist in the built
#   workbook, and `test_60` proves it against the artifacts. Every shape read
#   re-acquired `$Workbook.Worksheets` and released it again - dozens of times
#   per run against one underlying collection - and the lookup eventually failed
#   on a collection that had been released out from under it.
#
#   THE CONTROL REPORTED SUCCESS AND FAILURE AT ONCE. The write and the restore
#   shared a try block; the write succeeded and printed, the restore threw, and
#   the catch printed a refusal over the top. The target was also `inpDiscountRate`
#   - an EDITABLE INPUT, Locked=False - so it was never a locked-cell control at
#   all.

_WORKBOOK = BUILD / "PCCM_stageA.xlsx"
_MANIFEST_FILE = BUILD / "stage_b_manifest.json"


@pytest.mark.skipif(not _WORKBOOK.is_file() or not _MANIFEST_FILE.is_file(),
                    reason="Stage A has not been built into pccm/build")
def test_60_every_watched_identifier_resolves_against_the_built_workbook() -> None:
    """THE CONTROL THAT SETTLES "IS THE SHEET MISSING?" WITHOUT A WINDOWS RUN.

    Both identifiers come from the manifest, independently, and both are looked
    up in the workbook Stage A actually built. A tab name that drifted, a table
    renamed, or a register moved to another sheet fails here rather than as a
    DISP_E_BADINDEX an hour into a Windows run.
    """
    import json

    import openpyxl

    manifest = json.loads(_MANIFEST_FILE.read_text(encoding="utf-8"))
    workbook = openpyxl.load_workbook(_WORKBOOK)
    watched = ([(entry["key"], entry["sheet"], entry["table_name"])
                for entry in manifest["registers"]] +
               [(entry["key"], entry["sheet"], entry["table_name"])
                for entry in manifest["grids"]])
    assert len(watched) == 5, watched

    missing: list[str] = []
    for key, sheet, table in watched:
        if sheet not in workbook.sheetnames:
            missing.append(f"{key}: no worksheet named {sheet!r}")
            continue
        tables = getattr(workbook[sheet], "tables", {})
        if table not in tables:
            missing.append(f"{key}: {sheet!r} carries no table named {table!r}")
    assert missing == [], missing


def test_61_the_worksheets_collection_is_resolved_once_and_held() -> None:
    """THE CHURN THAT PRODUCED DISP_E_BADINDEX IS GONE. One acquire, one
    release, and every shape read works from the already-resolved objects."""
    code = _probe_code()
    assert "function Resolve-ProbeTargets" in code
    assert "function Release-ProbeTargets" in code
    assert code.count("$Workbook.Worksheets") == 2, (
        "the probe acquires the Worksheets collection more than twice: once to "
        "resolve the targets and once to read protection")
    assert code.count(".Item($sheetName)") == 1, (
        "a worksheet is still being looked up more than once")
    assert "$lo = $Target.ListObject" in code, (
        "the shape read still resolves its own ListObject")
    assert "Release-ProbeTargets -Resolution $resolution" in code


def test_62_an_unresolvable_sheet_or_table_is_a_probe_failure() -> None:
    """REQUIRED CONTROL 4 OF THIS ROUND. INCONCLUSIVE, never BLOCKED - and never
    a fallback to index 1."""
    code = _probe_code()
    assert "throw ('the workbook has no worksheet named '" in code
    assert "throw ('worksheet ' + [char]39 + $sheetName + [char]39 + ' carries no table named '" in code
    assert ".Item(1)" not in code, "the probe falls back to a positional index"
    # THE THROW LANDS IN THE HANDLER THAT LEAVES THE VERDICT ALONE.
    assert "A PROBE FAILURE IS NEVER A STATEMENT ABOUT PRODUCTION" in _probe()


def test_63_the_codename_is_recorded_and_never_used_to_find_anything() -> None:
    """REQUIRED CONTROL 3. Tab name and CodeName are different identifiers and
    the probe does not let one stand in for the other."""
    code = _probe_code()
    assert "CodeName  = [string]$ws.CodeName" in code
    assert ".Item($ws.CodeName)" not in code
    assert "$sheets.Item($codeName)" not in code
    # IT IS EVIDENCE, so it reaches the log in the target's own block.
    assert "'    codename = '" in code


def test_64_the_endpoint_is_not_called_invoked_until_application_run() -> None:
    """REQUIRED CONTROL 5. Probe Run 2 said "invoking the production entry
    point" while it was still collecting pre-command evidence, and the endpoint
    was never reached."""
    code = _probe_code()
    body = code.split("function Invoke-ProbeEndpoint")[1]
    body = body[:body.index("\n}")]
    assert "PREPARING TO TEST" in body
    assert "ENDPOINT INVOKED: Application.Run has been entered" in body
    # THE FLAG IS SET AFTER THE PREPARATION AND BEFORE THE CALL.
    assert body.index("$shapesBefore = Get-ProbeAllShapes") < body.index("$invoked = $true")
    assert body.index("$invoked = $true") < body.index("$Excel.Run($Endpoint)")
    assert "Invoked           = $invoked" in body


def test_65_the_evidence_order_is_the_declared_one() -> None:
    """REQUIRED ORDER: resolve, protection-before, shape-before, mark, invoke,
    announcement, shape-after, protection-after, classify."""
    body = _probe_code().split("function Invoke-ProbeEndpoint")[1]
    body = body[:body.index("\n}")]
    order = ["$protectionBefore = Get-ProbeProtectionState",
             "$shapesBefore = Get-ProbeAllShapes",
             "$invoked = $true",
             "$Excel.Run($Endpoint)",
             "$result = [string]$Excel.Run('PCCM_AutomationResult')",
             "$shapesAfter = Get-ProbeAllShapes",
             "$protectionAfter = Get-ProbeProtectionState"]
    positions = [body.index(step) for step in order]
    assert positions == sorted(positions), list(zip(order, positions))


def test_66_the_control_cannot_report_success_and_refusal_at_once() -> None:
    """REQUIRED CONTROL 6. Probe Run 2 printed SUCCEEDED, then REFUSED, then
    True - three contradictory lines out of one try block."""
    code = _probe_code()
    body = code.split("function Invoke-ProbeLockedCellControl")[1]
    body = body[:body.index("\n}\n")]
    # ONE RESULT PER RUN: every path returns, and every return names one word.
    results = re.findall(r"Result = '(\w+)'", body)
    assert set(results) == {"INCONCLUSIVE", "REFUSED", "SUCCEEDED"}, results
    assert body.count("return [pscustomobject]@{") == len(results)
    # THE WRITE AND THE RESTORE ARE IN SEPARATE TRY BLOCKS, so a cleanup failure
    # can never be printed as a refusal of the write.
    assert "try { Set-ProbeCellExact -Cell $cell -Value $probeText } catch { $writeRaised = (Format-Err $_) }" in body
    assert "try { Set-ProbeCellExact -Cell $cell -Value $original } catch { $restoreRaised = (Format-Err $_) }" in body
    # AND THE CALLER TAKES ONE ANSWER, as a WORD rather than a boolean.
    assert "$controlResult = [string]$control.Result" in code
    assert "$controlWorked" not in code, (
        "a boolean projection of the control is back; NOT ATTEMPTED has no False")


def test_67_the_control_proves_its_target_is_locked_before_it_writes() -> None:
    """REQUIRED CONTROL 7. Probe Run 2 wrote to an EDITABLE INPUT, which proves
    nothing about UserInterfaceOnly: a user can type in that cell."""
    body = _probe_code().split("function Invoke-ProbeLockedCellControl")[1]
    body = body[:body.index("\n}\n")]
    assert "if (-not $target.Worksheet.ProtectContents) {" in body
    assert "if (-not $cell.Locked) {" in body
    assert body.index("if (-not $cell.Locked) {") < body.index("-Value $probeText")
    # THE BAN IS ON THE CONTROL, NOT ON THE FILE. Probe Run 2 wrote its
    # locked-cell control to inpDiscountRate, which is Locked=False by design and
    # proves nothing about UserInterfaceOnly. The discount rate is now a DECLARED
    # Calculate prerequisite written elsewhere - as an ordinary editable Setup
    # input, which is exactly what it is - so the ban is stated where it belongs.
    assert "inpDiscountRate" not in _probe_code(), (
        "a defined name is hard-coded; it belongs in the inspection")
    assert "discount_rate" not in body, (
        "the locked-cell control still writes to an editable input")
    for banned in ("Set-NamedValue", "Set-ProbeDeclaredInputs"):
        assert banned not in body, f"the control writes through {banned}"


def test_68_the_original_value_is_restored_and_verified() -> None:
    """REQUIRED CONTROL 8, and a blank original is refused rather than
    'restored' - Probe Run 2's original was blank and that is what threw."""
    body = _probe_code().split("function Invoke-ProbeLockedCellControl")[1]
    body = body[:body.index("\n}\n")]
    assert "$original = $cell.Value2" in body
    assert "if ([string]::IsNullOrWhiteSpace([string]$original)) {" in body
    # THE READ IS NO LONGER STRINGIFIED, AND NEITHER IS THE COMPARISON. A
    # [string] comparison accepts a Double 2026 restored as the text '2026',
    # which is a different cell content wearing the same characters. The rule is
    # the accepted Test-Phase5ExactValue one: CLR type identity first.
    assert "$restored = $cell.Value2" in body
    assert "[string]$cell.Value2" not in body, "the control stringifies a captured value again"
    assert "if (-not (Test-ProbeExactValue -Actual $restored -Expected $original)) {" in body
    assert "the original value was restored and verified" in body
    comparator = _probe_code().split("function Test-ProbeExactValue")[1]
    comparator = comparator[: comparator.index("\n}\n")]
    assert "$Actual.GetType().FullName -cne $Expected.GetType().FullName" in comparator, (
        "the comparator does not establish exact CLR type identity first")
    assert "[string]$Actual -ceq [string]$Expected" in comparator


def test_69_a_cleanup_failure_cannot_produce_a_successful_control() -> None:
    """REQUIRED CONTROL 9. The capability answer and the cleanup answer are
    separate facts, and a failed cleanup makes the control INCONCLUSIVE - never
    'protection blocks value writes'."""
    body = _probe_code().split("function Invoke-ProbeLockedCellControl")[1]
    body = body[:body.index("\n}\n")]
    cleanup = body[body.index("$restoreRaised = ''"):]
    successes = re.findall(r"Result = '(\w+)'", cleanup)
    assert "SUCCEEDED" in successes
    assert successes.count("REFUSED") == 0, (
        "a cleanup failure can be reported as a refusal of the write")
    for detail in ("the original value could not be restored",
                   "the restored value did not match the original",
                   "protection was lost during the control"):
        assert detail in cleanup, detail


def test_70_an_untrustworthy_control_stops_the_probe() -> None:
    """A PROBE THAT CANNOT RESTORE WHAT IT CHANGED HAS ESTABLISHED NOTHING, so
    the production question is not asked over it."""
    code = _probe_code()
    assert "if ($controlResult -eq 'INCONCLUSIVE') {" in code
    assert "THE PROBE STOPS HERE, INCONCLUSIVE: the control did not settle." in code
    assert "The production question was not asked." in code


# ===========================================================================
# C5. PROBE RUN 3 - FIVE RECORDS THAT BECAME ONE
# ===========================================================================
# WHAT HAPPENED. `Get-ProbeWatchedTables` ended with `return ,@($watched)`. The
# unary comma makes the function emit ONE pipeline item that IS the array - and
# the caller wrote `@(Get-ProbeWatchedTables ...)`, which COLLECTS pipeline
# items rather than flattening nested arrays, so the five records arrived
# double-wrapped. `foreach ($entry in @($Watched))` then bound `$entry` to the
# inner array, `$entry.Sheet` became MEMBER ENUMERATION over five records, and
# `[string]` joined the result with spaces. The probe asked Excel for a
# worksheet named 'Cost Lines Risk Register Cost Profiling Risk Profiling
# Inflation'.
#
# THE COERCION IS WHAT HID IT. `[string]` turned a structural error into a
# plausible-looking name, and the only symptom was DISP_E_BADINDEX.
#
# AND IT EXPLAINS RUN 2 TOO. That run died on the same lookup with the same
# HRESULT; the earlier diagnosis blamed COM release churn. The joined name was
# already the cause - Run 2's message simply did not carry it.
#
# THE SECOND INSTANCE WAS WORSE THAN AN ABORT. `Get-ProbeShapeDelta` used the
# same idiom, so an EMPTY change list arrived as a one-element array and
# `StructuralEffect` would have been true for every endpoint. The probe could
# have reached PRODUCTION IS FINE having observed no shape change at all.

_MULTI_RECORD_PRODUCERS = ("Get-ProbeWatchedTables", "Get-ProbeShapeDelta")


def test_80_no_multi_record_helper_wraps_its_result_in_a_unary_comma() -> None:
    """THE EXACT IDIOM THAT COLLAPSED THE RECORDS, banned where it collapses.

    `return ,@($x)` is correct only when the caller does NOT wrap in `@()`.
    Every caller here does, so the producers emit their records normally and the
    caller's `@()` keeps a zero- or one-record result an array.
    """
    code = _probe_code()
    assert "return ,@(" not in code, "a helper still returns a comma-wrapped array"
    for producer in _MULTI_RECORD_PRODUCERS:
        body = code.split(f"function {producer} {{")[1]
        body = body[:body.index("\n}")]
        returns = re.findall(r"return (.+)", body)
        assert returns and all(not value.strip().startswith(",") for value in returns), (
            producer, returns)


def test_81_the_watched_shape_is_proved_before_excel_is_started() -> None:
    """REQUIRED CONTROLS 1, 6 AND 12. Five records, each with scalar fields,
    unique keys and unique pairs - checked at the boundary, not discovered by a
    COM lookup an hour later."""
    code = _probe_code()
    assert "function Assert-ProbeWatchedShape" in code
    assert "$null = Assert-ProbeWatchedShape -Watched $watched -Manifest $manifest" in code
    assert "REFUSED, BEFORE EXCEL WAS STARTED." in code
    # THE COUNT COMES FROM THE MANIFEST, never a literal five in this file.
    assert "$expected = (@($Manifest.registers).Count + @($Manifest.grids).Count)" in code
    assert "-ne 5" not in code and "-eq 5" not in code, "the record count is hard-coded"
    # UNIQUENESS, BOTH WAYS, AND EACH ASSERTED SEPARATELY - one check covering
    # for the other is how a dropped guard passes.
    assert "if ($keys -contains $key) { throw ($where + \": the key '\" + $key + " in code
    assert "if ($pairs -contains $pair) { throw ($where + ': ' + $pair + ' is not unique') }" in code
    assert "$pair = $sheet + '!' + $table" in code
    # AND THE RESOLVED SET IS RE-CHECKED, so a resolution that merged a target
    # cannot be measured happily.
    assert "Assert-ProbeWatchedShape -Watched $resolution.Targets" in code


def test_82_scalar_shape_is_proved_before_any_string_coercion() -> None:
    """REQUIRED CONTROL 3 OF THIS ROUND. `[string]` must never be the thing that
    discovers a property holds five values."""
    code = _probe_code()
    body = code.split("function Get-ProbeScalarString {")[1]
    body = body[:body.index("\n}")]
    # THE ARRAY TEST COMES BEFORE THE CAST.
    assert body.index("$value -is [System.Array]") < body.index("return [string]$value")
    assert "holds a collection of " in body
    assert "have been collapsed into an aggregate" in body
    # AND THE RESOLVER USES IT INSTEAD OF A BARE CAST.
    resolver = code.split("function Resolve-ProbeTargets {")[1]
    resolver = resolver[:resolver.index("\n}")]
    assert "[string]$entry.Sheet" not in resolver
    assert "[string]$entry.Table" not in resolver
    assert "[string]$entry.Key" not in resolver
    assert resolver.count("Get-ProbeScalarString") >= 3
    # AND EACH NAME IS USED FOR ITS OWN LOOKUP. Swapping them would resolve a
    # table name as a worksheet and produce exactly the error Run 3 reported.
    assert ("$sheetName = Get-ProbeScalarString -InputObject $entry -Name 'Sheet' "
            "-Where $where") in resolver
    assert ("$tableName = Get-ProbeScalarString -InputObject $entry -Name 'Table' "
            "-Where $where") in resolver
    assert "$ws = $sheets.Item($sheetName)" in resolver
    assert "$lo = $los.Item($tableName)" in resolver


def test_83_a_record_that_is_itself_an_array_is_refused() -> None:
    """THE COLLAPSE, NAMED. `$entry` bound to the inner array is exactly what
    happened, and it now refuses instead of member-enumerating."""
    body = _probe_code().split("function Assert-ProbeWatchedShape {")[1]
    body = body[:body.index("\n}")]
    assert "if ($record -is [System.Array]) {" in body
    assert "The collection boundary collapsed." in body


def test_84_every_watched_record_is_printed_in_its_own_block() -> None:
    """REQUIRED CONTROL 4 OF THIS ROUND. Five records printed as one line is how
    the collapse went unnoticed until Excel refused the name."""
    code = _probe_code()
    assert "Write-ProbeLine ('    tab   = ' + [string]$entry.Sheet)" in code
    assert "Write-ProbeLine ('    table = ' + [string]$entry.Table)" in code
    assert "Write-ProbeLine ('    tab      = ' + [string]$target.Sheet)" in code
    assert "Write-ProbeLine ('    codename = ' + [string]$target.CodeName)" in code
    assert "Write-ProbeLine ('    table    = ' + [string]$target.Table)" in code
    # AND THE COUNT IS SHOWN, so a collapse is visible in the log itself.
    assert "' targets resolved'" in code
    assert "' records)')" in code


def test_85_the_control_has_four_explicit_states() -> None:
    """REQUIRED CONTROLS 14 AND 15. Probe Run 3 printed "UserInterfaceOnly
    permits code VALUE writes: False" after failing before the control ever ran.
    False reads as "blocking was observed", which nothing had tested."""
    code = _probe_code()
    assert "$controlResult = 'NOT ATTEMPTED'" in code
    assert "UserInterfaceOnly code-value-write capability: NOT TESTED" in code
    assert "UserInterfaceOnly code-value-write capability: CONFIRMED" in code
    assert "UserInterfaceOnly code-value-write capability: REFUSED BY EXCEL" in code
    # NO BOOLEAN PROJECTION AT ALL, so there is no False to be misread.
    assert "permits code VALUE writes: ' + [string]" not in code
    assert "$controlWorked" not in code
    body = code.split("function Invoke-ProbeLockedCellControl")[1]
    body = body[:body.index("\n}\n")]
    states = set(re.findall(r"Result = '([A-Z ]+)'", body)) | {"NOT ATTEMPTED"}
    assert states == {"SUCCEEDED", "REFUSED", "INCONCLUSIVE", "NOT ATTEMPTED"}, states


def test_86_the_control_state_does_not_reach_the_verdict_branch() -> None:
    """REQUIRED: the reporting correction must not alter the verdict logic."""
    code = _probe_code()
    verdict = _verdict_decision_block(code)
    assert "$controlResult" not in verdict
    assert "$controlDetail" not in verdict


# ===========================================================================
# D. NOTHING WAS CHANGED WHILE THE QUESTION IS OPEN
# ===========================================================================
def test_30_no_production_source_changed() -> None:
    """THE CLASSIFICATION IS ESTABLISHED BEFORE THE CODE MOVES. Editing four
    accepted modules on a reading of the documentation is exactly what the
    Windows evidence exists to prevent."""
    changed = [line for line in _git("diff", "--name-only", ACCEPTED, "--",
                                     "pccm/src", "pccm/spec").splitlines() if line.strip()]
    # P10-RP. PRODUCTION HAS NOW MOVED, UNDER ITS OWN AUTHORISATION, and this
    # control says so rather than being deleted. The probe's classification was
    # established BEFORE the code moved - that is the claim, and it still holds:
    # PCCM_ApplyTimeline was invoked on Windows and refused with Error 1004
    # before one line of production changed. What changed afterwards is the
    # declared structural window, and the reversal proves it is the only thing.
    import sys as _sys
    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_structural_window import (DECLARED_STRUCTURAL_WINDOW_CHANGES,
                                       strip_structural_window)
    declared = {f"pccm/src/vba/{name}" for name in DECLARED_STRUCTURAL_WINDOW_CHANGES}
    for path in sorted(set(changed) & declared):
        name = Path(path).name
        current = strip_structural_window(
            name, (PCCM_ROOT / "src" / "vba" / name).read_bytes().decode("utf-8"))
        accepted = _git("show", f"{ACCEPTED}:{path}")
        assert current.replace("\r\n", "\n") == accepted.replace("\r\n", "\n"), (
            f"{path} moved outside the declared P10-RP structural window")
    changed = [path for path in changed if path not in declared]
    assert changed == [], f"production source changed: {changed}"


def test_31_the_benchmark_still_builds_its_scenario_through_production() -> None:
    """NO BYPASS. A fixture that wrote around a broken user command would hide a
    release defect behind a performance number."""
    benchmark = "\n".join(
        line for line in BENCHMARK.read_text(encoding="utf-8").splitlines()
        if not line.strip().startswith("#"))
    assert "Set-Phase5Fixture -Excel $excel -Workbook $wb" in benchmark
    assert "$Excel.Run([string]$Operation.endpoint)" in benchmark
    assert "Unprotect" not in benchmark, "the benchmark releases protection to get numbers"
    assert "ProtectionRelease" not in benchmark


def test_32_an_aborted_fixture_cannot_establish_a_baseline() -> None:
    """WINDOWS RUN 3 PRODUCED 0 OF 11 WARM MEDIANS, and the machinery already
    refuses to call that anything."""
    benchmark = "\n".join(
        line for line in BENCHMARK.read_text(encoding="utf-8").splitlines()
        if not line.strip().startswith("#"))
    assert "ABORTED BEFORE A COMPLETE BASELINE" in benchmark
    assert "NOT a partial warm median" in benchmark
    assert "@($completed).Count -eq $plannedCount" in benchmark
    assert "if (-not $runComplete) {" in benchmark
    assert "exit 1" in benchmark


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
