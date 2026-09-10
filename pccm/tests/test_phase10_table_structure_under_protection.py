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

import pytest  # noqa: E402
import yaml  # noqa: E402

PROBE = BOOTSTRAP / "phase10_protection_probe.ps1"
BENCHMARK = BOOTSTRAP / "phase10_benchmark.ps1"
ACCEPTED = "4ad035a"

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
    for endpoint in ("PCCM_ApplyTimeline", "PCCM_AddCostLine", "PCCM_AddRisk",
                     "PCCM_Calculate"):
        assert f"-Endpoint '{endpoint}'" in code, endpoint
    assert "$Excel.Run($Endpoint)" in code
    assert "PCCM_AutomationResult" in code


def test_21_the_probe_separates_the_two_capabilities() -> None:
    """A VALUE WRITE AND A TABLE STRUCTURAL OPERATION ARE DIFFERENT THINGS, and
    Windows Run 3 showed one working and the other refused. The probe tests the
    control case deliberately rather than as a side effect."""
    code = _probe_code()
    assert "CONTROL - CAN CODE WRITE A VALUE TO A LOCKED CELL?" in code
    assert "Set-ProbeNamedValue -Workbook $wb -DefinedName $discountName" in code
    assert "UserInterfaceOnly is honoured for code that writes VALUES" in code


def test_22_the_probe_checks_protection_is_actually_in_force() -> None:
    """OTHERWISE IT PROVES NOTHING. A workbook that came back unprotected would
    make every command below succeed for the wrong reason."""
    code = _probe_code()
    assert "function Get-ProbeProtectionState" in code
    assert "$before = Get-ProbeProtectionState -Workbook $wb" in code
    assert "$after = Get-ProbeProtectionState -Workbook $wb" in code
    assert "THE PROBE IS INCONCLUSIVE: the workbook did not come back protected" in code
    assert "ProtectionIsApplied" in code


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
# D. NOTHING WAS CHANGED WHILE THE QUESTION IS OPEN
# ===========================================================================
def test_30_no_production_source_changed() -> None:
    """THE CLASSIFICATION IS ESTABLISHED BEFORE THE CODE MOVES. Editing four
    accepted modules on a reading of the documentation is exactly what the
    Windows evidence exists to prevent."""
    changed = [line for line in _git("diff", "--name-only", ACCEPTED, "--",
                                     "pccm/src", "pccm/spec").splitlines() if line.strip()]
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
