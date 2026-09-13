#!/usr/bin/env python3
"""P10 FINAL WINDOWS ACCEPTANCE - the runner-source controls.

WHAT THIS FILE PROVES, ON LINUX, WITHOUT EXCEL. `phase10_final_acceptance.ps1`
is the contracted Phase-10 hardening runner: the one Windows session that
accepts the release workbook. Nothing here executes it. These controls read it
and prove the properties the authorisation named:

  it drives the SIX Setup commands and the FIVE structural commands through
  their real production entry points, and reads state through accessors that
  write nothing;
  it reuses the accepted machinery - the three dot-sourced files, the Stage-B
  bootstrap, the setup-only window shim - and carries the P9-1 helper copies
  byte for byte, so it holds no business logic and no model builder;
  it invokes no performance or equivalence work and asks for no stochastic
  request above the business minimum plus one;
  the Source Revision readback is mandatory, derived from the checkout and
  compared literally, with both values printed;
  every inventory is set-based against a projection, never a literal count;
  Reset preservation, Repair, protection-after-every-path, fail-fast and the
  clean COM lifecycle are asserted in the code, not promised in a comment;
  no UI automation exists; production VBA and spec are byte-identical to the
  commit this batch started from.

Runs standalone or under pytest.
"""
from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase10_final_acceptance.ps1"
# The accepted P9-1 runner: the shape this runner copies, and the authority for
# the ten helper copies the dot-sourced files call.
P9_RUNNER = WINDOWS / "phase9_p1_model_check.ps1"
BENCHMARK = WINDOWS / "phase10_benchmark.ps1"
RESET_VBA = PCCM_ROOT / "src" / "vba" / "modReset.bas"
REPAIR_VBA = PCCM_ROOT / "src" / "vba" / "modRepair.bas"
PWSH = "/opt/pwsh/pwsh"
RESOLUTION_AUDIT = PCCM_ROOT / "tests" / "powershell_command_resolution_audit.ps1"
UNINITIALISED_AUDIT = PCCM_ROOT / "tests" / "powershell_uninitialised_audit.ps1"
# The clean HEAD this batch started from. Production must be byte-identical to it.
ACCEPTED = "2d32f35"

# The helpers the dot-sourced scenario files call and this runner must define,
# exactly as the accepted P9-1 runner defines them.
COPIED_HELPERS = ("Write-RowObject", "Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
                  "Set-TableCell", "Get-TableBody", "Get-TableRowCount", "Add-BlankTableRow",
                  "Remove-TableRow", "Get-IdColumnValues")
# The window helpers copied from the benchmark: the same shim, the same reads.
SETUP_COMMANDS = ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                  "PCCM_RunAnnualStochastic", "PCCM_ResetResults", "PCCM_RepairProfiling")
STRUCTURAL_COMMANDS = ("PCCM_ApplyTimeline", "PCCM_AddCostLine", "PCCM_DeleteCostLineById",
                       "PCCM_AddRisk", "PCCM_DeleteRiskById")
# Every scenario line the authorisation asked for, and the runner must record.
REQUIRED_SCENARIOS = (
    "bootstrap", "compile", "sheets", "modules", "metadata.rows", "metadata.model-version",
    "metadata.builder-version", "metadata.build-phase", "metadata.source-revision",
    "protection.initial", "state.initial", "fixture", "structural.add-delete.",
    "structural.apply-timeline", "calculate.current", "modelcheck.calculated",
    "modelcheck.worksheet-safety", "simulation.current", "sensitivity", "annual",
    "modelcheck.simulated", "state.stale", "state.invalid", "refused.outcome.calculate",
    "modelcheck.invalid", "state.current-restored", "repair.noop", "repair.missing-row",
    "repair.order", "repair.duplicate-refused", "repair.restored", "repair.fingerprint",
    "reset.precondition", "reset.declined", "reset.confirmed", "reset.preserved",
    "reset.states", "modelcheck.after-reset", "reset.idempotent", "refused.outcome.annual",
    "reset.rollback", "protection.final", "protection.unlocked-writable",
    "protection.locked-refused", "shutdown.workbook-close", "shutdown.application-quit",
    "shutdown.natural-exit", "shutdown.no-emergency", "shutdown.com-released",
)
# Protection is asserted after every path that could leak a window.
PROTECTION_AFTER = ("protection.initial", "protection.after-structural", "protection.after-commands",
                    "protection.after-refusal", "protection.after-repair-noop",
                    "protection.after-repair-missing", "protection.after-repair-order",
                    "protection.after-repair-refusal", "protection.after-reset",
                    "protection.after-refused-annual", "protection.after-rollback",
                    "protection.final")

_MEMO: dict = {}


def _src(key: str, path: Path) -> str:
    if key not in _MEMO:
        _MEMO[key] = path.read_text(encoding="utf-8")
    return _MEMO[key]


def _runner() -> str:
    return _src("runner", RUNNER)


def _code() -> str:
    """The runner with its block comment and line comments removed. Every
    control that asserts what the runner DOES reads this."""
    if "runner_code" not in _MEMO:
        text = re.sub(r"<#.*?#>", "", _runner(), flags=re.S)
        _MEMO["runner_code"] = "\n".join(
            line for line in text.splitlines() if not line.strip().startswith("#"))
    return _MEMO["runner_code"]


def _function(name: str, source: str) -> str:
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
                return source[start: index + 1]
    raise AssertionError(f"{name} is not closed")


def _git(*args: str) -> str:
    return subprocess.run(["git", *args], cwd=PCCM_ROOT.parent,
                          capture_output=True, text=True, check=True).stdout


def _scenario_labels() -> list[str]:
    return re.findall(r"Add-FaCheck \(?'([^']+)'", _code()) + \
        re.findall(r"-Scenario \(?'([^']+)'", _code())


# ===========================================================================
# A. THE ACCEPTED SHAPE, REUSED
# ===========================================================================
def test_01_the_runner_dot_sources_exactly_the_three_accepted_files() -> None:
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([^']+)'\)", _code(), re.M)
    assert sourced == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                       "phase6_gate_b_scenarios.ps1"], sourced
    assert "Invoke-Phase5GateBScenarios" not in _code()
    assert "Invoke-Phase6GateBScenarios" not in _code()


def test_02_the_stage_b_bootstrap_builds_a_disposable_copy_under_temp() -> None:
    code = _code()
    assert "$bootstrap = Join-Path $scriptDir 'build_stage_b.ps1'" in code
    assert "& $bootstrap -BuildDir $tempRoot -Force" in code
    assert "[System.IO.Path]::GetTempPath()" in code
    assert "Add-FaCheck 'bootstrap' $bootstrapOk" in code
    # AND NEVER SAVED: the disposable copy is discarded.
    assert "$wb.Close($false)" in code
    assert ".Save(" not in code and "SaveAs" not in code


def test_03_the_ten_helper_copies_are_byte_identical_to_the_accepted_p9_runner() -> None:
    """THE DOT-SOURCED FILES CALL THESE AND DEFINE NONE OF THEM. The P9-1 copies
    are the accepted ones; a drift here is the Run-4 class of defect."""
    p9 = _src("p9", P9_RUNNER)
    for helper in COPIED_HELPERS:
        assert _function(helper, _runner()) == _function(helper, p9), helper


def test_04_the_window_helpers_are_the_benchmarks_shim_reads() -> None:
    """THE SAME SHIM, THE SAME FIVE FIELDS, THE SAME REFUSALS. The shim is never
    a manifest module, and the window is opened to depth 1 and never past it."""
    code = _code()
    assert "$script:FixtureWindowModule = 'modPhase10FixtureWindow'" in code
    assert "$script:FixtureWindowSource = 'phase10_fixture_window.bas'" in code
    importer = _function("Import-FaFixtureWindow", code)
    assert "will not import a test module over production" in importer
    assert "$Excel.Run('P10FW_Ping')" in importer
    reader = _function("Get-FaProtectionState", code)
    assert "$Excel.Run('P10FW_State')" in reader
    assert "@('applied', 'depth', 'structure', 'sheets', 'protected')" in reader
    opener = _function("Open-FaFixtureWindow", code)
    assert "$Excel.Run('P10FW_Begin')" in opener and "$state.Depth -ne 1" in opener
    assert "released WORKBOOK STRUCTURE protection" in opener
    closer = _function("Close-FaFixtureWindow", code)
    assert "$Excel.Run('P10FW_End')" in closer
    assert "Assert-FaProtectionApplied" in closer


def test_05_every_window_open_is_closed_in_a_finally() -> None:
    """A PRECONDITION MUST NOT BECOME A LEAK. Each Open-FaFixtureWindow at the
    scenario level is followed by a try whose finally closes the window."""
    code = _code()
    opens = [m.start() for m in re.finditer(r"\$null = Open-FaFixtureWindow", code)]
    assert len(opens) >= 5, len(opens)
    for start in opens:
        window = code[start: start + 2500]
        assert re.search(r"finally \{\s*\$null = Close-FaFixtureWindow", window), code[start: start + 120]


def test_06_the_stochastic_request_is_the_business_minimum_and_once_plus_one() -> None:
    code = _code()
    assert "$acceptanceIterations = [int]$gateBCases.bounds.business_minimum_iterations" in code
    assert "$staleIterations = $acceptanceIterations + 1" in code
    sets = re.findall(r"Set-NamedValue -Workbook \$wb -DefinedName \$iterationsName -Value \(\[double\](\$\w+)\)", code)
    assert sets and set(sets) == {"$acceptanceIterations", "$staleIterations"}, sets
    assert sets.count("$staleIterations") == 1, sets
    # NO LARGER REQUEST EXISTS AS A NUMBER ANYWHERE IN THE RUNNER.
    for literal in ("10000", "50000", "100000", "1000"):
        assert not re.search(rf"(?<![\w.]){literal}(?![\w.])", code), literal


def test_07_no_performance_or_equivalence_work_is_invoked() -> None:
    code = _code()
    for banned in ("phase10_benchmark", "Invoke-BenchmarkExecution", "Stopwatch", "PERF-",
                   "New-EquivalenceBundle", "Invoke-EquivalencePass", "Set-BenchmarkBulkFixture",
                   "New-BenchmarkModel", "New-BenchmarkDriver", "CalculateFull", "-FixtureMode"):
        assert banned not in code, banned


def test_08_the_model_is_the_accepted_w4_case_not_a_builder_in_the_runner() -> None:
    code = _code()
    assert "if ([string]$scenario.id -ceq 'W4') { $case = $scenario }" in code
    assert "$model = $case.model" in code
    assert "Set-Phase5Fixture -Excel $excel -Workbook $wb -Manifest $manifest" in code
    assert "Save-Phase5LockedFxSeed -Workbook $wb -Inspection $inspection" in code
    # NO DRIVER, WEIGHT OR REGISTER BLOCK IS BUILT HERE.
    for banned in ("permanent_id =", "profile_weights =", "ListRows.Add", "New-Fa", "-Drivers", "-Years"):
        assert banned not in code, banned
    # AND THE FIXTURE RUNS INSIDE THE WINDOW, CLOSED IN A FINALLY.
    fixture = code[code.index("Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'fixture'"):]
    fixture = fixture[: fixture.index("Add-FaCheck 'fixture'")]
    assert "Set-Phase5Fixture" in fixture and "finally" in fixture and "Close-FaFixtureWindow" in fixture


# ===========================================================================
# B. THE REAL ENTRY POINTS, AND NOTHING ELSE
# ===========================================================================
def test_10_every_setup_and_structural_command_is_invoked_by_its_production_name() -> None:
    code = _code()
    # PCCM_RunSimulation is reached through the accepted Phase-6 invoker, which
    # names it; every other command is named in this runner.
    phase6 = (WINDOWS / "phase6_gate_b_scenarios.ps1").read_text(encoding="utf-8")
    assert "Invoke-Phase6Simulation -Excel $excel" in code
    assert "'PCCM_RunSimulation'" in _function("Invoke-Phase6Simulation", phase6)
    # PCCM_RunAnnualStochastic is reached by the projected name, as W5 reaches it.
    import json as _json
    p7 = _json.loads((PCCM_ROOT / "build" / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert p7["command_surface"]["annual_endpoint"] == "PCCM_RunAnnualStochastic"
    assert code.count("-Operation ([string]$p7.command_surface.annual_endpoint)") == 4
    for command in SETUP_COMMANDS + STRUCTURAL_COMMANDS:
        if command in ("PCCM_RunSimulation", "PCCM_RunAnnualStochastic"):
            continue
        assert f"'{command}'" in code, command
    # THROUGH THE ACCEPTED SEAM: the two invokers begin automation and read the
    # announcement; no Run on a command bypasses them.
    invoker = _function("Invoke-FaEndpoint", code)
    assert "$Excel.Run('PCCM_AutomationBegin', $ConfirmReply, $FailAfterStage)" in invoker
    assert "$Excel.Run('PCCM_AutomationResult')" in invoker
    assert "$Excel.Run('PCCM_AutomationBegin', $true, '')" in invoker
    bare = re.findall(r"\$excel\.Run\('(PCCM_\w+)'\)", code)
    assert set(bare) <= {"PCCM_CalculationStatus", "PCCM_AutomationEnd"}, sorted(set(bare))


def test_11_states_are_read_through_accessors_that_write_nothing() -> None:
    """PCCM_CalculationStatus and PCCM_SimulationStatus persist the status they
    derive. After a reset that would rewrite a cleared cell, so every state
    observation goes through the read-only derivations."""
    states = _function("Get-FaStates", _code())
    assert "'PCCM_ModelCheckCalculationState'" in states
    assert "'SimReportDerivedStatus'" in states
    assert "'PCCM_AnnualDistributionState'" in states and "'PCCM_AnnualProfileState'" in states
    assert "PCCM_CalculationStatus" not in states and "PCCM_SimulationStatus" not in states
    # THE ONE PCCM_CalculationStatus CALL IS THE ACCEPTED COMPILE CHECK.
    assert _code().count("$excel.Run('PCCM_CalculationStatus')") == 1
    assert "PCCM_SimulationStatus" not in _code()


def test_12_the_runner_reaches_for_no_protection_of_its_own() -> None:
    code = _code()
    for banned in (".Unprotect", ".Protect(", "ProtectionRelease", "ProtectionBeginStructural",
                   "ProtectionEndStructural", "ProtectStructure =", "UserInterfaceOnly"):
        assert banned not in code, banned


def test_13_no_ui_automation_exists() -> None:
    code = _code()
    for banned in ("SendKeys", "AppActivate", "System.Windows.Forms", "OnAction", ".Click",
                   "MsgBox", "InputBox", "Shapes.Item", "mouse_event", "SetForegroundWindow",
                   "$excel.Visible = $true"):
        assert banned not in code, banned
    assert "$excel.Visible = $false" in code and "$excel.DisplayAlerts = $false" in code


# ===========================================================================
# C. SOURCE REVISION - MANDATORY, DERIVED, LITERAL
# ===========================================================================
def test_20_the_expected_revision_is_derived_the_way_the_builder_derives_it() -> None:
    reader = _function("Get-FaSourceRevision", _code())
    assert "git -C $RepoRoot rev-parse --short HEAD" in reader
    assert "git -C $RepoRoot status --porcelain --untracked-files=no" in reader
    assert "Expected = ($head + ' (clean)')" in reader
    builder = (PCCM_ROOT / "builder" / "pccm_builder" / "workbook_builder.py").read_text(encoding="utf-8")
    assert '"rev-parse", "--short", "HEAD"' in builder
    assert '"status", "--porcelain", "--untracked-files=no"' in builder


def test_21_a_dirty_tree_is_refused_before_excel_starts() -> None:
    code = _code()
    refusal = code[code.index("$revision = Get-FaSourceRevision"):]
    refusal = refusal[: refusal.index("$tempRoot = ")]
    assert "if ($revision.Dirty.Count -gt 0)" in refusal and "exit 1" in refusal
    assert "New-Object -ComObject Excel.Application" not in refusal


def test_22_the_workbook_value_is_read_from_the_projected_row_and_compared_literally() -> None:
    code = _code()
    assert "if ([string]$row.label -ceq 'Source Revision') { $sourceRevisionRow = $row }" in code
    assert "has no Source Revision row" in code
    assert "$observedRevision = $shown['Source Revision']" in code
    assert "Add-FaCheck 'metadata.source-revision' ($observedRevision -ceq $revision.Expected)" in code
    assert "('expected ' + $revision.Expected + '; observed ' + $observedRevision)" in code
    # PRINTED IN THE HEADER AND THE VERDICT, expected and observed both.
    assert "'expected source revision : ' + $revision.Expected" in code
    assert "'source revision          : expected ' + $revision.Expected + '; observed '" in code


def test_23_no_commit_hash_is_a_literal_in_the_runner() -> None:
    assert ACCEPTED not in _runner()
    assert not re.search(r"'[0-9a-f]{7,40}'", _code())
    assert not re.search(r"'[0-9a-f]{7}\s\(clean\)'", _runner())


def test_24_the_release_metadata_is_read_from_the_projection_and_the_workbook() -> None:
    code = _code()
    assert "$metadataRows = @($methodology.metadata)" in code
    assert "if ($label -cne [string]$row.label)" in code
    assert "if ($value -cne [string]$row.value)" in code
    assert "Add-FaCheck 'metadata.model-version' ($shown['PCCM Model Version'] -ceq '1.0.0')" in code
    assert "Add-FaCheck 'metadata.builder-version' ($shown['Builder Version'] -ceq '1.0.0')" in code
    assert "Add-FaCheck 'metadata.build-phase' ($shown['Build Phase'] -ceq 'Release 1.0 - Production')" in code


# ===========================================================================
# D. INVENTORIES ARE SETS AGAINST PROJECTIONS, NEVER LITERAL COUNTS
# ===========================================================================
def test_30_the_module_check_compares_the_observed_set_with_the_manifest_set() -> None:
    code = _code()
    block = code[code.index("$expectedModules = @{}"):]
    block = block[: block.index("Add-FaCheck 'modules'") + 200]
    assert "foreach ($module in @($manifest.vba.modules)) { $expectedModules[[string]$module.name] = $true }" in block
    assert "$expectedModules[[string]$manifest.vba.document_module.component] = $true" in block
    assert "foreach ($sheet in @($manifest.sheets)) { $expectedModules[[string]$sheet.codename] = $true }" in block
    assert "if (-not $observedModules.ContainsKey($name)) { $moduleProblems += ('missing ' + $name) }" in block
    assert "if (-not $expectedModules.ContainsKey($name)) { $moduleProblems += ('undeclared ' + $name) }" in block
    assert "Add-FaCheck 'modules' ($moduleProblems.Count -eq 0)" in block
    # REPORTED AFTER DERIVING, NEVER COMPARED TO A LITERAL.
    assert "$observedModules.Count + ' components observed" in block
    # AND THE SET IS READ BEFORE THE SHIM IS IMPORTED.
    assert code.index("Add-FaCheck 'modules'") < code.index("Import-FaFixtureWindow -Excel $excel")


def test_31_sheets_and_codenames_are_a_set_against_the_manifest() -> None:
    code = _code()
    assert "foreach ($sheet in @($manifest.sheets)) { $expectedSheets[[string]$sheet.name] = [string]$sheet.codename }" in code
    assert "$observedSheets[[string]$ws.Name] = [string]$ws.CodeName" in code
    assert "$sheetProblems += ('undeclared ' + $name)" in code
    assert "Add-FaCheck 'sheets' ($sheetProblems.Count -eq 0)" in code


def test_32_no_inventory_is_compared_to_a_literal_count() -> None:
    code = _code()
    for literal in ("33", "14", "11", "32"):
        assert not re.search(rf"-(eq|ne|ge|le|gt|lt) {literal}\b", code), literal
        assert not re.search(rf"\.Count -(eq|ne) {literal}\b", code), literal
    assertion = _function("Assert-FaProtectionApplied", code)
    assert "$expected = @($Protection.sheets).Count" in assertion
    assert "$state.Sheets -ne $expected" in assertion and "$state.Protected -ne $expected" in assertion
    assert "$state.Depth -ne 0" in assertion and "-not $state.Structure" in assertion


# ===========================================================================
# E. RESET RESULTS
# ===========================================================================
def test_40_the_preserved_digest_covers_every_contracted_preservation() -> None:
    digest = _function("Get-FaPreservedDigest", _code())
    for required in ("$Reset.preserved.editable_inputs", "$Reset.preserved.applied_timeline",
                     "$Reset.preserved.permanent_id_counters", "$Reset.publications.simulation.preserved",
                     "$identity.next_auto_nonce", "$identity.last_run_id", "$identity.pending_auto_nonce",
                     "-Address $MetadataRange"):
        assert required in digest, required


def test_41_the_cleared_rectangles_come_from_the_reset_projection_and_are_all_required_blank() -> None:
    code = _code()
    rects = _function("Get-FaClearedRectangles", code)
    for required in ("$calc.cleared.state", "$calc.cleared.totals", "$cleared.bank_snapshots.$bank",
                     "$cleared.summary.$bank", "$cleared.contingency.$bank", "$cleared.annual_stamps.$bank",
                     "$cleared.sensitivity_stamps.$bank", "$cleared.iteration_banks.$bank",
                     "$cleared.annual_blocks.$bank", "$cleared.sensitivity_records.$bank",
                     "$cleared.attempt_and_selector"):
        assert required in rects, required
    uncleared = _function("Get-FaUnclearedRectangles", code)
    assert "Test-FaBlockBlank" in uncleared and "$calc.cleared.tables" in uncleared
    assert "Add-FaCheck 'reset.confirmed'" in code and "($uncleared.Count -eq 0)" in code
    assert "'OK|Results reset. Model inputs and identity counters were preserved.'" in code
    assert "Results reset. Model inputs and identity counters were preserved." in _src("reset", RESET_VBA)


def test_42_the_reset_scenarios_are_declined_confirmed_idempotent_and_rolled_back() -> None:
    code = _code()
    assert "Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_ResetResults' -ConfirmReply $false" in code
    assert "'*Reset Results clears every published result*'" in code
    assert "Add-FaCheck 'reset.declined'" in code and "($publicationDeclined -ceq $publicationBefore)" in code
    assert "Add-FaCheck 'reset.preserved' ($preservedAfter -ceq $preservedBefore)" in code
    assert "Add-FaCheck 'reset.idempotent'" in code and "($again -ceq $confirmed)" in code
    assert "Add-FaCheck 'reset.states'" in code
    assert "($statesReset.Calculation -ceq $statusNotCalculated) -and ($statesReset.Simulation -eq '')" in code
    assert "$script:ResetFailpoint = 'Phase10ResetSimulation'" in code
    assert 'FAILPOINT_RESET_SIMULATION As String = "Phase10ResetSimulation"' in _src("reset", RESET_VBA)
    assert "-FailAfterStage $script:ResetFailpoint" in code
    rollback = code[code.index("Add-FaCheck 'reset.rollback'"):]
    rollback = rollback[: rollback.index("Assert-FaProtectionApplied")]
    for required in ("($injected -like 'FAIL|*')", "($publicationRolledBack -ceq $publicationFull)",
                     "($simStateRolledBack -ceq $simStateFull)", "($preservedRolledBack -ceq $preservedFull)",
                     "($statesRolledBack.Calculation -ceq $statusCurrent)"):
        assert required in rollback, required
    assert "'*put back*'" in rollback
    assert "Every publication this command had cleared was put back" in _src("reset", RESET_VBA)


# ===========================================================================
# F. REPAIR PROFILING
# ===========================================================================
def test_50_the_four_repair_scenarios_assert_the_contract() -> None:
    code = _code()
    noop = code[code.index("$noop = Invoke-FaEndpoint"):]
    noop = noop[: noop.index("Add-FaCheck 'repair.noop'") + 400]
    assert "'*Nothing was changed*'" in noop and "-ceq $inputFingerprint0" in noop
    assert "Nothing was changed." in _src("repair", REPAIR_VBA)
    missing = code[code.index("Add-FaCheck 'repair.missing-row'") - 2600: code.index("Add-FaCheck 'repair.missing-row'")]
    assert "Remove-TableRow -Workbook $wb -SheetName $gridSheet -TableName $gridTable -RowIndex $missingRowIndex" in missing
    assert "$repairProblems += ($missingId + ' was not restored')" in missing
    assert "', not blank')" in missing and "$repairProblems += ($id + ' changed')" in missing
    assert "are not the register order" in missing
    order = code[code.index("Add-FaCheck 'repair.order'"):]
    order = order[: order.index("Assert-FaProtectionApplied")]
    assert "($orderDigest -ceq $gridDigest2)" in order and "($inputFingerprint2 -ceq $inputFingerprint1)" in order
    duplicate = code[code.index("Add-FaCheck 'repair.duplicate-refused'"):]
    duplicate = duplicate[: duplicate.index("Assert-FaProtectionApplied")]
    assert "($repairDuplicate -like 'FAIL|*')" in duplicate
    assert "('*' + $duplicateId + '*')" in duplicate and "'*more than one row*'" in duplicate
    assert "($afterRefusal -ceq $damagedDigest)" in duplicate
    assert "on more than one row, and" in _src("repair", REPAIR_VBA)
    assert "Add-FaCheck 'repair.fingerprint'" in code and "($fingerprintAfterRepairs -ceq $fingerprint)" in code


def test_51_every_repair_precondition_is_created_inside_the_window_and_repaired_outside_it() -> None:
    """THE WINDOW MAKES THE DEFECT; PRODUCTION REPAIRS UNDER ITS OWN PROTECTION."""
    code = _code()
    for scenario in ("repair.missing-row", "repair.order", "repair.duplicate", "repair.duplicate-undo"):
        open_at = code.index(f"Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario '{scenario}'")
        close_at = code.index(f"Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario '{scenario}'")
        assert open_at < close_at
        assert "'PCCM_RepairProfiling'" not in code[open_at: close_at], scenario
        # The window closes BEFORE the next repair invocation.
        next_repair = code.index("'PCCM_RepairProfiling'", close_at)
        assert "Open-FaFixtureWindow" not in code[close_at: next_repair], scenario


# ===========================================================================
# G. STATES, REFUSAL, PROTECTION AFTER EVERY PATH, FAIL FAST, SHUTDOWN
# ===========================================================================
def test_60_the_state_vocabulary_is_projected_and_the_four_states_are_observed() -> None:
    code = _code()
    assert "$statusNotCalculated = [string]$p7.model_states.derived_status[0]" in code
    assert "$statusCurrent       = [string]$p7.model_states.derived_status[1]" in code
    assert "$statusStale         = [string]$p7.model_states.derived_status[2]" in code
    assert "$statusInvalid       = [string]$p7.model_states.derived_status[3]" in code
    assert "$attemptRefused      = [string]$p7.model_states.attempt_result[2]" in code
    for check, word in (("state.initial", "$statusNotCalculated"), ("calculate.current", "$statusCurrent"),
                        ("state.stale", "$statusStale"), ("state.invalid", "$statusInvalid")):
        tail = code[code.index(f"Add-FaCheck '{check}'"): code.index(f"Add-FaCheck '{check}'") + 400]
        assert word in tail, check
    assert "for literal in" not in code
    for word in ("'NOT CALCULATED'", "'CURRENT'", "'STALE'", "'INVALID'", "'REFUSED'"):
        assert word not in code, word


def test_61_refused_is_observed_only_as_an_attempt_outcome() -> None:
    code = _code()
    calc = code[code.index("Add-FaCheck 'refused.outcome.calculate'"):]
    calc = calc[: calc.index("$excel.Calculate()")]
    assert "($attemptInvalid -ceq $attemptRefused)" in calc
    assert "($derivedVocabulary -cnotcontains $attemptRefused)" in calc
    annual = code[code.index("Add-FaCheck 'refused.outcome.annual'"):]
    annual = annual[: annual.index("Assert-FaProtectionApplied")]
    assert "($refusedAnnual -like 'FAIL|*')" in annual
    assert "($statesAfterRefusal.Calculation -ceq $statusNotCalculated)" in annual


def test_62_protection_is_asserted_after_every_success_refusal_and_error_path() -> None:
    labels = _scenario_labels()
    for scenario in PROTECTION_AFTER:
        assert scenario in labels, scenario
    code = _code()
    # ORDER: the refusal, the repair refusal, the reset, the refused annual and the
    # rollback are each followed by an assertion before the next scenario.
    for after, assertion in (("Add-FaCheck 'state.current-restored'", "protection.after-refusal"),
                             ("Add-FaCheck 'repair.duplicate-refused'", "protection.after-repair-refusal"),
                             ("Add-FaCheck 'modelcheck.after-reset'", "protection.after-reset"),
                             ("Add-FaCheck 'refused.outcome.annual'", "protection.after-refused-annual"),
                             ("Add-FaCheck 'reset.rollback'", "protection.after-rollback")):
        start = code.index(after)
        assert f"-Scenario '{assertion}'" in code[start: start + 1500], assertion


def test_63_the_protection_behaviour_pair_is_asserted() -> None:
    code = _code()
    assert "foreach ($candidate in @($setupProtection.unlocked))" in code
    assert "Add-FaCheck 'protection.unlocked-writable' ($unlockedFailure -eq '')" in code
    assert "$lockedAddress = [string]$methodology.label_column + [string]$sourceRevisionRow.row" in code
    assert "catch { $lockedRefused = $true" in code
    assert "Add-FaCheck 'protection.locked-refused' $lockedRefused $lockedDetail" in code


def test_64_every_required_scenario_line_is_recorded_and_fail_fast_holds() -> None:
    labels = _scenario_labels()
    for scenario in REQUIRED_SCENARIOS:
        assert any(label.startswith(scenario) for label in labels), scenario
    check = _function("Add-FaCheck", _code())
    assert "Write-FaLine ($verdict + '|' + $Scenario + '|' + $Detail)" in check
    assert "if ((-not $Ok) -and (-not $Continue)) {" in check
    assert "throw ('ACCEPTANCE FAILED at ' + $Scenario + ': ' + $Detail)" in check
    # -Continue is reserved for the shutdown checks, which run after the session.
    continued = re.findall(r"Add-FaCheck '([^']+)'[^\n]*-Continue", _code())
    assert continued and all(label.startswith("shutdown.") for label in continued), continued
    assert "$fatal = (Format-Err $_)" in _code() and "'FINAL ACCEPTANCE ' + $(if ($ok) { 'PASS' } else { 'FAIL' })" in _code()
    assert "if ($ok) { exit 0 } else { exit 1 }" in _code()


def test_65_the_shutdown_is_the_accepted_lifecycle_and_is_asserted() -> None:
    code = _code()
    shutdown = code[code.index("} finally {"):]
    for required in ("$wb.Close($false)", "$rel.WorkbookClosed = $true", "$excel.Quit()", "$rel.QuitCalled = $true",
                     "Wait-ExcelExit -Identity $excelIdentity -TimeoutSeconds 90",
                     "Invoke-EmergencyExcelCleanup -Identity $excelIdentity", "Format-ReleaseLedger -Ledger $rel",
                     "Get-TransientFailures"):
        assert required in shutdown, required
    assert shutdown.index("$wb.Close($false)") < shutdown.index("Invoke-FaRelease -Ledger $rel -Obj $workbooks") \
        < shutdown.index("$excel.Quit()")
    for check in ("shutdown.workbook-close' $rel.WorkbookClosed", "shutdown.application-quit' $rel.QuitCalled",
                  "shutdown.natural-exit' $rel.NaturalExit", "shutdown.no-emergency' (-not $rel.EmergencyRequired)",
                  "shutdown.com-released'"):
        assert check in shutdown, check
    assert "([int]$rel.Attempted -eq $comAcquired)" in shutdown and "($rel.Failed.Count -eq 0)" in shutdown
    assert "$excel.Run('PCCM_AutomationEnd')" in code


# ===========================================================================
# H. THE TREE
# ===========================================================================
def test_70_production_vba_spec_and_builder_are_byte_identical_to_the_accepted_head() -> None:
    changed = _git("diff", "--name-only", ACCEPTED, "--", "pccm/src", "pccm/spec", "pccm/builder").strip()
    assert changed == "", changed


def test_71_the_runner_is_declared_and_documented() -> None:
    benchmark_suite = (PCCM_ROOT / "tests" / "test_phase10_benchmark_harness.py").read_text(encoding="utf-8")
    assert '"phase10_final_acceptance.ps1",' in benchmark_suite
    readme = (WINDOWS / "README.md").read_text(encoding="utf-8")
    assert "`phase10_final_acceptance.ps1`" in readme


def test_72_the_runner_keeps_crlf_and_no_ps7_only_construct() -> None:
    raw = RUNNER.read_bytes()
    assert b"\r\n" in raw and b"\r\n" not in raw.replace(b"\r\n", b"")
    code = _code()
    for banned, pattern in (("ternary", r"[^'\"]\s\?\s[^?]"), ("null-coalescing", r"\?\?"),
                            ("pipeline chain", r"(?<![&|])(&&|\|\|)(?![&|])")):
        assert not re.search(pattern, code), banned


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_73_every_reachable_command_resolves_exactly_once() -> None:
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(RESOLUTION_AUDIT), "-Path", str(RUNNER)],
                          capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout
    assert done.stdout.strip().startswith("CLEAN"), done.stdout


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_74_no_variable_can_be_read_before_assignment() -> None:
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(UNINITIALISED_AUDIT), "-Path", str(RUNNER)],
                          capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout
    assert done.stdout.strip() == "CLEAN", done.stdout


if __name__ == "__main__":
    failures = 0
    for name in sorted(n for n in dir() if n.startswith("test_")):
        try:
            globals()[name]()
            print(f"PASS {name}")
        except BaseException as exc:  # noqa: BLE001
            failures += 1
            print(f"FAIL {name}: {exc}")
    sys.exit(1 if failures else 0)
