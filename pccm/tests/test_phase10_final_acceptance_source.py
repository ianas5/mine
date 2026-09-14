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
HANDLER_VBA = PCCM_ROOT / "src" / "vba" / "ThisWorkbook.vba"
APPSTATE_VBA = PCCM_ROOT / "src" / "vba" / "modAppState.bas"
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
    "protection.initial", "state.initial.persisted", "state.initial.live", "fixture", "structural.add-delete.",
    "structural.apply-timeline", "calculate.current", "modelcheck.calculated",
    "modelcheck.worksheet-safety", "simulation.current", "sensitivity", "annual",
    "modelcheck.simulated", "state.stale", "state.invalid", "refused.outcome.calculate",
    "modelcheck.invalid", "modelcheck.invalid.subject", "state.invalid.annual-historical", "state.current-restored", "repair.noop", "repair.missing-row",
    "repair.order", "repair.duplicate-refused", "repair.restored", "repair.fingerprint",
    "reset.precondition", "reset.declined", "reset.confirmed", "reset.preserved",
    "reset.states", "modelcheck.after-reset", "modelcheck.after-reset.adapter", "reset.idempotent", "refused.outcome.annual",
    "reset.rollback", "protection.final", "protection.unlocked-writable",
    "protection.locked-cell.user-protected", "protection.locked-cell.code-write-permitted",
    "shutdown.workbook-close", "shutdown.application-quit",
    "shutdown.natural-exit", "shutdown.no-emergency", "shutdown.com-released",
    # Added at the bounded correction round after the independent review: the
    # Repair contract scenarios of section 5 / matrix rows F, G, G2, M, and the
    # distribution-copy session of matrix row J.
    "repair.width-growth", "repair.shrink-blank", "repair.shrink-zero-refused",
    "repair.shrink-nonzero-refused", "repair.semantic-non1-refused", "repair.blank-profile-allowed",
    "repair.signed-zero-total-refused", "repair.rollback.", "repair.grids-restored",
    "copy.compile", "copy.source-revision", "copy.protection", "copy.fixture", "copy.run",
    # Added at the P10-R3 closure: contract matrix row O, the injected
    # Workbook_Open failure, in a third disposable session.
    "rowo.open-suppressed", "rowo.handler-entered", "rowo.failure-path", "rowo.disclosed-once",
    "rowo.screen-updating-restored", "rowo.calculation-unchanged", "rowo.events-unchanged",
    "rowo.released-not-half-protected", "rowo.usable", "rowo.reopen-silent", "rowo.reopen-applied",
)
# The runner the starting authority of the bounded correction round carried, and
# that Windows executed at final acceptance run 7 (ee6e9fb; the runner is
# byte-identical between the two). Everything this round changed in the runner
# is declared below and proved by reversal against it.
STARTING_AUTHORITY = "347f42e"
# The starting authority of the P10-R3 closure round, whose runner is the R2
# runner: the row-O session is one more delimited block on top of it.
R3_AUTHORITY = "889b6b5"
R2_BLOCKS = ("P10-R2 REPAIR CONTRACT SCENARIOS", "P10-R2 DISTRIBUTION COPY")
R3_BLOCKS = ("P10-R3 WORKBOOK_OPEN FAILURE",)
# The one region of the executed tail that was SUBSTITUTED rather than inserted:
# the final protection behaviour, from its heading to the automation end.
PROTECTION_REGION = ("    # 18/19. FINAL PROTECTION AND ITS BEHAVIOUR.",
                     "    $excel.Run('PCCM_AutomationEnd') | Out-Null\n")
# Protection is asserted after every path that could leak a window.
PROTECTION_AFTER = ("protection.initial", "protection.after-structural", "protection.after-commands",
                    "protection.after-refusal", "protection.after-repair-noop",
                    "protection.after-repair-missing", "protection.after-repair-order",
                    "protection.after-repair-refusal", "protection.after-reset",
                    "protection.after-refused-annual", "protection.after-rollback",
                    "protection.final",
                    "protection.after-repair-growth", "protection.after-repair-shrink",
                    "protection.after-repair-shrink-zero", "protection.after-repair-shrink-nonzero",
                    "protection.after-repair-semantic", "protection.after-repair-signed",
                    "protection.after-repair-rollback.", "protection.after-locked-cell",
                    "copy.protection", "copy.protection-after-run")

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


def _without_r2_blocks(text: str, blocks: tuple[str, ...] = R2_BLOCKS + R3_BLOCKS) -> str:
    """The runner with the delimited insertions of the bounded correction round
    and the closure round taken out - each must be present exactly once, begin
    and end."""
    for name in blocks:
        begin = f"    # --- {name}: begin"
        end = f"    # --- {name}: end"
        assert text.count(begin) == 1 and text.count(end) == 1, name
        start = text.index(begin)
        stop = text.index("\n", text.index(end)) + 1
        text = text[:start] + text[stop:]
    return text


def _split_protection_region(tail: str) -> tuple[str, str, str]:
    start = tail.index(PROTECTION_REGION[0])
    stop = tail.index(PROTECTION_REGION[1], start)
    return tail[:start], tail[start:stop], tail[stop:]


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
    # THE PCCM_CalculationStatus CALL IS THE ACCEPTED COMPILE CHECK: once per
    # opened workbook - the original, and (since the bounded correction round)
    # the distribution copy - each immediately after the open and before any
    # state is observed in that session.
    code = _code()
    calls = [m.start() for m in re.finditer(re.escape("$excel.Run('PCCM_CalculationStatus')"), code)]
    assert len(calls) == 2, len(calls)
    opened = [m.start() for m in re.finditer(re.escape("$workbooks.Open("), code)]
    assert len(opened) == 3, len(opened)
    for open_at, call_at in zip(opened, calls):
        assert open_at < call_at
        assert "Get-FaStates" not in code[open_at: call_at] and "Get-FaPersistedStates" not in code[open_at: call_at]
    # The third open (the P10-R3 row-O session) runs no compile check and reads
    # no state: its first Run is the shim's ping, and the status writer never runs there.
    assert "PCCM_CalculationStatus" not in code[opened[2]:]
    assert "PCCM_SimulationStatus" not in code


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


def test_50b_the_missing_row_expectation_requires_blank_weights_and_never_a_zero() -> None:
    """RUN 7's GENUINE FINDING. The runner requires the recreated row's
    project-year cells to be BLANK, and that is the settled §5 contract; it is
    not weakened to accept the owners' zero."""
    code = _code()
    missing = code[code.index("$repairMissing = Invoke-FaEndpoint"):code.index("Add-FaCheck 'repair.missing-row'")]
    assert "if ([string]$restoredRow[$c] -ne '') { $repairProblems += ($missingId + ' column ' + [string]($c + 1) + ' is ' + [string]$restoredRow[$c] + ', not blank') }" in missing
    assert "-ne '0'" not in missing and "-ne 0" not in missing
    contract = (PCCM_ROOT / "docs" / "phase10_step1_contract.md").read_text(encoding="utf-8")
    assert "create it, id only, **weights blank** — a blank is an unmade assumption, not a zero" in contract
    assert "extend with blanks" in contract


def test_51_every_repair_precondition_is_created_inside_the_window_and_repaired_outside_it() -> None:
    """THE WINDOW MAKES THE DEFECT; PRODUCTION REPAIRS UNDER ITS OWN PROTECTION."""
    code = _code()
    for scenario in ("repair.missing-row", "repair.order", "repair.duplicate", "repair.duplicate-undo",
                     "repair.width-growth", "repair.shrink-blank", "repair.shrink-zero", "repair.shrink-zero-clear",
                     "repair.shrink-nonzero", "repair.shrink-nonzero-clear", "repair.rollback"):
        open_at = code.index(f"Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario '{scenario}'")
        close_at = code.index(f"Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario '{scenario}'")
        assert open_at < close_at
        assert "'PCCM_RepairProfiling'" not in code[open_at: close_at], scenario
        # The window closes BEFORE the next repair invocation.
        next_repair = code.index("'PCCM_RepairProfiling'", close_at)
        assert "Open-FaFixtureWindow" not in code[close_at: next_repair], scenario
    # The two column helpers are structural and are called ONLY inside a window.
    for helper in ("Add-FaTableColumn -Workbook", "Remove-FaLastTableColumn -Workbook"):
        for match in re.finditer(re.escape(helper), code):
            opened = code.rfind("Open-FaFixtureWindow", 0, match.start())
            closed = code.rfind("Close-FaFixtureWindow", 0, match.start())
            assert opened > closed, (helper, match.start())


def test_52_width_growth_regrows_blank_and_the_blank_only_shrink_changes_nothing_else() -> None:
    """MATRIX ROWS F AND G, ADDED AT THE BOUNDED CORRECTION ROUND. A project
    year lost inside the window is regrown BLANK on every row with no existing
    cell touched; a blank project year added inside the window is trimmed and
    the grid is byte-identical to before it was added."""
    code = _code()
    growth = code[code.index("$growthBefore = @(Get-TableBody"): code.index("Add-FaCheck 'repair.width-growth'") + 300]
    assert "Remove-FaLastTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable" in growth
    assert "$grown = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'" in growth
    assert "if ($grown -notlike 'OK|*')" in growth and "if ($grown -notlike '*left blank*')" in growth
    assert "if ([string]$after[$c] -cne [string]$before[$c])" in growth
    assert "if ([string]$after[$before.Count - 1] -ne '') { $growthProblems += ([string]$before[0] + ' regrown project year '" in growth
    assert "-ne '0'" not in growth and "-ne 0)" not in growth
    assert "Add-FaCheck 'repair.width-growth' ($growthProblems.Count -eq 0)" in growth
    assert "are left blank for you to complete" in _src("repair", REPAIR_VBA)
    shrink = code[code.index("Scenario 'repair.shrink-blank'"): code.index("Add-FaCheck 'repair.shrink-blank'") + 200]
    assert "$null = Add-FaTableColumn -Workbook $wb -SheetName $gridSheet -TableName $gridTable" in shrink
    assert "Add-FaCheck 'repair.shrink-blank' (($shrunkBlank -like 'OK|*') -and ($costAfterShrink -ceq $script:FaCostBefore))" in shrink
    # the setup for growth leaves every retained row at 100%, so only structure is repaired
    setup = code[code.index("# (a) WIDTH GROWTH") if "# (a) WIDTH GROWTH" in code else code.index("$rowOneId = [string]$rowOneWeights[0]"): code.index("$growthBefore = @(Get-TableBody")]
    assert "-RowIndex 1 -Year 1 -Weight 0.5" in setup and "-RowIndex 1 -Year 2 -Weight 0.5" in setup


def test_53_every_repair_refusal_is_asserted_with_both_grids_unchanged_and_names_the_id_and_year() -> None:
    """MATRIX ROWS G2 AND M, ADDED AT THE BOUNDED CORRECTION ROUND. A typed
    zero beyond the applied duration, a nonzero weight beyond it, a populated
    fixed-width row that is not 100%, and a populated signed row totalling zero
    are each REFUSED - the announcement is a FAIL, both profiling grids are
    byte-identical to before the call, and the refusal names the permanent id
    (and, for a trim, the project year). An all-blank row passes. Every wording
    the runner expects is production's."""
    code = _code()
    unchanged = _function("Assert-FaGridsUnchanged", code)
    assert "($Announcement -like 'FAIL|*')" in unchanged
    assert "($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)" in unchanged
    saver = _function("Save-FaGridsBefore", code)
    assert "$script:FaCostBefore = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable" in saver
    assert "$script:FaRiskBefore = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable" in saver
    repair = _src("repair", REPAIR_VBA)
    for scenario, variable, wordings in (
            ("repair.shrink-zero-refused", "$zeroRefused",
             ("('*project year ' + [string]($durationYears + 1) + '*')", "'*typed zero counts as populated*'")),
            ("repair.shrink-nonzero-refused", "$nonzeroRefused",
             ("('*project year ' + [string]($durationYears + 1) + '*')",)),
            ("repair.semantic-non1-refused", "$non1Refused", ("'*populated and total*'", "'*not 100%*'")),
            ("repair.signed-zero-total-refused", "$signedRefused", ("'*populated and total 0*'", "'*not 100%*'"))):
        assert f"Assert-FaGridsUnchanged -Scenario '{scenario}' -Announcement {variable}" in code, scenario
        names = code[code.index(f"Add-FaCheck '{scenario}.names'"):]
        names = names[: names.index("Assert-FaProtectionApplied")]
        assert f"({variable} -like ('*' + $rowOneId + '*'))" in names, scenario
        for wording in wordings:
            assert f"({variable} -like {wording})" in names, (scenario, wording)
        # the digests are saved AFTER the precondition and BEFORE the call
        call = code.index(f"{variable} = Invoke-FaEndpoint -Excel $excel -Operation 'PCCM_RepairProfiling'")
        saved = max(code.rfind("Save-FaGridsBefore\n", 0, call), code.rfind("$script:FaCostBefore = $costWithZero", 0, call))
        assert saved > 0 and "Set-FaWeight" not in code[saved: call], scenario
    assert "A typed zero counts as populated." in repair
    assert "are populated and total " in repair and "which is not 100%." in repair
    assert "even when those weights total zero" in repair
    # the signed row is [1, -1, 0, 0]: populated, zero total, never blank
    signed = code[code.index("# (g) POPULATED ZERO-TOTAL SIGNED PROFILE") if "# (g)" in code else code.index("Add-FaCheck 'repair.blank-profile-allowed'"): code.index("$signedRefused = Invoke-FaEndpoint")]
    assert "-RowIndex 1 -Year 1 -Weight 1.0" in signed and "-RowIndex 1 -Year 2 -Weight -1.0" in signed
    blank = code[code.index("Add-FaCheck 'repair.blank-profile-allowed'"):]
    blank = blank[: blank.index("\n")]
    assert "(($blankAllowed -like 'OK|*') -and ($blankAllowed -like '*Nothing was changed*') -and ($costAfterBlank -ceq $script:FaCostBefore))" in blank
    # the fixed-width non-100% case has no width drift: no column helper between the weight and the call
    non1 = code[code.rfind("Set-FaWeight", 0, code.index("$non1Refused = Invoke-FaEndpoint")): code.index("$non1Refused = Invoke-FaEndpoint")]
    assert "TableColumn" not in non1 and "-Weight 0.55" in code[: code.index("$non1Refused = Invoke-FaEndpoint")]


def test_54_both_repair_failpoints_roll_back_and_the_grids_return_to_their_baseline() -> None:
    """CONTRACT SECTION 5 TRANSACTION, ADDED AT THE BOUNDED CORRECTION ROUND.
    Both contracted Repair failpoints - the runner's names are production's
    constants - turn a repairable fault into a FAIL whose rollback leaves BOTH
    grids byte-identical to before the call; the real repair then recreates the
    row blank; and after the scenarios both grids equal their baseline digests."""
    code = _code()
    repair = _src("repair", REPAIR_VBA)
    for script_name, constant in (("RepairFailpointCost", "FAILPOINT_REPAIR_COST"),
                                  ("RepairFailpointRisk", "FAILPOINT_REPAIR_RISK")):
        value = re.search(rf'^Public Const {constant} As String = "([^"]+)"', repair, re.M).group(1)
        assert f"$script:{script_name} = '{value}'" in code, (script_name, value)
    rollback = code[code.index("foreach ($failpoint in @($script:RepairFailpointCost, $script:RepairFailpointRisk))"):]
    rollback = rollback[: rollback.index("Add-FaCheck 'repair.grids-restored'") + 400]
    assert "Save-FaGridsBefore" in rollback
    assert "-Operation 'PCCM_RepairProfiling' -FailAfterStage $failpoint" in rollback
    assert "(($rolledBack -like 'FAIL|*') -and ($rolledBack -like ('*' + $failpoint + '*')) -and ($rolledBack -like '*restored to the state they were in before*')" in rollback
    assert "($costNow -ceq $script:FaCostBefore) -and ($riskNow -ceq $script:FaRiskBefore)" in rollback
    assert "restored to the state they were in before the" in repair
    assert "-Scenario ('protection.after-repair-rollback.' + $failpoint)" in rollback
    assert "Add-FaCheck 'repair.rollback.then-repaired' (($repairedAfterRollback -like 'OK|*') -and $rowTwoBlank)" in rollback
    assert "if ([string]$rowTwo[$c] -ne '') { $rowTwoBlank = $false }" in rollback
    assert "Add-FaCheck 'repair.grids-restored' (($costRestored -ceq $baselineCost) -and ($riskRestored -ceq $baselineRisk))" in rollback
    assert "$baselineCost = Get-FaTableDigest -Workbook $wb -SheetName $gridSheet -TableName $gridTable" in code
    assert "$baselineRisk = Get-FaTableDigest -Workbook $wb -SheetName $riskSheet -TableName $riskTable" in code
    # the grid variables are initialised at script scope before any helper runs
    preamble = code[: code.index("function ")]
    assert "$script:FaCostBefore = ''" in preamble and "$script:FaRiskBefore = ''" in preamble


def test_55_the_distribution_copy_is_a_renamed_copy_of_the_built_file_run_in_a_second_session_and_never_saved() -> None:
    """MATRIX ROW J, ADDED AT THE BOUNDED CORRECTION ROUND. The original is
    closed unsaved; the BUILT Stage-B file is copied under a different name in a
    different folder; the copy is opened in the same owned Excel (so
    Workbook_Open runs in it), proved to compile, to carry the expected Source
    Revision, to be protected, to take the fixture and to run to CURRENT; it is
    closed unsaved by the accepted shutdown, and its COM acquisition is counted."""
    code = _code()
    copy = code[code.index("try { $wb.Close($false); $rel.WorkbookClosed = $true }"): code.index("Add-FaCheck 'copy.run'") + 600]
    assert "Invoke-FaRelease -Ledger $rel -Obj $wb -Label 'Workbook(original)'" in copy
    assert copy.index("$wb = $null") < copy.index("Copy-Item")
    assert "$copyDir = Join-Path $tempRoot 'distribution copy'" in copy
    assert "$copyPath = Join-Path $copyDir 'PCCM distribution copy.xlsm'" in copy
    assert "Copy-Item -LiteralPath $stageBPath -Destination $copyPath -Force" in copy
    assert "$wb = $workbooks.Open($copyPath)" in copy and "$comAcquired = $comAcquired + 1" in copy
    assert copy.index("Copy-Item") < copy.index("$wb = $workbooks.Open($copyPath)") < copy.index("$comAcquired = $comAcquired + 1")
    assert "try { $null = $excel.Run('PCCM_CalculationStatus') } catch { $copyCompile = (Format-Err $_) }" in copy
    assert "Add-FaCheck 'copy.compile' ($copyCompile -eq '')" in copy
    assert "Add-FaCheck 'copy.source-revision' ($copyRevision -ceq $revision.Expected)" in copy
    assert "Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'copy.protection'" in copy
    assert "Import-FaFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir" in copy
    assert copy.index("Open-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'copy.fixture'") \
        < copy.index("Set-Phase5Fixture -Excel $excel -Workbook $wb") \
        < copy.index("Close-FaFixtureWindow -Excel $excel -Protection $protection -Scenario 'copy.fixture'")
    assert "Set-NamedValue -Workbook $wb -DefinedName $iterationsName -Value ([double]$acceptanceIterations)" in copy
    assert "(($copyCalc -like 'OK|*') -and ($copySim -like 'OK|*') -and ($copyStates.Calculation -ceq $statusCurrent) -and ($copyStates.Simulation -ceq $statusCurrent))" in copy
    assert "-Scenario 'copy.protection-after-run'" in copy
    # nothing is ever saved, in either session
    assert ".Save(" not in code and ".SaveAs(" not in code and ".SaveCopyAs(" not in code
    # the copy is closed by the accepted shutdown: the finally still closes $wb unsaved
    shutdown = code[code.rindex("} finally {"):]
    assert "try { $wb.Close($false); $rel.WorkbookClosed = $true }" in shutdown


def _handler_body() -> str:
    return "\n".join(line for line in _src("handler", HANDLER_VBA).splitlines()
                     if not line.strip().startswith("'"))


def _handler_failpoint_name() -> str:
    match = re.search(r'^Private Const FAILPOINT_WORKBOOK_OPEN As String = "([^"]+)"$', _src("handler", HANDLER_VBA), re.M)
    assert match, "the handler declares no private failpoint constant"
    return match.group(1)


def test_56_the_row_o_session_reaches_the_real_handler_through_excels_event_switch_and_proves_the_failure_contract() -> None:
    """CONTRACT MATRIX ROW O, ADDED AT THE P10-R3 CLOSURE. A third disposable
    copy is opened with EnableEvents OFF so the handler does not run at open
    (proved: no protection, nothing recorded); the accepted seam is begun with
    the handler's own failpoint; the switch is put back in a finally; the REAL
    production handler is then run in the workbook's own document module. The
    checks: entered, the exact record, disclosed once with no prompt and no
    dialog, ScreenUpdating / Calculation / EnableEvents compared to their
    pre-run values with nothing written to restore them, fully released,
    usable, and the same handler applies protection again afterwards."""
    code = _code()
    block = code[code.index("$openPath = Join-Path $openDir 'PCCM_open_failure_copy.xlsm'"):]
    block = block[: block.index("Assert-FaProtectionApplied -Excel $excel -Protection $protection -Scenario 'rowo.reopen-applied'") + 120]
    # the copy is closed unsaved and released before the third session
    before = code[code.index("Add-FaCheck 'copy.run'"): code.index("$openPath = Join-Path")]
    assert "try { $wb.Close($false); $rel.WorkbookClosed = $true }" in before
    assert "Invoke-FaRelease -Ledger $rel -Obj $wb -Label 'Workbook(copy)'" in before and "$wb = $null" in before
    # the pre-open trigger: Excel's own switch, off across the open and the arming, back in a finally
    assert "if (-not [bool]$excel.EnableEvents) { throw" in block
    assert "$excel.EnableEvents = $false" in block
    armed = block[block.index("$excel.EnableEvents = $false"): block.index("} finally { $excel.EnableEvents = $true }")]
    assert "$wb = $workbooks.Open($openPath)" in armed and "$comAcquired = $comAcquired + 1" in armed
    assert "Import-FaFixtureWindow -Excel $excel -Workbook $wb -Manifest $manifest -ScriptDir $scriptDir" in armed
    assert "$suppressed = Get-FaProtectionState -Excel $excel" in armed
    assert "$resultBeforeHandler = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationResult'" in armed
    assert "$excel.Run('PCCM_AutomationBegin', $true, $script:OpenFailpoint)" in armed
    assert "Workbook_Open" not in armed, "the handler is run while events are off"
    # the failpoint name is the handler's own constant
    assert f"$script:OpenFailpoint = '{_handler_failpoint_name()}'" in code
    # the handler did not run at open: no protection, nothing recorded
    assert "Add-FaCheck 'rowo.open-suppressed' `\n        (($null -ne $suppressed) -and (-not $suppressed.Applied) -and ($suppressed.Protected -eq 0) -and (-not $suppressed.Structure) -and ($suppressed.Depth -eq 0) -and ($resultBeforeHandler -eq ''))" in block
    # the REAL handler, in the workbook's own document module, after the switch is back
    run = "$excel.Run(\"'\" + [string]$wb.Name + \"'!ThisWorkbook.Workbook_Open\")"
    assert block.count(run) == 2
    first_run = block.index(run)
    assert block.index("} finally { $excel.EnableEvents = $true }") < first_run
    assert "$screenBefore = [bool]$excel.ScreenUpdating" in block[: first_run]
    assert "$calcBefore = [int]$excel.Calculation" in block[: first_run]
    assert "$eventsBefore = [bool]$excel.EnableEvents" in block[: first_run]
    assert "catch { $handlerFailure = (Format-Err $_) }" in block
    # entered, exact record derived from the failpoint owner's wording, once
    appstate = _src("appstate", APPSTATE_VBA)
    assert "\"Injected structural failure after stage '\" & StageName & \"'.\"" in appstate
    assert "$expectedRecord = \"Workbook_Open: Injected structural failure after stage '\" + $script:OpenFailpoint + \"'.\"" in block
    assert "Add-FaCheck 'rowo.handler-entered' (($handlerFailure -eq '') -and ($recorded -clike 'Workbook_Open: *'))" in block
    assert "Add-FaCheck 'rowo.failure-path' ($recorded -ceq $expectedRecord)" in block
    assert "Add-FaCheck 'rowo.disclosed-once' (($recorded -ceq $expectedRecord) -and ($prompted -eq '') -and $eventsBefore)" in block
    assert "$prompted = Get-FaRunText -Excel $excel -Procedure 'PCCM_AutomationPrompt'" in block
    assert 'modAppState.RecordResult "Workbook_Open: " & detail' in _handler_body()
    # application state compared, never written back by the runner
    assert "Add-FaCheck 'rowo.screen-updating-restored' ([bool]$excel.ScreenUpdating -eq $screenBefore)" in block
    assert "Add-FaCheck 'rowo.calculation-unchanged' ([int]$excel.Calculation -eq $calcBefore)" in block
    assert "Add-FaCheck 'rowo.events-unchanged' ([bool]$excel.EnableEvents -eq $eventsBefore)" in block
    after_run = block[first_run:]
    for write in ("$excel.ScreenUpdating =", "$excel.Calculation =", "$excel.EnableEvents ="):
        assert write not in after_run, write
    # not half-protected: released entirely, no window depth
    assert "Add-FaCheck 'rowo.released-not-half-protected' `\n        ((-not $released.Applied) -and ($released.Protected -eq 0) -and (-not $released.Structure) -and ($released.Depth -eq 0))" in block
    # usable: a value write to the once-locked reference cell succeeds and reads back
    assert "$usableCell = $methodWs.Range($lockedAddress)" in block
    assert "$usableCell.Value2 = [string]$sourceRevisionRow.label" in block
    assert "Add-FaCheck 'rowo.usable' (($usableFailure -eq '') -and ($valueAfterOpenFailure -ceq [string]$sourceRevisionRow.label))" in block
    # protection attempted again by the same handler, seam begun with no failpoint, silent success
    assert block.index("$excel.Run('PCCM_AutomationBegin', $true, '')") < block.rindex(run)
    assert "Add-FaCheck 'rowo.reopen-silent' (($reopenFailure -eq '') -and ($recordedAfterReopen -eq ''))" in block
    assert "-Scenario 'rowo.reopen-applied'" in block
    assert "$excel.Run('PCCM_AutomationEnd')" in code[code.index("-Scenario 'rowo.reopen-applied'"):]
    # nothing is saved; the session is closed by the accepted shutdown
    assert ".Save(" not in code and ".SaveAs(" not in code


def test_57_the_handler_carries_one_dormant_failpoint_after_the_apply_and_reads_err_before_the_release() -> None:
    """THE P10-R3 SEAM, AND ITS REVERSAL. One modAppState.FailPointCheck, named
    by a private constant, after a successful ProtectionApply and before the
    restore - so an injected failure exercises the whole failure path from a
    fully protected workbook. It is dormant: FailPointCheck exits unless the
    accepted seam is active with exactly that stage, and nothing in production
    begins the seam. The failure path reads Err.Description before the release
    owner's On Error clears it. Taking the closure off reproduces 889b6b5."""
    handler = _src("handler", HANDLER_VBA)
    body = _handler_body()
    name = _handler_failpoint_name()
    assert name.startswith("Phase10") and " " not in name
    assert body.count("FailPointCheck") == 1
    assert "    modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN\n" in body
    apply_at = body.index("If Not modProtection.ProtectionApply(detail) Then GoTo Failed")
    check_at = body.index("modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN")
    restore_at = body.index("Application.ScreenUpdating = previousUpdating")
    assert body.index("On Error GoTo Failed") < apply_at < check_at < restore_at < body.index("Failed:")
    # the constant is private to the document module and is the only one there
    assert handler.count("Const ") == 1 and "Private Const FAILPOINT_WORKBOOK_OPEN" in handler
    assert handler.count("Sub ") == 1
    # dormant at the owner: inactive seam or empty stage exits before the comparison
    check = _src("appstate", APPSTATE_VBA)
    check = check[check.index("Public Sub FailPointCheck"): check.index("End Sub", check.index("Public Sub FailPointCheck"))]
    assert "If Not gAutomationActive Then Exit Sub" in check
    assert "If Len(gAutomationFailAfterStage) = 0 Then Exit Sub" in check
    assert "Err.Raise vbObjectError + 5001" in check
    # nothing in production begins the seam
    for path in sorted((PCCM_ROOT / "src" / "vba").glob("*.bas")) + [HANDLER_VBA]:
        code = "\n".join(line for line in path.read_text(encoding="utf-8").splitlines() if not line.strip().startswith("'"))
        if path.name != "modAppState.bas":
            assert "gAutomationActive = True" not in code and "PCCM_AutomationBegin" not in code, path.name
    appstate_code = "\n".join(line for line in _src("appstate", APPSTATE_VBA).splitlines() if not line.strip().startswith("'"))
    assert appstate_code.count("gAutomationActive = True") == 1
    assert "PCCM_AutomationBegin" not in _src("handler", HANDLER_VBA)
    # the failure path: Err read first, then the state restore, then the release, then the record, then the guarded dialog
    failed = body[body.index("Failed:"):]
    order = ("If Len(detail) = 0 Then detail = Err.Description",
             "If Not restored Then Application.ScreenUpdating = previousUpdating",
             "If Not modProtection.ProtectionRelease(releaseDetail) Then",
             'modAppState.RecordResult "Workbook_Open: " & detail',
             "If Not modAppState.gAutomationActive Then",
             "modAppState.ReportFailure")
    positions = [failed.index(step) for step in order]
    assert positions == sorted(positions), positions
    assert failed.count("Err.Description") == 1
    assert "On Error GoTo Failed" in (PCCM_ROOT / "src" / "vba" / "modProtection.bas").read_text(encoding="utf-8")[
        (PCCM_ROOT / "src" / "vba" / "modProtection.bas").read_text(encoding="utf-8").index("Public Function ProtectionRelease"):]
    # the successful path is unchanged in behaviour: apply, restore, exit, and nothing else
    success = body[body.index("On Error GoTo Failed"): body.index("Failed:")]
    statements = [line.strip() for line in success.splitlines() if line.strip()]
    assert statements == ["On Error GoTo Failed",
                          "previousUpdating = Application.ScreenUpdating",
                          "Application.ScreenUpdating = False",
                          "restored = False",
                          "If Not modProtection.ProtectionApply(detail) Then GoTo Failed",
                          "modAppState.FailPointCheck FAILPOINT_WORKBOOK_OPEN",
                          "Application.ScreenUpdating = previousUpdating",
                          "restored = True",
                          "Exit Sub"], statements
    # the reversal
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_open_failpoint import ACCEPTED_BEFORE_OPEN_FAILPOINT, strip_open_failpoint
    tested = _git("show", f"{ACCEPTED_BEFORE_OPEN_FAILPOINT}:pccm/src/vba/ThisWorkbook.vba")
    assert strip_open_failpoint("ThisWorkbook.vba", handler) == tested
    assert "FailPointCheck" not in tested


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
    for check, word in (("state.initial.persisted", "$statusNotCalculated"), ("state.initial.live", "$statusInvalid"),
                        ("calculate.current", "$statusCurrent"),
                        ("state.stale", "$statusStale"), ("state.invalid", "$statusInvalid")):
        tail = code[code.index(f"Add-FaCheck '{check}'"): code.index(f"Add-FaCheck '{check}'") + 400]
        assert word in tail, check
    assert "for literal in" not in code
    for word in ("'NOT CALCULATED'", "'CURRENT'", "'STALE'", "'INVALID'", "'REFUSED'"):
        assert word not in code, word


def test_60b_the_untouched_workbook_is_asserted_as_phase_9_accepted_it_live_and_persisted_apart() -> None:
    """RUN 1 AT 6770cb8 FAILED HERE. The runner expected the LIVE calculation
    state of the untouched workbook to read NOT CALCULATED; Phase 9 had already
    settled that the untouched workbook's required inputs are unresolved, so its
    LIVE state is INVALID, the simulation INVALID as its consequence, and annual
    and profile NOT PRODUCED, while NOT CALCULATED is a PERSISTED history fact:
    no calculation has ever been committed. The two are now asserted apart."""
    code = _code()
    live = code[code.index("Add-FaCheck 'state.initial.live'"):]
    live = live[: live.index("Save-Phase5LockedFxSeed")]
    assert "($states0.Calculation -ceq $statusInvalid)" in live
    assert "($states0.Simulation -ceq $statusInvalid)" in live
    assert "($states0.Annual -like 'NOT PRODUCED*')" in live and "($states0.Profile -like 'NOT PRODUCED*')" in live
    assert "$statusNotCalculated" not in live
    persisted = code[code.index("$persistedStatus = Format-FaCell"):]
    persisted = persisted[: persisted.index("$compileFailure = ''")]
    assert "($persistedStatus -ceq $statusNotCalculated)" in persisted
    assert "($persistedAttempt -ceq $attemptNone)" in persisted
    assert "($persistedFingerprint -ceq '<blank>')" in persisted
    assert "$statusInvalid" not in persisted
    # THE PERSISTED CELLS ARE THE MODEL'S OWN, from the Phase-5 block, read with
    # no accessor - and read BEFORE the compile check, which evaluates and
    # persists the live status.
    assert "$calcPersistedStatusCell = ($calcValueColumn + [string]$calcStateBlock.rows.calculation_status)" in code
    assert "$calcFingerprintCell = ($calcValueColumn + [string]$calcStateBlock.rows.last_successful_fingerprint)" in code
    assert "-Address $calcPersistedStatusCell" in persisted and "-Address $calcAttemptCell" in persisted \
        and "-Address $calcFingerprintCell" in persisted
    assert "$excel.Run(" not in persisted
    assert code.index("$wb = $workbooks.Open($stageBPath)") < code.index("Add-FaCheck 'state.initial.persisted'") \
        < code.index("$excel.Run('PCCM_CalculationStatus')") < code.index("Add-FaCheck 'state.initial.live'")
    # AND NO MODEL CHECK IS READ ON THE UNTOUCHED WORKBOOK: the first Model Check
    # read follows the first Calculate, so no PASS is forced on an invalid model.
    assert code.index("Add-FaCheck 'calculate.current'") < code.index("Assert-FaModelCheck -Workbook $wb")


def test_60c_every_scenario_from_the_fixture_onward_is_the_executed_runner_plus_the_declared_round_changes() -> None:
    """RESTATED at the bounded correction round after the independent review.
    The tail of the runner Windows executed at run 7 (the starting authority's
    runner; byte-identical to ee6e9fb's) is the tail of the runner now, with
    exactly three declared changes: the Repair contract scenarios INSERTED as
    one delimited block, the distribution-copy session INSERTED as one
    delimited block, and the final protection behaviour SUBSTITUTED between its
    heading and the automation end. Taking the insertions out and the
    substitution aside reproduces the executed tail byte for byte."""
    tested = _git("show", f"{STARTING_AUTHORITY}:pccm/bootstrap/windows/phase10_final_acceptance.ps1").replace("\r\n", "\n")
    assert _git("diff", "--stat", "ee6e9fb", STARTING_AUTHORITY, "--",
                "pccm/bootstrap/windows/phase10_final_acceptance.ps1").strip() == ""
    now = _runner().replace("\r\n", "\n")
    marker = "    # 7. THE ACCEPTED W4 FIXTURE"
    tested_tail = tested[tested.index(marker):]
    now_tail = _without_r2_blocks(now[now.index(marker):])
    assert now_tail != now[now.index(marker):], "the two declared insertions are absent"
    tested_head, tested_protection, tested_rest = _split_protection_region(tested_tail)
    now_head, now_protection, now_rest = _split_protection_region(now_tail)
    assert now_head == tested_head
    assert now_rest == tested_rest
    assert now_protection != tested_protection, "the declared protection substitution is absent"
    # The substitution keeps the unlocked half of the pair and replaces the
    # locked half: the old "must be refused" expectation is gone with its variables.
    assert "Add-FaCheck 'protection.unlocked-writable' ($unlockedFailure -eq '')" in now_protection
    assert "protection.locked-refused" in tested_protection and "$lockedRefused" in tested_protection
    assert "protection.locked-refused" not in now_protection and "lockedRefused" not in now_protection
    assert "$lockedDetail" not in now_protection
    # And before the fixture: the executed head plus the declared helper and
    # preamble insertions, nothing else.
    head_now = now[: now.index(marker)]
    head_tested = tested[: tested.index(marker)]
    for insertion in _runner_head_insertions():
        assert head_now.count(insertion) == 1, insertion[:60]
        head_now = head_now.replace(insertion, "")
    assert head_now == head_tested


def test_60c2_the_executed_runner_is_the_run_2_runner_plus_the_model_check_corrections() -> None:
    """THE HISTORICAL CHAIN, KEPT. The runner Windows executed at run 7 is the
    runner Windows executed at 9686baf with exactly the four Model Check
    substitutions applied to its tail, and the reader/expectation region moved
    in its head - both sides now immutable in git."""
    tested = _git("show", "9686baf:pccm/bootstrap/windows/phase10_final_acceptance.ps1").replace("\r\n", "\n")
    now = _git("show", f"{STARTING_AUTHORITY}:pccm/bootstrap/windows/phase10_final_acceptance.ps1").replace("\r\n", "\n")
    marker = "    # 7. THE ACCEPTED W4 FIXTURE"
    tail = tested[tested.index(marker):]
    for old, new in MODEL_CHECK_CORRECTIONS:
        assert tail.count(old) == 1, old[:60]
        tail = tail.replace(old, new)
    assert now[now.index(marker):] == tail
    head_marker = "$calcAttemptCell = ($calcValueColumn + [string]$calcStateBlock.rows.last_attempt_result)\n"
    preflight = "# ===========================================================================\n# PREFLIGHT"

    def outside_readers(text: str, reader_marker: str) -> str:
        head = text[: text.index(head_marker)]
        return head[: head.index(reader_marker)] + head[head.index(preflight):]

    tested_head = outside_readers(tested, "# Read the Model Check summary through the Phase-9 projection")
    now_head = outside_readers(now, "# Read the Model Check surface through the Phase-9 projection")
    expectations = now_head[now_head.index("$severityError       = "): now_head.index("# THE CALCULATION STATE BLOCK")]
    assert "$notCalculatedExpected = @{ AnyOf = $calcWarningIds" in expectations
    tested_expectations = tested_head[tested_head.index("$overallPass         = "): tested_head.index("# THE CALCULATION STATE BLOCK")]
    assert now_head.replace(expectations, "") == tested_head.replace(tested_expectations, "")


def _runner_head_insertions() -> tuple[str, ...]:
    """The exact text this round added before the fixture: two failpoint names,
    two grid-digest variables, and three helpers (a column added, the last
    column removed, one weight written) - read from the runner by their own
    delimiters, so that the reversal in test_60c is a statement about THESE
    fragments and nothing else."""
    now = _runner().replace("\r\n", "\n")
    preamble_start = now.index("# The two contracted Repair failpoints: after the cost grid and after the risk grid.\n")
    preamble_stop = now.index("$script:FaRiskBefore = ''\n") + len("$script:FaRiskBefore = ''\n")
    helpers_start = now.index("# A PROJECT-YEAR COLUMN ADDED TO OR REMOVED FROM A GRID")
    helpers_stop = now.index("function Get-IdColumnValues {")
    return (now[preamble_start:preamble_stop], now[helpers_start:helpers_stop]) + _r3_head_insertions()


def _r3_head_insertions() -> tuple[str, ...]:
    """The one line the P10-R3 closure added before the fixture: the row-O
    failpoint name, with its comment."""
    now = _runner().replace("\r\n", "\n")
    start = now.index("# The contracted Workbook_Open failpoint (matrix row O)")
    stop = now.index("$script:OpenFailpoint = ") 
    stop = now.index("\n", stop) + 1
    return (now[start:stop],)


def test_60c3_the_row_o_session_is_the_one_declared_change_since_the_closure_authority() -> None:
    """P10-R3 REVERSAL. The runner at 889b6b5 plus exactly one delimited block
    and one preamble line is the runner now, byte for byte."""
    tested = _git("show", f"{R3_AUTHORITY}:pccm/bootstrap/windows/phase10_final_acceptance.ps1").replace("\r\n", "\n")
    now = _runner().replace("\r\n", "\n")
    stripped = _without_r2_blocks(now, R3_BLOCKS)
    assert stripped != now, "the row-O block is absent"
    for insertion in _r3_head_insertions():
        assert stripped.count(insertion) == 1
        stripped = stripped.replace(insertion, "")
    assert stripped == tested


MODEL_CHECK_CORRECTIONS = (
    ("    $summaryCalc = Get-FaModelCheckSummary -Workbook $wb -Projection $projection\n"
     "    $null = Add-FaCheck 'modelcheck.calculated' ([string]$summaryCalc['overall_status'] -ceq $overallPass) `\n"
     "        ('overall=' + (Format-FaCell $summaryCalc['overall_status']) + ' errors=' + (Format-FaCell $summaryCalc['error_count']))\n",
     "    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.calculated' `\n"
     "        -Expected @($advisoryExpected)\n"),
    ("    $summaryFull = Get-FaModelCheckSummary -Workbook $wb -Projection $projection\n"
     "    $null = Add-FaCheck 'modelcheck.simulated' ([string]$summaryFull['overall_status'] -ceq $overallPass) `\n"
     "        ('overall=' + (Format-FaCell $summaryFull['overall_status']) + ' errors=' + (Format-FaCell $summaryFull['error_count']))\n",
     "    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.simulated' `\n"
     "        -Expected @($advisoryExpected)\n"),
    ("    $summaryInvalid = Get-FaModelCheckSummary -Workbook $wb -Projection $projection\n"
     "    $null = Add-FaCheck 'modelcheck.invalid' ([string]$summaryInvalid['overall_status'] -ceq $overallError) `\n"
     "        ('overall=' + (Format-FaCell $summaryInvalid['overall_status']) + ' errors=' + (Format-FaCell $summaryInvalid['error_count']))\n",
     "    $refusalSubject = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckRefusalSubject'\n"
     "    $null = Add-FaCheck 'modelcheck.invalid.subject' ($refusalSubject -ceq $victimId) `\n"
     "        ('the refusal subject is ' + $refusalSubject + '; the invalidated driver is ' + $victimId)\n"
     "    # THE ANNUAL OUTPUTS PUBLISHED BEFORE THE INVALIDATION ARE NOW HISTORICAL.\n"
     "    $null = Add-FaCheck 'state.invalid.annual-historical' `\n"
     "        (($statesInvalid.Annual -ceq $annualHistorical) -and ($statesInvalid.Profile -ceq $profileHistorical)) `\n"
     "        ('annual=' + $statesInvalid.Annual + ' profile=' + $statesInvalid.Profile)\n"
     "    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.invalid' `\n"
     "        -Expected @(@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId },\n"
     "                    $advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)\n"),
    ("    $summaryReset = Get-FaModelCheckSummary -Workbook $wb -Projection $projection\n"
     "    $modelCheckReset = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckCalculationState'\n"
     "    $null = Add-FaCheck 'modelcheck.after-reset' ($modelCheckReset -ceq $statusNotCalculated) `\n"
     "        ('adapter=' + $modelCheckReset + ' overall=' + (Format-FaCell $summaryReset['overall_status']) + ' errors=' + (Format-FaCell $summaryReset['error_count']))\n",
     "    $modelCheckReset = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckCalculationState'\n"
     "    $null = Add-FaCheck 'modelcheck.after-reset.adapter' ($modelCheckReset -ceq $statusNotCalculated) ('adapter=' + $modelCheckReset)\n"
     "    $null = Assert-FaModelCheck -Workbook $wb -Projection $projection -Scenario 'modelcheck.after-reset' `\n"
     "        -Expected @($advisoryExpected, $notCalculatedExpected)\n"),
)


def _spec_checks() -> list[dict]:
    import yaml
    return yaml.safe_load((PCCM_ROOT / "spec" / "workbook.yaml").read_text(encoding="utf-8"))["phase9_shell"]["model_check"]["checks"]


def _projection() -> dict:
    import json
    return json.loads((PCCM_ROOT / "build" / "phase9_model_check_inspection.json").read_text(encoding="utf-8"))


def test_60e_a_valid_model_at_the_business_minimum_expects_the_advisory_warning_not_pass() -> None:
    """RUN 2 AT 9686baf FAILED HERE: the runner expected PASS after Calculate. The
    Phase-9 contract places the business minimum strictly below the recommended
    iterations, so the low-iteration advisory is an actionable WARNING with zero
    errors, and it refuses nothing. The runner now expects exactly that."""
    code = _code()
    projection = _projection()
    import json
    cases = json.loads((PCCM_ROOT / "build" / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    assert int(cases["bounds"]["business_minimum_iterations"]) < int(projection["advisory"]["threshold"])
    assert projection["advisory"]["severity"] == "WARNING" and projection["advisory"]["refuses"] is False
    assert "if (-not ($acceptanceIterations -lt [int]$projection.advisory.threshold)) {" in code
    assert "$advisoryExpected = @{ Id = [string]$projection.advisory.check_id; Severity = [string]$projection.advisory.severity" in code
    assert "Message = [string]$projection.advisory.message }" in code
    for scenario in ("modelcheck.calculated", "modelcheck.simulated"):
        assert f"-Scenario '{scenario}' `\n        -Expected @($advisoryExpected)\n" in code, scenario
    assert "$overallPass" not in code and "overall_states[0]" not in code.replace("$states[0]", "")


def test_60f_the_invalid_checkpoint_expects_the_calculation_error_the_advisory_and_both_historical_annual_warnings() -> None:
    """THE ACTUAL SEQUENCE, NOT AN ABSTRACT INVALID MODEL. Run 5 at 46100b0 showed
    what the runner's own order produces: Calculate, Simulation, Sensitivity and
    Annual all succeed, THEN a driver is invalidated, so the annual outputs
    published for the earlier run are HISTORICAL and the Phase-9 contract raises
    the historical-profile and historical-distribution Annual WARNINGs beside the
    one Calculation ERROR and the advisory. The runner expected only the error and
    the advisory; it now expects the exact four. The simulation's invalidity under
    a non-CURRENT calculation is still context, never a second actionable row."""
    code = _code()
    assert "$calcErrorChecks = @($projection.evaluation.declared_checks | Where-Object {" in code
    assert "if ($calcErrorChecks.Count -ne 1) { throw" in code
    assert "$groupAnnual = [string]$projection.vocabulary.group_order[4]" in code
    assert "$annualWarningIds = @($projection.evaluation.declared_checks | Where-Object {" in code
    assert "([string]$_.group -ceq $groupAnnual) -and ([string]$_.severity -ceq $severityWarning) } |" in code
    assert "$annualHistoricalExpected = @{ AnyOf = $annualWarningIds; Severity = $severityWarning; Subject = '' }" in code
    assert "$annualHistorical  = [string]$p7.handoff.distribution_states[2]" in code
    assert "$profileHistorical = [string]$p7.handoff.profile_states[3]" in code
    invalid = code[code.index("$refusalSubject = Get-FaRunText -Excel $excel -Procedure 'PCCM_ModelCheckRefusalSubject'"):]
    invalid = invalid[: invalid.index("Set-TableCell -Workbook $wb")]
    assert "Add-FaCheck 'modelcheck.invalid.subject' ($refusalSubject -ceq $victimId)" in invalid
    assert "(($statesInvalid.Annual -ceq $annualHistorical) -and ($statesInvalid.Profile -ceq $profileHistorical))" in invalid
    assert "-Expected @(@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId }," in invalid
    assert "$advisoryExpected, $annualHistoricalExpected, $annualHistoricalExpected)" in invalid
    # NO ANNUAL CHECK ID IS A LITERAL.
    assert "ANN-" not in _runner()
    # FROM THE PROJECTION AND THE SPEC: the projection's Annual WARNING population
    # is three; the spec fires two of them at HISTORICAL profile and HISTORICAL
    # distributions, both without a subject; the third needs a profile state of
    # OTHER Px, which a HISTORICAL profile cannot be, and carries a subject.
    import json
    projection = _projection()
    assert projection["vocabulary"]["group_order"][4] == "Annual"
    annual_warnings = [c["check_id"] for c in projection["evaluation"]["declared_checks"]
                       if c["group"] == "Annual" and c["severity"] == "WARNING"]
    assert annual_warnings == ["ANN-010", "ANN-020", "ANN-050"]
    checks = {c["check_id"]: c for c in _spec_checks()}
    assert checks["ANN-010"]["condition"] == '{annual_profile_state}="HISTORICAL"' and not checks["ANN-010"].get("subject")
    assert checks["ANN-050"]["condition"] == '{annual_distribution_state}="HISTORICAL"' and not checks["ANN-050"].get("subject")
    assert checks["ANN-020"]["condition"] == '{annual_profile_state}="OTHER Px"' and checks["ANN-020"]["subject"] == "{annual_profile_px}"
    p7 = json.loads((PCCM_ROOT / "build" / "phase7_acceptance_inspection.json").read_text(encoding="utf-8"))
    assert p7["handoff"]["distribution_states"][2] == "HISTORICAL" and p7["handoff"]["profile_states"][3] == "HISTORICAL"
    # The Calculation ERROR and the simulation context, as before.
    calc_errors = [c for c in checks.values() if c["group"] == "Calculation" and c["severity"] == "ERROR"]
    assert [c["check_id"] for c in calc_errors] == ["CAL-010"]
    assert calc_errors[0]["condition"] == '{calculation_state}="INVALID"'
    assert calc_errors[0]["subject"] == "{calculation_refusal_subject}"
    assert checks["SIM-010"]["condition"] == 'AND({simulation_state}="INVALID",{calculation_state}="CURRENT")'
    assert checks["SIM-020"]["severity"] == "INFO"
    assert checks["SIM-020"]["condition"] == 'AND({simulation_state}="INVALID",{calculation_state}<>"CURRENT")'


def test_60g_the_after_reset_checkpoint_expects_the_not_calculated_warning_and_the_advisory() -> None:
    """FROM THE SPEC: two Calculation WARNING checks exist, one for NOT
    CALCULATED and one for STALE; the calculation state is one word, so at
    NOT CALCULATED - which the runner asserts separately at that point - the one
    Calculation WARNING shown is the NOT CALCULATED check, with no subject; and
    the advisory still fires because the request is still the business minimum."""
    code = _code()
    assert "$calcWarningIds = @($projection.evaluation.declared_checks | Where-Object {" in code
    assert "$notCalculatedExpected = @{ AnyOf = $calcWarningIds; Severity = $severityWarning; Subject = '' }" in code
    after = code[code.index("Add-FaCheck 'modelcheck.after-reset.adapter'"):]
    after = after[: after.index("Assert-FaProtectionApplied")]
    assert "($modelCheckReset -ceq $statusNotCalculated)" in after
    assert "-Scenario 'modelcheck.after-reset' `\n        -Expected @($advisoryExpected, $notCalculatedExpected)" in after
    # The request is still the business minimum there: the only two writes of
    # the iterations name are the minimum and the minimum plus one, and the last
    # write of the main session, which precedes the reset block, restores the
    # minimum. The distribution-copy session (bounded correction round) writes
    # the minimum once more, into the fresh copy.
    pattern = r"Set-NamedValue -Workbook \$wb -DefinedName \$iterationsName -Value \(\[double\](\$\w+)\)"
    main = code[: code.index("Copy-Item -LiteralPath $stageBPath")]
    writes = [m.group(1) for m in re.finditer(pattern, main)]
    assert writes[-1] == "$acceptanceIterations"
    assert main.rindex("Set-NamedValue -Workbook $wb -DefinedName $iterationsName") < main.index("Add-FaCheck 'reset.precondition'")
    copy_writes = [m.group(1) for m in re.finditer(pattern, code[len(main):])]
    assert copy_writes == ["$acceptanceIterations"], copy_writes
    checks = _spec_checks()
    calc_warnings = {c["check_id"]: c["condition"] for c in checks if c["group"] == "Calculation" and c["severity"] == "WARNING"}
    assert calc_warnings == {"CAL-020": '{calculation_state}="NOT CALCULATED"', "CAL-030": '{calculation_state}="STALE"'}
    assert all(not c.get("subject") for c in checks if c["check_id"] in calc_warnings)


def test_60h_the_model_check_assertion_derives_everything_from_the_expected_set_and_refuses_unrelated_rows() -> None:
    fn = _function("Assert-FaModelCheck", _code())
    assert "$states = @($Projection.vocabulary.overall_states | ForEach-Object { [string]$_ })" in fn
    assert "$actionable = @($Projection.vocabulary.actionable_severities | ForEach-Object { [string]$_ })" in fn
    assert "if ($expectedWarnings -gt 0) { $expectedOverall = $states[1] }" in fn
    assert "if ($expectedErrors -gt 0) { $expectedOverall = $states[2] }" in fn
    # A MATCHED ROW LEAVES THE POOL, so two identical AnyOf entries consume two
    # DISTINCT rows and the same row can never satisfy both.
    assert "$unmatched = @($unmatched | Where-Object { -not [object]::ReferenceEquals($_, $found) })" in fn
    for required in ("if ($overall -cne $expectedOverall)", "if ($errors -ne $expectedErrors)", "if ($warnings -ne $expectedWarnings)",
                     "if ($unmatched.Count -gt 0) { $problems += ('unexpected actionable row(s): '",
                     "is not shown as expected", "($entry.HasMessage -and ((Format-FaCell $row.message) -cne $entry.Message))",
                     "($entry.HasSubject -and ((Format-FaCell $row.subject) -cne $entry.Subject))"):
        assert required in fn, required
    reader = _function("Get-FaModelCheckSurface", _code())
    assert "$Projection.register.first_row" in reader and "$Projection.register.last_row" in reader
    assert "$script:FaErrorCodes.ContainsKey([int]$id)" in reader
    assert "Get-FaModelCheckSummary" not in _code()


def test_60i_the_accepted_phase_9_formulas_produce_exactly_the_runner_expectations() -> None:
    """SOURCE-PROVEN BY THE ACCEPTED STATIC EVALUATOR, not by reading. The
    Phase-9 suite evaluates the sheet's own formulas over the readings each
    checkpoint will present at the business minimum."""
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    import test_phase9_model_check as p9
    import json
    plan = p9._plan()
    minimum = int(json.loads((PCCM_ROOT / "build" / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))["bounds"]["business_minimum_iterations"])

    def actionable(readings: dict) -> tuple[str, list[tuple[str, str, str]]]:
        result = p9._evaluate(plan, dict({"requested_iterations": minimum, "structural_report": ""}, **readings))
        rows = [(str(r["check_id"]), str(r["severity"]), str(r["subject"])) for r in result["shown"]
                if str(r["severity"]) in ("ERROR", "WARNING")]
        return str(result["summary"]["overall_status"]), sorted(rows)

    advisory = ("INP-010", "WARNING", "Monte Carlo Iterations")
    assert actionable({"calculation_state": "CURRENT", "simulation_state": "", "calculation_attempt_result": "SUCCESS",
                       "calculation_attempt_detail": "Calculation committed."}) == ("WARNING", [advisory])
    assert actionable({"calculation_state": "CURRENT", "simulation_state": "CURRENT", "annual_distribution_state": "CURRENT",
                       "annual_profile_state": "CURRENT", "annual_profile_px": "P80", "annual_year_count": 4,
                       "calculation_attempt_result": "SUCCESS", "calculation_attempt_detail": "Calculation committed.",
                       "simulation_publication": "stamp", "published_run_id": "run", "published_iterations_run": minimum,
                       "simulation_status_last_evaluated": "CURRENT", "sensitivity_availability": "Available"}) == ("WARNING", [advisory])
    # THE ACTUAL SEQUENCE: annual outputs were published before the driver was
    # invalidated, so annual and profile are HISTORICAL (run 5 observed exactly
    # this), and the two historical Annual WARNINGs fire beside the error and the
    # advisory - four actionable rows, ERROR 1/3.
    invalid_after_annual = {"calculation_state": "INVALID", "simulation_state": "INVALID",
                            "annual_distribution_state": "HISTORICAL", "annual_profile_state": "HISTORICAL",
                            "annual_profile_px": "P50", "annual_year_count": 4,
                            "calculation_refusal_detail": "maximum below minimum", "calculation_refusal_subject": "CL-001",
                            "calculation_attempt_result": "REFUSED", "calculation_attempt_detail": "refused",
                            "simulation_publication": "stamp", "published_run_id": "run", "published_iterations_run": minimum,
                            "simulation_status_last_evaluated": "CURRENT", "sensitivity_availability": "Available"}
    assert actionable(invalid_after_annual) == ("ERROR", sorted([("CAL-010", "ERROR", "CL-001"), advisory,
                                                                  ("ANN-010", "WARNING", ""), ("ANN-050", "WARNING", "")]))
    result = p9._evaluate(plan, dict({"requested_iterations": minimum, "structural_report": ""}, **invalid_after_annual))
    assert (int(result["summary"]["error_count"]), int(result["summary"]["warning_count"])) == (1, 3)
    # And an abstract INVALID model with CURRENT annual outputs is a different
    # state the runner never reaches: it shows no Annual WARNING at all.
    assert actionable(dict(invalid_after_annual, annual_distribution_state="CURRENT", annual_profile_state="CURRENT")) \
        == ("ERROR", sorted([("CAL-010", "ERROR", "CL-001"), advisory]))
    assert actionable({"calculation_state": "NOT CALCULATED", "simulation_state": "", "calculation_attempt_result": "NONE"}) \
        == ("WARNING", sorted([("CAL-020", "WARNING", ""), advisory]))
    # And the same valid model at the recommendation is PASS: the WARNING the
    # runner expects is the request size, not the model.
    threshold = int(_projection()["advisory"]["threshold"])
    result = p9._evaluate(plan, {"requested_iterations": threshold, "calculation_state": "CURRENT", "simulation_state": ""})
    assert str(result["summary"]["overall_status"]) == "PASS"


def test_60k_every_projection_path_the_runner_reads_exists_in_the_generated_phase_9_projection() -> None:
    """THE SCHEMA-PATH CONTROL. Run 3 at 2bc10e8 died in preflight with
    PropertyNotFoundStrict on `declared_checks`: the runner read it at the
    projection root while build_phase9_inspection() emits it under
    `evaluation`, and the controls had pinned the same wrong path. This walks
    EVERY `$projection.<path>` the runner spells against the generated JSON, so
    the projection builder - never a copy of its schema - is the authority."""
    import json
    projection = json.loads((PCCM_ROOT / "build" / "phase9_model_check_inspection.json").read_text(encoding="utf-8"))
    paths = sorted({m.group(1) for m in re.finditer(r"\$[Pp]rojection\.([A-Za-z_][A-Za-z0-9_.]*)", _code())})
    assert paths, "the runner reads nothing from the projection"
    for path in paths:
        node = projection
        for segment in path.split("."):
            if segment == "PSObject":
                break
            assert isinstance(node, dict) and segment in node, (path, segment, sorted(node) if isinstance(node, dict) else type(node).__name__)
            node = node[segment]
    # THE TWO DISCOVERIES CONSUME THE AUTHORITATIVE PATH, AND NOTHING AT THE ROOT.
    assert "$projection.evaluation.declared_checks" in _code()
    assert "$projection.declared_checks" not in _code()
    assert "evaluation.declared_checks" in paths
    assert isinstance(projection["evaluation"]["declared_checks"], list) and projection["evaluation"]["declared_checks"]
    for required in ("vocabulary.overall_states", "vocabulary.actionable_severities", "vocabulary.group_order",
                     "advisory.check_id", "advisory.severity", "advisory.message", "advisory.threshold",
                     "summary.value_column", "summary.rows", "register.columns", "register.first_row",
                     "register.last_row", "sheet"):
        assert any(path == required or path.startswith(required + ".") for path in paths), required
    # And the builder is the authority the runner follows: its own validator
    # reads the same path.
    builder = (PCCM_ROOT / "builder" / "pccm_builder" / "phase9_model_check.py").read_text(encoding="utf-8")
    assert 'inspection["evaluation"]["declared_checks"]' in builder


MATCHER_HARNESS = PCCM_ROOT / "tests" / "phase10_final_acceptance_matcher_flow.ps1"
PROJECTION_JSON = PCCM_ROOT / "build" / "phase9_model_check_inspection.json"


def _matcher_lines(runner: Path = RUNNER) -> dict[str, str]:
    done = subprocess.run([PWSH, "-NoProfile", "-File", str(MATCHER_HARNESS), "-Runner", str(runner),
                           "-Projection", str(PROJECTION_JSON)], capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout + done.stderr
    lines = {}
    for line in done.stdout.splitlines():
        parts = line.split("|", 2)
        if len(parts) == 3 and parts[0] in ("MATCH", "REFUSED", "ERROR"):
            lines[parts[1]] = line
    assert lines, done.stdout
    return lines


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_60m_the_matcher_executes_every_real_expectation_shape_under_strict_mode() -> None:
    """RUN 4 AT 95e322f DIED HERE: `$entry.Subject` on the advisory expectation,
    which names no Subject, under Set-StrictMode. The real matcher is lifted
    out of the runner by AST and driven over the three real shapes, three
    refusals and four malformed definitions, with no Excel and no COM."""
    lines = _matcher_lines()
    for case in ("A.advisory", "B.invalid", "C.after-reset"):
        assert lines[case].startswith(f"MATCH|{case}|ok=True|"), lines[case]
    assert "actionable=INP-010(WARNING)[Monte Carlo Iterations]" in lines["A.advisory"]
    assert "overall=ERROR errors=1 warnings=1 actionable=CAL-010(ERROR)[CL-001], INP-010(WARNING)" in lines["B.invalid"]
    assert "overall=WARNING errors=0 warnings=2" in lines["C.after-reset"] and "CAL-020(WARNING)[]" in lines["C.after-reset"]
    for case, why in (("D.unrelated", "unexpected actionable row(s): ANN-050(WARNING)"),
                      ("E.missing", "expected INP-010(WARNING) is not shown as expected"),
                      ("F.wrong-subject", "expected CAL-010(ERROR) is not shown as expected")):
        assert lines[case].startswith(f"MATCH|{case}|ok=False|") and why in lines[case], lines[case]
    for case, why in (("M1.both-selectors", "names both Id and AnyOf"), ("M2.no-selector", "names neither Id nor AnyOf"),
                      ("M3.no-severity", "has no Severity"), ("M4.empty-anyof", "has an empty AnyOf")):
        assert lines[case].startswith(f"REFUSED|{case}|RUNNER DEFINITION ERROR") and why in lines[case], lines[case]
    assert not [line for line in lines.values() if line.startswith("ERROR|")], lines


def test_60n_no_optional_expected_field_is_read_before_its_presence_is_established() -> None:
    code = _code()
    validator = _function("Test-FaExpectedEntry", code)
    for key in ("Id", "AnyOf", "Severity", "Subject", "Message"):
        assert f"$Entry.ContainsKey('{key}')" in validator, key
    for read in ("[string]$Entry['Id']", "$Entry['AnyOf']", "[string]$Entry['Severity']", "[string]$Entry['Subject']", "[string]$Entry['Message']"):
        assert read in validator, read
    assert "if ($hasSubject) { $subject = [string]$Entry['Subject'] }" in validator
    assert "if ($hasMessage) { $message = [string]$Entry['Message'] }" in validator
    assert "if ($hasId) { $ids = @([string]$Entry['Id']) }" in validator
    assert "if ($hasAnyOf) { $wanted = 'one of ' + ($ids -join '/') }" in validator
    for refusal in ("names both Id and AnyOf", "names neither Id nor AnyOf", "has no Severity", "has a blank Severity",
                    "has an empty AnyOf", "names a blank check id", "is not a hashtable"):
        assert refusal in validator, refusal
    assert validator.count("throw ($where") == 7
    matcher = _function("Assert-FaModelCheck", code)
    assert "Test-FaExpectedEntry -Entry $raw -Scenario $Scenario" in matcher
    # NO RAW HASHTABLE READ, no selector read, and Subject and Message read only
    # on the line that first tests the descriptor's presence flag.
    assert not re.search(r"\$entry\.Id\b", matcher) and "$entry.AnyOf" not in matcher
    for banned in ("$raw.", "$entry['", "$Expected."):
        assert banned not in matcher, banned
    for field in ("Subject", "Message"):
        for line in matcher.splitlines():
            if f"$entry.{field}" in line:
                assert f"$entry.Has{field} -and" in line, line
    for guarded in ("if ($entry.HasSubject -and ((Format-FaCell $row.subject) -cne $entry.Subject))",
                    "if ($entry.HasMessage -and ((Format-FaCell $row.message) -cne $entry.Message))",
                    "if (-not ($entry.Ids -ccontains [string]$row.check_id))", "$entry.Wanted"):
        assert guarded in matcher, guarded
    # THE THREE REAL SHAPES ARE STILL THE SHAPES THE RUNNER BUILDS.
    assert "$advisoryExpected = @{ Id = [string]$projection.advisory.check_id; Severity = [string]$projection.advisory.severity" in code
    assert "$notCalculatedExpected = @{ AnyOf = $calcWarningIds; Severity = $severityWarning; Subject = '' }" in code
    assert "@{ Id = [string]$calcErrorChecks[0].check_id; Severity = $severityError; Subject = $victimId }" in code


def test_60p_omitting_either_historical_annual_warning_from_the_invalid_expectation_is_refused() -> None:
    """THE REGRESSION CONTROL: with annual and profile HISTORICAL the exact set
    has two Annual WARNING entries, and the matcher's set equality refuses a
    runner that expects one or none. Proved on the runner's own text and on the
    executed matcher: the two-entry expectation matches run 5's rows; a
    one-entry expectation leaves an unexpected row."""
    code = _code()
    invalid = code[code.index("-Scenario 'modelcheck.invalid' `"):]
    invalid = invalid[: invalid.index("Set-TableCell -Workbook $wb")]
    assert invalid.count("$annualHistoricalExpected") == 2
    if not Path(PWSH).exists():
        return
    lines = _matcher_lines()
    assert lines["G.invalid-after-annual"].startswith("MATCH|G.invalid-after-annual|ok=True|"), lines["G.invalid-after-annual"]
    assert "overall=ERROR errors=1 warnings=3" in lines["G.invalid-after-annual"]
    assert lines["H.one-annual-omitted"].startswith("MATCH|H.one-annual-omitted|ok=False|"), lines["H.one-annual-omitted"]
    assert "warnings 3, expected 2" in lines["H.one-annual-omitted"] and "unexpected actionable row(s): ANN-0" in lines["H.one-annual-omitted"]


def test_60q_the_record_states_run_5_as_a_runner_expectation_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 5 — 46100b0 — FAILED AT modelcheck.invalid — RUNNER EXPECTATION DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("351 passed", "expected 46100b0 (clean)", "observed 46100b0 (clean)", "modelcheck.calculated", "modelcheck.simulated",
                 "5 ranked of 5", "4 project years", "P50", "STALE", "calc INVALID", "sim INVALID", "annual HISTORICAL", "profile HISTORICAL",
                 "REFUSED", "CL-001", "1 ERROR and 3 WARNINGS", "ANN-010", "ANN-050", "The runner expected only one warning",
                 "runner expectation defect, NOT a production defect", "FINAL ACCEPTANCE IS NOT PASSED",
                 "Workbook.Close True", "Application.Quit True", "natural PID exit True"):
        assert fact in plain, fact


def test_60r_every_model_wide_blank_subject_expectation_is_the_worksheets_empty_string() -> None:
    """RUN 6 AT 2be9761 FAILED HERE. The runner expected the two historical
    Annual subjects as the '<blank>' token, which Format-FaCell reserves for a
    $null cell read; the Phase-9 builder writes a model-wide subject as `=""`,
    an empty STRING, so Excel answered '' and both correct rows were rejected and
    then reported as unexpected. Every expectation whose contract means "no
    subject" now carries the worksheet's representation, and the constraint
    stays so the confidence-level Annual WARNING cannot pass as historical."""
    code = _code()
    builder = (PCCM_ROOT / "builder" / "pccm_builder" / "phase9_model_check.py").read_text(encoding="utf-8")
    assert "return '=\"\"'" in builder and "`=\"\"` is an empty STRING" in builder
    blank_subject = re.findall(r"^\$(\w+) = @\{ AnyOf = \$\w+; Severity = \$severityWarning; Subject = '' \}$", code, re.M)
    assert sorted(blank_subject) == ["annualHistoricalExpected", "notCalculatedExpected"], blank_subject
    assert "Subject = '<blank>'" not in code
    # Every expected entry the runner builds, and its subject stance.
    entries = re.findall(r"@\{ (?:Id|AnyOf) = [^}]*\}", code)
    assert len(entries) == 4, entries
    assert sum(1 for e in entries if "Subject = ''" in e) == 2
    assert sum(1 for e in entries if "Subject = $victimId" in e) == 1
    assert sum(1 for e in entries if "Subject" not in e) == 1  # the advisory, matched on its message
    # Format-FaCell is unchanged: '<blank>' still names a $null read, for the
    # persisted-cell evidence that relies on the distinction.
    cell = _function("Format-FaCell", code)
    assert "if ($null -eq $Value) { return '<blank>' }" in cell and "if ($Value -is [string]) { return $Value }" in cell
    assert "($persistedFingerprint -ceq '<blank>')" in code


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_60s_the_executed_matcher_models_excels_empty_string_subjects() -> None:
    """THE HARNESS CARRIES '' FOR MODEL-WIDE ROWS, NEVER $null, and proves: the
    '<blank>' token fails against real rows exactly as run 6 did; a Px-subject
    Annual WARNING cannot satisfy the historical expectation; two identical
    AnyOf entries consume two distinct rows."""
    harness = (PCCM_ROOT / "tests" / "phase10_final_acceptance_matcher_flow.ps1").read_text(encoding="utf-8")
    assert "-Subject $null" not in harness
    assert "Subject = '<blank>' }" in harness  # only the deliberate run-6 case I
    assert harness.count("Subject = '<blank>'") == 1
    lines = _matcher_lines()
    blank = lines["I.blank-token"]
    assert blank.startswith("MATCH|I.blank-token|ok=False|"), blank
    assert blank.count("is not shown as expected") == 2 and "unexpected actionable row(s): ANN-010(WARNING)[], ANN-050(WARNING)[]" in blank
    px = lines["J.px-subject"]
    assert px.startswith("MATCH|J.px-subject|ok=False|") and "unexpected actionable row(s): ANN-020(WARNING)[P50]" in px, px
    assert px.count("is not shown as expected") == 1
    distinct = lines["K.distinct-rows"]
    assert distinct.startswith("MATCH|K.distinct-rows|ok=True|") and "ANN-010(WARNING)[], ANN-050(WARNING)[]" in distinct, distinct
    assert lines["G.invalid-after-annual"].startswith("MATCH|G.invalid-after-annual|ok=True|")


def test_60t_the_record_states_run_6_as_a_runner_representation_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 6 — 2be9761 — FAILED AT modelcheck.invalid — RUNNER REPRESENTATION DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("351 passed", "expected 2be9761 (clean)", "observed 2be9761 (clean)", "state.invalid.annual-historical",
                 "modelcheck.calculated", "modelcheck.simulated", "calc INVALID", "sim INVALID", "annual HISTORICAL", "profile HISTORICAL",
                 "CAL-010(ERROR)[CL-001]", "INP-010(WARNING)[Monte Carlo Iterations]", "ANN-010(WARNING)[]", "ANN-050(WARNING)[]",
                 "'<blank>'", "empty string", "runner representation defect, NOT a production defect",
                 "FINAL ACCEPTANCE IS NOT PASSED", "Workbook.Close True", "Application.Quit True", "natural PID exit True",
                 "You cannot call a method on a null-valued expression", "non-repeatable"):
        assert fact in plain, fact


def test_60u_the_record_states_run_7_as_a_genuine_production_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 7 — ee6e9fb — FAILED AT repair.missing-row — GENUINE PRODUCTION DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("351 passed", "expected ee6e9fb (clean)", "observed ee6e9fb (clean)", "repair.missing-row.window-after",
                 "modelcheck.invalid passed exactly", "CAL-010 ERROR [CL-001]", "INP-010 WARNING", "ANN-010 WARNING", "ANN-050 WARNING",
                 "identical CURRENT fingerprint", "Repair Profiling as a no-op", "CL-002 column 3 is 0, not blank",
                 "GENUINE PRODUCTION DEFECT", "not weakened", "cell.Value = PROFILE_INITIAL_VALUE", "extend with blanks",
                 "FINAL ACCEPTANCE IS NOT PASSED", "Workbook.Close True", "Application.Quit True", "natural PID exit True",
                 "modProfiling.SetValueFor", "OPEN — reported, not changed"):
        assert fact in plain, fact


def test_60v_the_record_states_the_bounded_correction_round_without_windows() -> None:
    """THE RECORD: four blockers corrected at source, contract row O reported
    not invented, the reviewer's Stage-A artefacts explained as stale, and no
    Windows, benchmark or full sweep claimed."""
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Bounded correction round after the independent review — from 347f42e — NO WINDOWS EXECUTED")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("Nothing was executed on Windows in this round", "No Excel was executed", "No benchmark was run",
                 "No complete full static sweep was run", "byte-identical between ee6e9fb and 347f42e",
                 "An explicit numeric 0 is populated", "REFUSES BEFORE ANY MUTATION", "A typed zero counts as populated",
                 "modProfiling.CountDataBeyond and modWorkbook.IsDataCell are byte-identical to ee6e9fb",
                 "item 5 as OPEN", "now CLOSED — corrected", "item 8 was OPEN",
                 "EVERY retained valid driver row", "independent of width drift", "[1, -1] or [0, 0] — is NOT blank",
                 'THERE IS NO "EMPTY TOTAL"', "No normalisation is performed",
                 "reproduces ee6e9fb byte for byte", "still reproduces 58b2394",
                 "UserInterfaceOnly:=True, which the Phase-10 protection probe proved PERMITS",
                 "protection.locked-cell.user-protected", "protection.locked-cell.code-write-permitted",
                 "never evidence of protection", "protection.locked-refused no longer exists",
                 "repair.width-growth", "repair.shrink-blank", "repair.shrink-zero-refused", "repair.shrink-nonzero-refused",
                 "repair.semantic-non1-refused", "repair.blank-profile-allowed", "repair.signed-zero-total-refused",
                 "repair.rollback.Phase10RepairCost", "repair.rollback.Phase10RepairRisk", "repair.grids-restored",
                 "Contract row J", "PCCM distribution copy.xlsm", "Nothing is saved in either session",
                 "Contract row O", "is NOT implemented and is reported instead", "It stays open",
                 "build/ is git-ignored", "identify ee6e9fb", "STALE with respect to 347f42e",
                 "do not represent 347f42e", "rebuild on the Windows host is deferred"):
        assert fact in plain, fact
    # THE DEFERRAL WAS TRUE OF THE SOURCE AT THAT ROUND (889b6b5 carried no
    # failpoint in the handler) and has since been CLOSED by the P10-R3 seam,
    # declared and reversed in test_57; the record keeps the history.
    handler_then = _git("show", f"{R3_AUTHORITY}:pccm/src/vba/ThisWorkbook.vba")
    assert "FailPointCheck" not in handler_then and "modProtection.ProtectionApply(detail)" in handler_then
    assert "Workbook_Open" not in _code()[: _code().index("Copy-Item -LiteralPath $stageBPath -Destination $openPath")]


def test_60o_the_record_states_run_4_as_a_runner_matcher_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 4 — 95e322f — TERMINATED IN THE MODEL CHECK MATCHER — RUNNER MATCHER DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("351 passed", "Stage-B bootstrap completed", "expected 95e322f (clean)", "observed 95e322f (clean)",
                 "PASS|calculate.current", "23DA06D35152CFF9", "No acceptance check had failed",
                 "PropertyNotFoundException", "'Subject'", "runner matcher defect, NOT a production or Model Check defect",
                 "modelcheck.calculated itself was NOT evaluated", "FINAL ACCEPTANCE IS NOT PASSED",
                 "Workbook.Close True", "Application.Quit True", "natural PID exit True"):
        assert fact in plain, fact


def test_60l_the_record_states_the_preflight_attempt_at_2bc10e8() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance attempt 3 — 2bc10e8 — TERMINATED IN PREFLIGHT — RUNNER SCHEMA-PATH DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("PropertyNotFoundStrict", "declared_checks", "No Stage-B bootstrap output", "no Excel acceptance scenario",
                 "projection.evaluation.declared_checks", "static controls had pinned the same incorrect path",
                 "Production was not implicated", "FINAL ACCEPTANCE IS NOT PASSED", "351 passed"):
        assert fact in plain, fact


def test_60j_the_record_states_run_2_as_a_runner_expectation_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 2 — 9686baf — FAILED AT modelcheck.calculated — RUNNER EXPECTATION DEFECT")
    plain = " ".join(record[start:].replace("`", "").replace("**", "").split())
    for fact in ("PASS|calculate.current", "expected 9686baf (clean)", "observed 9686baf (clean)", "23DA06D35152CFF9",
                 "FAIL|modelcheck.calculated|overall=WARNING errors=0", "the runner expected PASS",
                 "runner expectation defect, NOT a production defect", "1,000", "Phase-9",
                 "Workbook.Close True", "Application.Quit True", "natural PID exit True", "FINAL ACCEPTANCE IS NOT PASSED"):
        assert fact in plain, fact


def test_60d_the_record_states_run_1_as_a_runner_expectation_defect() -> None:
    record = (PCCM_ROOT / "docs" / "phase10_windows_run_evidence.md").read_text(encoding="utf-8")
    start = record.index("## Final acceptance run 1 — 6770cb8 — FAILED AT state.initial — RUNNER EXPECTATION DEFECT")
    section = record[start:]
    plain = " ".join(section.replace("`", "").replace("**", "").split())
    for fact in ("PASS|protection.initial", "expected 6770cb8 (clean)", "observed 6770cb8 (clean)",
                 "FAIL|state.initial|calc=INVALID sim=INVALID annual=NOT PRODUCED profile=NOT PRODUCED",
                 "runner expectation defect, NOT a production defect", "Phase 9",
                 "Workbook.Close True", "Application.Quit True", "natural PID exit True",
                 "FINAL ACCEPTANCE IS NOT PASSED"):
        assert fact in plain, fact


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
                             ("Add-FaCheck 'modelcheck.after-reset.adapter'", "protection.after-reset"),
                             ("Add-FaCheck 'refused.outcome.annual'", "protection.after-refused-annual"),
                             ("Add-FaCheck 'reset.rollback'", "protection.after-rollback")):
        start = code.index(after)
        assert f"-Scenario '{assertion}'" in code[start: start + 1500], assertion


def test_63_the_protection_behaviour_pair_separates_user_edit_protection_from_code_write_capability() -> None:
    """RESTATED at the bounded correction round after the independent review.
    The accepted design protects with UserInterfaceOnly:=True, which the
    protection probe proved PERMITS a code-driven value write to a locked cell.
    So the runner may not expect that write to fail. It establishes the target
    sheet and cell explicitly, proves the cell LOCKED on a sheet with
    ProtectContents (the user-edit protection), then proves the code write
    SUCCEEDS and reads back (the code-write capability); any exception on either
    path is a FAIL of that check and never counts as a refusal."""
    code = _code()
    assert "foreach ($candidate in @($setupProtection.unlocked))" in code
    assert "Add-FaCheck 'protection.unlocked-writable' ($unlockedFailure -eq '')" in code
    region = code[code.index("Add-FaCheck 'protection.unlocked-writable'"):]
    region = region[: region.index("-Scenario 'protection.after-locked-cell'")]
    # the target, established from the projection before any COM call
    assert "if (-not [bool]$methodologyProtection.protect) { throw" in region
    assert "if (@($methodologyProtection.unlocked).Count -ne 0) { throw" in region
    assert "$lockedAddress = [string]$methodology.label_column + [string]$sourceRevisionRow.row" in region
    # the user-edit protection: Locked on a protected sheet, read failure = FAIL
    assert "$sheetIsProtected = [bool]$methodWs.ProtectContents" in region
    assert "$lockedIsLocked = [bool]$lockedCell.Locked" in region
    assert "} catch { $lockedReadFailure = (Format-Err $_) }" in region
    assert "Add-FaCheck 'protection.locked-cell.user-protected' (($lockedReadFailure -eq '') -and $lockedIsLocked -and $sheetIsProtected)" in region
    # the code-write capability: the write succeeds and reads back, exception = FAIL
    assert "$lockedCell.Value2 = [string]$sourceRevisionRow.label" in region
    assert "$valueAfterWrite = [string]$lockedCell.Value2" in region
    assert "} catch { $codeWriteFailure = (Format-Err $_) }" in region
    assert "Add-FaCheck 'protection.locked-cell.code-write-permitted' (($codeWriteFailure -eq '') -and ($valueAfterWrite -ceq [string]$sourceRevisionRow.label))" in region
    assert region.index("'protection.locked-cell.user-protected'") < region.index("$lockedCell.Value2 = ")
    # no exception is ever read as protection, and the old expectation is gone
    assert "protection.locked-refused" not in code and "lockedRefused" not in code
    assert not re.search(r"catch \{[^}]*=\s*\$true", region), "an exception sets a success flag"
    assert "UserInterfaceOnly:=True" in (PCCM_ROOT / "src" / "vba" / "modProtection.bas").read_text(encoding="utf-8")
    assert "-Scenario 'protection.after-locked-cell'" in code


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
def test_70_production_vba_spec_and_builder_are_byte_identical_to_the_accepted_head_except_the_declared_repair_correction() -> None:
    """RESTATED at final acceptance run 7, which found a GENUINE production
    defect: Repair Profiling recreated a missing row seeded at zero where the
    contract says blank. The correction is one declared layer in modRepair.bas;
    taking it off reproduces the tree run 7 executed byte for byte, and no other
    production, spec or builder byte has moved since the accepted head."""
    changed = _git("diff", "--name-only", ACCEPTED, "--", "pccm/src", "pccm/spec", "pccm/builder").split()
    assert sorted(changed) == ["pccm/src/vba/ThisWorkbook.vba", "pccm/src/vba/modRepair.bas"], changed
    sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_open_failpoint import ACCEPTED_BEFORE_OPEN_FAILPOINT, strip_open_failpoint
    from vba_repair_reconstruction import (ACCEPTED_BEFORE_REPAIR_RECONSTRUCTION,
                                           strip_repair_reconstruction)
    current = (PCCM_ROOT / "src" / "vba" / "modRepair.bas").read_bytes().decode("utf-8")
    tested = _git("show", f"{ACCEPTED_BEFORE_REPAIR_RECONSTRUCTION}:pccm/src/vba/modRepair.bas")
    assert strip_repair_reconstruction("modRepair.bas", current) == tested
    assert current != tested, "the declared correction is absent"
    assert _git("diff", "--name-only", ACCEPTED, ACCEPTED_BEFORE_REPAIR_RECONSTRUCTION, "--",
                "pccm/src", "pccm/spec", "pccm/builder").strip() == ""
    # AND THE P10-R3 CLOSURE on top: ThisWorkbook.vba only, reversing to 889b6b5,
    # which is byte-identical to the accepted head for that file.
    assert ACCEPTED_BEFORE_OPEN_FAILPOINT == R3_AUTHORITY
    handler = _src("handler", HANDLER_VBA)
    handler_then = _git("show", f"{ACCEPTED_BEFORE_OPEN_FAILPOINT}:pccm/src/vba/ThisWorkbook.vba")
    assert strip_open_failpoint("ThisWorkbook.vba", handler) == handler_then
    assert handler != handler_then, "the declared closure is absent"
    assert _git("diff", "--name-only", ACCEPTED, ACCEPTED_BEFORE_OPEN_FAILPOINT, "--",
                "pccm/src/vba/ThisWorkbook.vba").strip() == ""
    assert _git("diff", "--name-only", ACCEPTED_BEFORE_OPEN_FAILPOINT, "--",
                "pccm/src", "pccm/spec", "pccm/builder").split() == ["pccm/src/vba/ThisWorkbook.vba"]


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
