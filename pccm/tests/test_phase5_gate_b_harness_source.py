#!/usr/bin/env python3
"""PCCM Phase 5 Gate B Step B1: STATIC tests over the Windows harness extension.

NO WINDOWS RUN HAPPENS HERE, AND NONE CAN. Linux has no Excel and no PowerShell
host; every assertion below is a statement about SOURCE TEXT and about the build
artifacts the Stage-A builder emits. Nothing here starts a COM session, drives a
workbook, or claims that any VBA has executed.

What this file DOES establish:

  * the accepted Phase-4 harness remains the base, with its 35 scenario IDs
    intact and its 35/35 result a Gate-B prerequisite
  * every plan-case ID the corpus emits is mapped to a Windows scenario, and the
    preflight that proves it runs before Excel is started
  * expected analytical values are LOADED from build/phase5_cases.json and are
    not hand-maintained anywhere in the harness
  * the transient diagnostic module is absent from the production manifest,
    declares no PCCM_ endpoint, is imported only after the A1 production compile
    proof, and is removed again
  * the locked direct-vector sets are complete, both decimal separators are
    injected, and the reference stream is asserted by BOTH unit count and digest
  * the rollback comparison uses C13:C16, C23:C32 and the five tables, and does
    NOT demand that all of C13:C20 stay unchanged
  * no Calculate button is introduced and no production module is added

What it does NOT establish is anything about behaviour. Whether Excel agrees is
Gate B's, on Windows, and that run has not been made.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import re
import sys
import tempfile
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
HARNESS = BOOTSTRAP / "phase4_functional_test.ps1"
SCENARIOS = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"
DIAGNOSTIC = BOOTSTRAP / "phase5_gate_b_diagnostics.bas"
BUILD_STAGE_B = BOOTSTRAP / "build_stage_b.ps1"
SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"

DIAGNOSTIC_MODULE_NAME = "modPhase5GateBDiagnostics"

# The Phase-4 matrix, as the accepted harness reports it. The timeline chain D..J
# is ten sequential steps, so the matrix is 35 RESULTS.
PHASE4_SCENARIO_IDS = (
    "PRE0", "PRE", "A", "A1", "A2", "B", "B2", "C", "D0",
    "D-J.1", "D-J.2", "D-J.3", "D-J.4", "D-J.5",
    "D-J.6", "D-J.7", "D-J.8", "D-J.9", "D-J.10",
    "K", "K2", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W",
    "Y", "Z",
)

PHASE5_API = (
    "PCCM_Calculate",
    "PCCM_CalculationStatus",
    "PCCM_CalculationAttemptResult",
    "PCCM_CalculationAttemptDetail",
    "PCCM_CalculationFingerprint",
    "PCCM_CurrentInputFingerprint",
)

STATUS_ROW_IDS = ("P5-S1", "P5-S2", "P5-S3", "P5-S4", "P5-S5", "P5-S6")


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _text(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def _executable(path: Path) -> str:
    """PowerShell source with comment lines and here-string docs removed.

    A rule satisfied by a comment is not satisfied. Every structural assertion
    below runs over this, not over the raw file.
    """
    raw = _text(path)
    raw = re.sub(r"<#.*?#>", "", raw, flags=re.S)
    kept = []
    for line in raw.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            continue
        kept.append(re.sub(r"(?<!`)#.*$", "", line) if " #" in line else line)
    return "\n".join(kept)


def _vba_executable(path: Path) -> str:
    kept = []
    for line in _text(path).splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("'"):
            continue
        kept.append(stripped)
    return "\n".join(kept)


def _emitted() -> dict:
    """The build artifacts, PRODUCED by the real builder into a fresh temp tree.

    Never read from `build/`. An assertion about an emitted artifact that returns
    early when the file happens to be absent proves nothing - it passes loudest
    exactly when the build is broken.
    """
    from pccm_builder import (
        emit_calc_artifacts, emit_inspection, emit_stage_b, load_calc_contract,
        load_contract, load_driver_contract, load_spec, load_structure_contract,
    )

    tmp = Path(tempfile.mkdtemp(prefix="pccm-gateb-"))
    spec = load_spec(SPEC / "workbook.yaml")
    contract = load_contract(SPEC / "input_contract.yaml")
    drivers = load_driver_contract(SPEC / "driver_contract.yaml")
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    calc = load_calc_contract(SPEC / "calc_contract.yaml")

    stage_b = emit_stage_b(tmp, spec, contract, drivers, structure)
    calc_artifacts = emit_calc_artifacts(tmp, spec, calc)
    inspection = emit_inspection(tmp, calc, contract)
    for path in (stage_b.manifest_path, calc_artifacts.cases_path, inspection.path):
        assert path.is_file(), f"the builder produced no {path.name}"
    return {
        "manifest": json.loads(_text(stage_b.manifest_path)),
        "cases": json.loads(_text(calc_artifacts.cases_path)),
        "inspection": json.loads(_text(inspection.path)),
        "calc_module": _text(calc_artifacts.module_path),
        "constants": _text(stage_b.module_path),
        "dir": tmp,
    }


def _ps_string_literals(source: str) -> list[str]:
    return re.findall(r"'([^']*)'", source)


def _synthetic(body: str) -> str:
    """Planted PowerShell text for a negative control. Nothing is executed."""
    return body


# ===========================================================================
# 1. the harness extension is an EXTENSION
# ===========================================================================
def test_01_the_phase_4_harness_is_still_the_base() -> None:
    assert HARNESS.is_file(), "the accepted Phase-4 harness must still exist"
    source = _executable(HARNESS)
    assert ". (Join-Path $scriptDir 'com_lifecycle.ps1')" in source, (
        "the accepted COM lifecycle must still be the one in use"
    )
    assert ". (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')" in source, (
        "the Phase-5 scenarios must be dot-sourced INTO this harness"
    )
    assert "build_stage_b.ps1" in source, "the accepted bootstrap path must still be used"


def test_02_there_is_exactly_one_com_lifecycle_and_one_bootstrap() -> None:
    """No second Excel process, no second bootstrap, no second shutdown."""
    scenarios = _executable(SCENARIOS)
    for forbidden in ("New-Object -ComObject Excel.Application",   # refusal-list
                      "build_stage_b.ps1", "$excel.Quit()", "Get-PreExistingExcelPids",
                      "New-ReleaseLedger", "Invoke-EmergencyExcelCleanup"):
        assert forbidden not in scenarios, (
            f"the Phase-5 scenarios create a competing lifecycle ({forbidden})"
        )
    harness = _executable(HARNESS)
    assert harness.count("New-Object -ComObject Excel.Application") == 1, (   # refusal-list
        "more than one Excel instance is created"
    )
    assert harness.count("& $bootstrap") == 1, "more than one bootstrap invocation exists"


def test_03_every_phase_4_scenario_id_survives() -> None:
    """The accepted matrix is not rewritten to make room for Phase 5."""
    source = _text(HARNESS)
    reported = set(re.findall(r"Add-Result\s+'([^']+)'", source))
    reported |= {f"D-J.{index}" for index in range(1, 11)} if "'D-J.' + $stepIndex" in source else set()
    missing = [name for name in PHASE4_SCENARIO_IDS if name not in reported]
    assert not missing, f"Phase-4 scenario IDs disappeared: {missing}"


def test_04_the_full_phase_4_matrix_is_a_gate_b_prerequisite() -> None:
    """35/35, 0 FAIL, 0 SKIP - checked in source, not assumed in prose."""
    source = _executable(SCENARIOS)
    assert "Get-Phase4RequiredScenarioIds" in source
    declared = re.search(
        r"\$script:Phase4RequiredScenarioIds\s*=\s*@\((.*?)\)", source, re.S
    )
    assert declared, "the required Phase-4 matrix is not declared"
    names = _ps_string_literals(declared.group(1))
    assert len(names) == 35, f"the prerequisite names {len(names)} scenarios, not 35"
    assert set(names) == set(PHASE4_SCENARIO_IDS)
    assert "'the Phase-4 matrix has 0 FAIL'" in source
    assert "'the Phase-4 matrix has 0 SKIP'" in source
    assert re.search(r"\$passed\.Count -eq 35", source), "35/35 is not asserted"
    # A broken prerequisite is a FAIL, not a quiet SKIP.
    block = source[source.index("if (-not $phase4Ok)"):]
    assert "Add-Result 'P5-ALL'" in block[:600]
    assert "'FAIL'" in block[:600], "an unmet prerequisite is reported as something other than FAIL"
    assert "'SKIP'" not in block[:600]


# ===========================================================================
# 2. the expected-value authority
# ===========================================================================
def test_05_phase5_cases_is_loaded_from_the_supplied_build_dir() -> None:
    harness = _executable(HARNESS)
    assert "Join-Path $BuildDir 'phase5_cases.json'" in harness, (
        "the corpus must come from the supplied BuildDir, not from a fixed path"
    )
    assert "Join-Path $BuildDir 'phase5_gate_b_inspection.json'" in harness
    scenarios = _executable(SCENARIOS)
    assert "Join-Path $BuildDir 'phase5_cases.json'" in scenarios, (
        "the preflight must read the corpus it is validating"
    )


def test_06_no_expected_analytical_value_is_hand_maintained() -> None:
    """Not one oracle number is written into the harness.

    Every analytical comparison reads `$case.expected...`, `$vector.expected`,
    `$reference.digest` or a tolerance from the corpus. A bare numeric literal
    that looks like an oracle value - anything with a decimal point or an
    exponent - is a hand-copied expectation.
    """
    # Two kinds of numeric literal are legitimate and everything else is not:
    #
    #   arithmetic identities  0.0 as an accumulator seed, 1.0 as the relative
    #                          scale floor. Neither is a value the model produces.
    #   fixture INPUTS         0.99, the deliberately invalid profiling weight,
    #                          and 3.75, the rate on the deliberately UNREFERENCED
    #                          currency. Both are things the harness writes INTO
    #                          the workbook, never things it expects back.
    #
    # String literals are stripped first, because the locked vector LABELS
    # ('0.1', '1e-20', '0.1 + 0.2') are names the preflight checks the corpus
    # against, not numbers the harness expects.
    # Scoped to the PHASE-5 surface. The accepted Phase-4 harness states its own
    # structural fixture inputs - inflation rates, profiling percentages - and
    # Step B1 may not rewrite Phase-4 scenario semantics to satisfy a Phase-5
    # rule. test_09 covers what Phase 5 added to that file.
    allowed = {"0.0", "1.0", "0.99", "3.75"}
    for path in (SCENARIOS,):
        source = re.sub(r"'[^']*'", "''", _executable(path))
        source = re.sub(r'"[^"]*"', '""', source)
        # `Set-StrictMode -Version 2.0` is a PowerShell language level.
        source = re.sub(r"-Version\s+[\d.]+", "-Version", source)
        offenders = []
        for match in re.finditer(r"(?<![\w.\$])(-?\d+\.\d+(?:[eE][-+]?\d+)?)", source):
            if match.group(1) in allowed:
                continue
            offenders.append(match.group(1))
        assert not offenders, (
            f"{path.name} carries hand-maintained numeric expectations: {sorted(set(offenders))}"
        )


def test_06a_the_diagnostic_module_states_no_expected_value() -> None:
    """The VBA wrapper marshals and calls. It is not an oracle."""
    source = _vba_executable(DIAGNOSTIC)
    offenders = re.findall(r"(?<![\w.])(-?\d+\.\d+(?:[eE][-+]?\d+)?)", source)
    assert not offenders, (
        f"the diagnostic module carries numeric expectations: {sorted(set(offenders))}"
    )
    for forbidden in ("E+308", "E-324", "E-01", "E+20"):
        assert forbidden not in source, (
            f"the diagnostic restates a canonical expectation ({forbidden})"
        )


def test_07_the_locked_digest_and_stream_length_are_never_literals() -> None:
    for path in (SCENARIOS, HARNESS, DIAGNOSTIC):
        source = _text(path)
        assert "50B6EB0E26857EA7" not in source, (
            f"{path.name} hard-codes the reference digest instead of reading the corpus"
        )
        assert not re.search(r"\b366\b", _executable(path) if path.suffix == ".ps1"
                             else _vba_executable(path)), (
            f"{path.name} hard-codes the 366-unit count instead of reading the corpus"
        )


def test_08_the_expected_values_come_from_the_corpus_by_reference() -> None:
    source = _executable(SCENARIOS)
    for reference in ("$vector.expected", "$reference.digest", "$reference.code_units",
                      "$probe.digest", "$case.expected", "$vector.point", "$vector.comma",
                      "$Cases.tolerances"):
        assert reference in source, f"the harness never reads {reference} from the corpus"


def test_09_what_phase_5_added_to_the_phase_4_harness_is_plumbing_only() -> None:
    """The edit to the accepted harness is dot-source, load, copy and invoke.

    Nothing analytical was written into it, and no Phase-4 scenario was reshaped
    to make room.
    """
    source = _executable(HARNESS)
    added = [
        ". (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')",
        "Invoke-Phase5CoveragePreflight -BuildDir $BuildDir",
        "Join-Path $BuildDir 'phase5_cases.json'",
        "Join-Path $BuildDir 'phase5_gate_b_inspection.json'",
        "Invoke-Phase5GateBScenarios",
    ]
    for fragment in added:
        assert fragment in source, f"the harness is missing the Phase-5 plumbing: {fragment}"
    # The Phase-5 scenarios themselves live in their own file. The only two
    # P5 results the accepted harness may report are its own ERROR paths: the
    # preflight raising, and the scenario invocation raising.
    reported = set(re.findall(r"Add-Result '(P5-[A-Z]+)'", source))
    assert reported == {"P5-PRE", "P5-XX"}, (
        f"Phase-5 scenario bodies were written into the Phase-4 harness: {sorted(reported)}"
    )


# ===========================================================================
# 3. the 37-case coverage ledger
# ===========================================================================
def test_10_every_emitted_plan_case_is_mapped() -> None:
    """Driven from the CORPUS, so a case added upstream cannot go uncovered."""
    emitted = {str(case["id"]) for case in _emitted()["cases"]["plan_cases"]}
    source = _executable(SCENARIOS)
    block = source[source.index("function Get-Phase5CoverageLedger"):
                   source.index("function Get-Phase5ScenarioIds")]
    mapped = set()
    for match in re.finditer(r"\$ledger\.Add\('([^']+)'", block):
        mapped.add(match.group(1))
    for match in re.finditer(r"foreach \(\$id in ([\d, ]+)\)", block):
        for token in match.group(1).split(","):
            mapped.add(token.strip())
    missing = sorted(emitted - mapped, key=int)
    assert not missing, f"plan cases with no Windows scenario mapping: {missing}"
    ghosts = sorted(mapped - emitted, key=lambda value: int(value))
    assert not ghosts, f"ledger entries for cases the corpus does not emit: {ghosts}"
    assert len(emitted) == 37, f"the corpus emits {len(emitted)} plan cases, not 37"


def test_11_every_mapping_names_a_scenario_the_harness_defines() -> None:
    source = _executable(SCENARIOS)
    declared = set(_ps_string_literals(
        source[source.index("function Get-Phase5ScenarioIds"):
               source.index("function Get-Phase5FailpointNames")]
    ))
    referenced = set(re.findall(r"'(P5-[A-Z0-9]+)'", source))
    # Reported directly, or through the shared rollback runner's -ScenarioId.
    reported = set(re.findall(r"Add-Result\s+'(P5-[A-Z0-9]+)'", source))
    reported |= set(re.findall(r"-ScenarioId\s+'(P5-[A-Z0-9]+)'", source))
    unknown = sorted(name for name in declared if name not in reported)
    assert not unknown, f"scenario IDs declared but never reported: {unknown}"
    assert declared <= referenced


def test_12_the_coverage_preflight_runs_before_excel_is_started() -> None:
    """A missing mapping stops the run; it does not become a line in a summary."""
    source = _executable(HARNESS)
    preflight = source.index("Invoke-Phase5CoveragePreflight")
    excel = source.index("New-Object -ComObject Excel.Application")   # refusal-list
    assert preflight < excel, "the coverage preflight runs after Excel is started"
    tail = source[preflight:excel]
    assert "exit 1" in tail, "a failed coverage preflight does not abort the run"
    scenarios = _executable(SCENARIOS)
    body = scenarios[scenarios.index("function Invoke-Phase5CoveragePreflight"):]
    body = body[:body.index("\nfunction ", 1)] if "\nfunction " in body[1:] else body
    for forbidden in ("$Excel.Run", "ComObject", "$Workbook."):
        assert forbidden not in body, f"the preflight touches Excel ({forbidden})"


def test_13_the_preflight_checks_the_corpus_in_both_directions() -> None:
    source = _executable(SCENARIOS)
    body = source[source.index("function Invoke-Phase5CoveragePreflight"):]
    assert "'every emitted plan-case ID maps to a Windows scenario'" in body
    assert "'no ledger entry names a case the corpus does not emit'" in body
    assert "'every mapping names a scenario the harness defines'" in body
    assert "'every mapped fixture carries the evidence its kind promises'" in body
    # The unmapped scan is driven from $emitted, never from the ledger's own keys.
    unmapped = body[body.index("$unmapped = @()"):body.index("$orphan = @()")]
    assert "foreach ($id in $emitted)" in unmapped, (
        "the coverage scan iterates the ledger, so a missing entry cannot be seen"
    )


# ===========================================================================
# 4. the six API names, and the six status rows
# ===========================================================================
def test_14_all_six_api_procedures_are_exercised() -> None:
    source = _executable(SCENARIOS)
    for name in PHASE5_API:
        assert f"'{name}'" in source, f"{name} is never invoked by the harness"
    emitted = _emitted()["manifest"]["vba"]["api_procedures"]
    assert sorted(emitted) == sorted(PHASE5_API)
    # api_procedures is consumed AS api_procedures, not folded into entry_points.
    assert "$Manifest.vba.api_procedures" in source, (
        "the manifest's api_procedures projection is never read"
    )
    assert "'no API procedure is also an entry point'" in source
    assert "'no API procedure is bound to a button'" in source


def test_15_all_six_status_rows_exist_and_assert_all_four_accessors() -> None:
    source = _executable(SCENARIOS)
    for row in STATUS_ROW_IDS:
        assert f"Add-Result '{row}'" in source, f"status-matrix row {row} is missing"
    helper = source[source.index("function Add-StatusRowChecks"):]
    helper = helper[:helper.index("\n    }") + 6]
    for accessor in ("PCCM_CalculationStatus", "PCCM_CalculationAttemptResult",
                     "PCCM_CalculationAttemptDetail", "PCCM_CalculationFingerprint",
                     "PCCM_CurrentInputFingerprint"):
        assert accessor in helper, f"the status-row helper never reads {accessor}"


def test_16_the_six_rows_expect_the_locked_matrix() -> None:
    """Row 6 is STALE, and rows 5 and 6 keep the two axes in disagreement."""
    source = _executable(SCENARIOS)
    expectations = {
        "row 1": ("CURRENT", "SUCCESS"),
        "row 2": ("STALE", "SUCCESS"),
        "row 3": ("INVALID", "SUCCESS"),
        "row 4": ("INVALID", "REFUSED"),
        "row 5": ("CURRENT", "REFUSED"),
    }
    for row, (status, attempt) in expectations.items():
        call = re.search(
            rf"-Row '{row}' [^\n]*\n?[^\n]*-ExpectedStatus '([A-Z ]+)' -ExpectedAttempt '([A-Z]+)'",
            source,
        )
        assert call, f"{row} does not state an expected status and attempt"
        assert call.group(1) == status, f"{row} expects status {call.group(1)}, not {status}"
        assert call.group(2) == attempt, f"{row} expects attempt {call.group(2)}, not {attempt}"
    # Row 6 is the injected-failure row and is asserted inside the rollback
    # scenario: STALE, never CURRENT.
    rollback = source[source.index("function Invoke-Phase5RollbackScenario"):]
    assert "'PCCM_CalculationStatus() = STALE, not CURRENT'" in rollback
    assert "$status -eq 'STALE'" in rollback
    assert "'C19 is not FAILED: an attempt result may never be a status'" in rollback


def test_17_row_5_keeps_the_required_disagreement() -> None:
    source = _executable(SCENARIOS)
    assert "'row 5: CURRENT and a historical REFUSED coexist by design'" in source
    assert "'row 5: the refusal detail is STILL readable, unchanged'" in source
    assert "$row5.Detail -ceq $refusalDetail" in source, (
        "the refusal detail is not compared against the one that was recorded"
    )


def test_18_status_is_asked_for_before_it_is_claimed() -> None:
    """The status cell is last-evaluated, not live."""
    source = _executable(SCENARIOS)
    helper = source[source.index("function Add-StatusRowChecks"):]
    first = helper.index("$Excel.Run(")
    assert "PCCM_CalculationStatus" in helper[first:first + 60], (
        "the helper reads a stored status before asking for a fresh evaluation"
    )
    # And the raw C19 cell IS inspected too, so the endpoint can be cross-checked
    # against what the workbook actually holds rather than believed on its own.
    assert "State['calculation_status']" in source, (
        "the raw status cell is never inspected, so the endpoint cannot be cross-checked"
    )


# ===========================================================================
# 5. the transient diagnostic module
# ===========================================================================
def test_19_the_diagnostic_module_is_not_a_production_module() -> None:
    emitted = _emitted()
    declared = [module["name"] for module in emitted["manifest"]["vba"]["modules"]]
    assert DIAGNOSTIC_MODULE_NAME not in declared, (
        "the diagnostic module is declared in the production manifest"
    )
    assert len(declared) == 15, f"the manifest declares {len(declared)} modules, not 15"
    # Not in the structure contract either, so it can never be emitted into one.
    contract = _text(SPEC / "structure_contract.yaml")
    assert DIAGNOSTIC_MODULE_NAME not in contract
    # And the Stage-B bootstrap does not import it.
    assert DIAGNOSTIC_MODULE_NAME not in _text(BUILD_STAGE_B), (
        "build_stage_b.ps1 imports the diagnostic module into the production workbook"
    )
    assert "phase5_gate_b_diagnostics" not in _text(BUILD_STAGE_B)
    # It lives outside src/vba, so no production sweep can pick it up.
    assert not (SRC_VBA / f"{DIAGNOSTIC_MODULE_NAME}.bas").exists()
    assert DIAGNOSTIC.parent == BOOTSTRAP


def test_20_the_diagnostic_module_declares_no_production_endpoint() -> None:
    source = _vba_executable(DIAGNOSTIC)
    assert "PCCM_" not in source, "the diagnostic module declares or calls a PCCM_ endpoint"
    publics = re.findall(r"^Public (?:Sub|Function) (\w+)", source, re.M)
    assert publics, "the diagnostic module exposes nothing"
    for name in publics:
        assert name.startswith("GBD_"), f"{name} is not clearly test-only"
    # No button, no event handler, no persistence.
    for forbidden in ("OnAction", "Shapes", "Workbook_Open", "Worksheet_Change",
                      "SaveAs", "ThisWorkbook.Save", "Rnd", "Randomize"):
        assert forbidden not in source, f"the diagnostic module contains {forbidden}"


def test_21_the_diagnostic_module_only_wraps_accepted_public_helpers() -> None:
    """It marshals and calls. It computes nothing of its own."""
    source = _vba_executable(DIAGNOSTIC)
    called = sorted(set(re.findall(r"modCalc(?:Fingerprint|Analytical|Factors)\.(\w+)", source)))
    assert called == [
        "CalcFpCanonicalInteger", "CalcFpCanonicalNumber", "CalcFpCanonicalText",
        "CalcFpDigestStream", "CalcFpNormaliseCodeUnit", "CalcFpReduceDouble",
        "CalcFpUtf16Length", "PertMean", "TriangularMean", "UniformMean",
    ], f"unexpected diagnostic surface: {called}"
    # Every one of those is already Public in the accepted production source: the
    # diagnostic reopens nothing.
    fingerprint = _text(SRC_VBA / "modCalcFingerprint.bas")
    analytical = _text(SRC_VBA / "modCalcAnalytical.bas")
    for name in called:
        assert (re.search(rf"^Public Function {name}\b", fingerprint, re.M)
                or re.search(rf"^Public Function {name}\b", analytical, re.M)), (
            f"{name} is not Public in the accepted production source"
        )
    # No hash recurrence, no canonical formatting and no statistic of its own.
    for forbidden in ("FP_BASE", "FP_MOD_1", "FP_MOD_2", "Format(", "Format$(",
                      "Exp(", "Log("):
        assert forbidden not in source, f"the diagnostic module reimplements ({forbidden})"


def test_22_the_diagnostic_import_happens_only_after_the_a1_compile_proof() -> None:
    """A1 IS the first real VBA compilation boundary, and it must stay production-only."""
    harness = _executable(HARNESS)
    a1 = harness.index("Add-Result 'A1'")
    invoke = harness.index("Invoke-Phase5GateBScenarios")
    assert a1 < invoke, "the Phase-5 scenarios run before the A1 production compile proof"
    # A1 is still the FIRST Application.Run in the harness.
    first_run = harness.index("$excel.Run(")
    assert first_run < a1, "A1 does not contain the first Application.Run"
    assert "PCCM_AutomationBegin" in harness[first_run:first_run + 80], (
        "the first Application.Run is not a production procedure"
    )
    # And the import is inside the scenario file, after the prerequisite gate.
    scenarios = _executable(SCENARIOS)
    assert "$components.Import($source)" in scenarios
    imported = scenarios.index("$components.Import($source)")
    prerequisite = scenarios.index("Add-Result 'P5-P4'")
    assert prerequisite < imported, (
        "the diagnostic module is imported before the Phase-4 prerequisite is checked"
    )
    # Nothing in the harness or the bootstrap imports it earlier.
    assert "phase5_gate_b_diagnostics.bas" not in harness, (
        "the accepted harness imports the diagnostic module itself"
    )


def test_23_the_diagnostic_module_is_removed_after_the_vector_section() -> None:
    source = _executable(SCENARIOS)
    assert "$components.Remove($target)" in source, "the diagnostic module is never removed"
    removal = source.index("Add-Result 'P5-D8'")
    for vector in ("P5-D1", "P5-D2", "P5-D3", "P5-D4", "P5-D5", "P5-D6", "P5-D7"):
        assert source.index(f"Add-Result '{vector}'") < removal, (
            f"{vector} runs after the diagnostic module was removed"
        )
    tail = source[removal:]
    assert "'the diagnostic module is absent from the project'" in source
    assert "'the inventory is exactly the 15 manifest modules again'" in source
    assert "'no diagnostic procedure is callable any more'" in source
    # The inventory re-assertion happens BEFORE the analytical acceptance work.
    assert source.index("Add-Result 'P5-AN'") > removal, (
        "analytical acceptance runs while the diagnostic module is still installed"
    )
    # And no workbook is saved anywhere in the Phase-5 scenarios.
    for forbidden in (".Save()", ".SaveAs(", "SaveCopyAs"):
        assert forbidden not in source, f"the Phase-5 scenarios persist the workbook ({forbidden})"


# ===========================================================================
# 6. the direct vectors are complete
# ===========================================================================
def test_24_the_canonical_number_vectors_are_complete_and_locked() -> None:
    cases = _emitted()["cases"]
    labels = [vector["label"] for vector in cases["fingerprint"]["numeric_encodings"]["vectors"]]
    assert labels == ["0", "-0", "1", "-1", "0.1", "1e-20", "1e+20", "0.1 + 0.2",
                      "MAX_DOUBLE", "minimum subnormal"]
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-D1'") - 3000:source.index("Add-Result 'P5-D1'")]
    assert "$vectors = @($Cases.fingerprint.numeric_encodings.vectors)" in block
    assert "foreach ($vector in $vectors)" in block, "the vectors are not all iterated"
    assert "$vectors.Count -eq 10" in block, "the vector count is not asserted"
    # THE MINIMUM SUBNORMAL IS NOT SKIPPED, and it is also built on target.
    assert "'minimum subnormal'" in block
    assert "GBD_CanonicalNumberConstructed" in block, (
        "the extreme vectors are never constructed on target, so a lossy COM "
        "marshalling would be reported as an encoder defect"
    )
    diagnostic = _vba_executable(DIAGNOSTIC)
    assert "Case \"minimum subnormal\"" in diagnostic
    assert "For halving = 1 To 1074" in diagnostic, (
        "the minimum subnormal is not constructed by exact halving"
    )
    assert "value = MAX_DOUBLE" in diagnostic, (
        "MAX_DOUBLE is retyped rather than taken from the accepted kernel constant"
    )


def test_25_both_decimal_separators_are_injected() -> None:
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-D2'") - 3000:source.index("Add-Result 'P5-D2'")]
    assert "separator = '.'" in block and "separator = ','" in block, (
        "only one decimal separator is exercised"
    )
    assert "$vector.point" in block and "$vector.comma" in block
    assert "GBD_CanonicalNumber" in block
    # INJECTION, not environment. No regional setting is read or altered.
    for forbidden in ("Application.International", "UseSystemSeparators",
                      "DecimalSeparator", "CurrentCulture", "Set-Culture"):
        assert forbidden not in _executable(SCENARIOS), (
            f"the separator proof uses {forbidden} instead of injecting the argument"
        )
        assert forbidden not in _vba_executable(DIAGNOSTIC)


def test_26_all_four_reduction_vectors_are_exercised_in_vba() -> None:
    cases = _emitted()["cases"]
    assert len(cases["fingerprint"]["reduction_vectors"]) == 4
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-D3'") - 2500:source.index("Add-Result 'P5-D3'")]
    assert "$vectors = @($Cases.fingerprint.reduction_vectors)" in block
    assert "$vectors.Count -eq 4" in block, "the four-vector count is not asserted"
    assert "GBD_ReduceDouble" in block, "the accepted VBA reducer is never called"
    # PowerShell must not compute a reduction of its own and compare with itself.
    for forbidden in ("%", "[Math]::Floor", "-shr", "[math]::Truncate"):
        assert forbidden not in block, (
            f"the harness reduces in PowerShell ({forbidden}) instead of testing VBA"
        )
    diagnostic = _vba_executable(DIAGNOSTIC)
    assert "modCalcFingerprint.CalcFpReduceDouble" in diagnostic
    assert "Mod " not in diagnostic and " \\ " not in diagnostic


def test_27_the_utf16_vectors_including_non_bmp_are_exercised() -> None:
    cases = _emitted()["cases"]
    keys = [vector["key"] for vector in cases["fingerprint"]["utf16_vectors"]["vectors"]]
    assert keys == ["bmp_above_7fff", "non_bmp", "mixed_length_prefix"]
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-D4'") - 4000:source.index("Add-Result 'P5-D4'")]
    assert "$Cases.fingerprint.utf16_vectors.vectors" in block
    assert "GBD_Utf16Length" in block, "the unit count is never asked of VBA"
    assert "GBD_RawAscW" in block, "the SIGNED AscW result is never observed"
    assert "GBD_NormaliseCodeUnit" in block, "the normalisation is never exercised"
    assert "GBD_CanonicalTextField" in block, "the length prefix is never checked"
    assert "$key -eq 'non_bmp'" in block, "the non-BMP vector is not distinguished"
    assert "'the non-BMP character contributes TWO surrogate units, not one'" in block
    # The text is rebuilt from CODE UNITS, so no console encoding can reshape it.
    assert "$vector.code_units" in block
    assert "GBD_TextFromUnits" in _vba_executable(DIAGNOSTIC)


def test_28_the_reference_stream_asserts_units_and_digest_together() -> None:
    cases = _emitted()["cases"]
    reference = cases["fingerprint"]["reference"]
    assert reference["code_units"] == len(reference["stream"])
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-D5'") - 1600:source.index("Add-Result 'P5-D5'")]
    assert "GBD_StreamLength" in block, "the code-unit count is never asserted"
    assert "GBD_DigestStream" in block, "the digest is never asserted"
    assert "$reference.code_units" in block and "$reference.digest" in block
    # BOTH. A digest asserted alone would agree with itself over a truncated stream.
    length_at = block.index("GBD_StreamLength")
    digest_at = block.index("GBD_DigestStream")
    assert length_at < digest_at, "the digest is checked before the stream length"


def test_29_the_probe_and_statistics_vectors_are_exercised() -> None:
    cases = _emitted()["cases"]
    assert len(cases["fingerprint"]["collision_probes"]) == 8
    source = _executable(SCENARIOS)
    assert "GBD_ProbeDigest" in source
    assert "'every probe digest is distinct'" in source
    assert "GBD_ConvexStatistic" in source
    assert "$candidate.id -eq '28'" in source


# ===========================================================================
# 7. rollback, refusal and the failpoints
# ===========================================================================
def test_30_both_phase_5_failpoint_names_are_used() -> None:
    source = _executable(SCENARIOS)
    names = re.search(
        r"function Get-Phase5FailpointNames.*?AnalyticalWrite\s*=\s*'([^']+)'.*?"
        r"SuccessCommit\s*=\s*'([^']+)'",
        source, re.S,
    )
    assert names, "the harness declares no Phase-5 failpoint names"
    analytical, commit = names.group(1), names.group(2)
    # THE PRODUCTION MODULE IS THE AUTHORITY. These two strings are a checked
    # copy of the accepted declaration, not a second one.
    report = _text(SRC_VBA / "modCalcReport.bas")
    assert f'FAILPOINT_ANALYTICAL_WRITE As String = "{analytical}"' in report, (
        "the analytical failpoint name does not match the accepted production source"
    )
    assert f'FAILPOINT_SUCCESS_COMMIT As String = "{commit}"' in report, (
        "the commit failpoint name does not match the accepted production source"
    )
    # Both are actually armed, in two SEPARATE scenarios.
    assert f"-Failpoint $failpoints.AnalyticalWrite" in source
    assert f"-Failpoint $failpoints.SuccessCommit" in source
    assert "-ScenarioId 'P5-FA'" in source and "-ScenarioId 'P5-FC'" in source


def test_31_the_accepted_injection_mechanism_is_reused() -> None:
    source = _executable(SCENARIOS)
    assert "PCCM_AutomationBegin', $true, $Failpoint" in source, (
        "the failpoint is not armed through the accepted Phase-4 mechanism"
    )
    for invented in ("gAutomationFailAfterStage", "FailAfterStage", "Set-FailPoint",
                     "PCCM_InjectFailure"):
        assert invented not in source, f"a second injection system was created ({invented})"
    # And the production hook is where the source says it is: at the commit
    # assignment, not upstream of it. Gate B must exercise THAT hook.
    report = _text(SRC_VBA / "modCalcReport.bas")
    writer = report[report.index("Private Sub WriteSuccessCommit"):]
    writer = writer[:writer.index("End Sub")]
    lines = [line.strip() for line in writer.splitlines()
             if line.strip() and not line.strip().startswith("'")]
    hook = lines.index("modAppState.FailPointCheck FAILPOINT_SUCCESS_COMMIT")
    assignment = lines.index("CalcSheet.Range(CALC_STATE_VALUE_RANGE).Value2 = block")
    assert hook == assignment - 1, (
        "the production commit hook is no longer adjacent to the C13:C20 assignment"
    )


def test_32_the_rollback_comparison_uses_the_right_three_groups() -> None:
    """C13:C16, C23:C32 and the five tables. NOT all of C13:C20."""
    source = _executable(SCENARIOS)
    helper = source[source.index("function Add-SnapshotUnchangedChecks"):]
    helper = helper[:helper.index("\n    }\n\n") + 6]
    assert "$SuccessFields" in helper, "the success record is not compared as its own group"
    assert "$Before.Totals.Keys" in helper, "C23:C32 is not compared"
    assert "$Before.Tables.Keys" in helper, "the five tables are not compared"
    # The unchanged group is the SUCCESS RECORD only, and the attempt fields are
    # deliberately absent from it.
    declared = re.search(r"\$successRecordFields = @\((.*?)\)", source, re.S)
    assert declared, "the success-record field group is not declared"
    fields = set(_ps_string_literals(declared.group(1)))
    assert fields == {"last_successful_stamp", "last_successful_fingerprint",
                      "fingerprint_version", "last_successful_applied_timeline"}
    attempt = re.search(r"\$attemptFields = @\((.*?)\)", source, re.S)
    assert attempt, "the attempt field group is not declared"
    attempt_fields = set(_ps_string_literals(attempt.group(1)))
    assert attempt_fields == {"last_attempt_result", "last_attempt_detail",
                              "calculation_status", "status_evaluated_at"}
    assert not (fields & attempt_fields), "the two mutation groups overlap"
    # And the whole eight-cell range is NEVER asserted unchanged.
    assert "value_range" not in helper, (
        "the comparison reaches for the whole C13:C20 range instead of the two groups"
    )


def test_33_the_rollback_asserts_every_required_final_state() -> None:
    source = _executable(SCENARIOS)
    block = source[source.index("function Invoke-Phase5RollbackScenario"):
                   source.index("$analyticalOk = Invoke-Phase5RollbackScenario")]
    for claim in ("'C17 = FAILED'",
                  "'C18 carries a specific failure detail'",
                  "'C19 is a freshly DERIVED status, not the attempt result'",
                  "'C20 carries a fresh evaluation timestamp'",
                  "'PCCM_CalculationStatus() = STALE, not CURRENT'",
                  "'no mixed old/new analytical state survived the rollback'",
                  "'EnableEvents was restored'",
                  "'ScreenUpdating was restored'",
                  "'Calculation mode was restored to automatic'"):
        assert claim in block, f"the rollback scenario never asserts {claim}"
    assert "Add-SnapshotUnchangedChecks" in block, (
        "the rollback never compares against the previous successful snapshot"
    )


def test_34_the_refusal_compares_the_two_groups_separately() -> None:
    source = _executable(SCENARIOS)
    row4 = source[source.index("$Excel.Run('PCCM_Calculate') | Out-Null\n        $row4"):
                  source.index("Add-Result 'P5-S4'")]
    assert "Add-SnapshotUnchangedChecks" in row4, (
        "row 4 never asserts the prior snapshot survived"
    )
    for changed in ("'row 4: C17 CHANGED to REFUSED'",
                    "'row 4: C18 CHANGED to a specific refusal detail'",
                    "'row 4: C19 is the freshly derived status'",
                    "'row 4: C20 carries a status-evaluation timestamp'"):
        assert changed in row4, (
            f"row 4 does not assert that the attempt axis CHANGED ({changed}); "
            "asserting all of C13:C20 unchanged would assert the refusal was never recorded"
        )
    # No partial analytical output may survive a refusal.
    assert "'carries no refused output'" in source or "carries no refused output" in source


def test_35_the_analytical_scenario_asserts_every_emitted_value() -> None:
    """Not "it succeeded" plus a handful of totals."""
    source = _executable(SCENARIOS)
    checks = source[source.index("function Add-Phase5AnalyticalChecks"):
                    source.index("function Invoke-Phase5GateBScenarios")]
    for table in ("calc_years", "calc_inflation_factors", "calc_fx", "calc_drivers",
                  "calc_annual"):
        assert f"'{table}'" in checks, f"{table} is never compared"
    assert "-Block 'calc_totals'" in checks, "calc_totals is never compared"
    # The driver and annual comparisons are driven from the FIXTURE's own field
    # names, so a field added upstream is asserted without editing the harness.
    assert "foreach ($field in $wanted.PSObject.Properties.Name)" in checks
    assert "row count" in checks, "a row-count mismatch is not itself a failure"


def test_36_blank_is_never_equal_to_numeric_zero() -> None:
    source = _executable(SCENARIOS)
    body = source[source.index("function Test-CalcValue"):]
    body = body[:body.index("\nfunction ")]
    assert "if ($null -eq $Expected) { return (Test-CalcBlank -Actual $Actual) }" in body, (
        "an expected BLANK is not compared as a blank"
    )
    assert "if (Test-CalcBlank -Actual $Actual) { return $false }" in body, (
        "a BLANK can satisfy a numeric expectation"
    )
    blank = source[source.index("function Test-CalcBlank"):]
    blank = blank[:blank.index("\nfunction ")]
    # A ZERO-LENGTH STRING is blank; a numeric zero is not. The only length test
    # allowed here is on .Length, never on the value.
    assert "$Actual -eq 0" not in blank and "-eq 0.0" not in blank, (
        "the blank test coerces a numeric zero to blank"
    )
    assert "$Actual.Length -eq 0" in blank, "an empty string is not treated as blank"


# ===========================================================================
# 8. what must NOT have happened
# ===========================================================================
def test_37_no_calculate_button_and_no_new_production_module() -> None:
    emitted = _emitted()
    buttons = emitted["manifest"]["buttons"]
    assert len(buttons) == 5, f"the manifest declares {len(buttons)} buttons, not 5"
    entry_points = {button["entry_point"] for button in buttons}
    assert "PCCM_Calculate" not in entry_points, "a Calculate button was introduced"
    assert entry_points == set(emitted["manifest"]["vba"]["entry_points"])
    modules = {module["name"] for module in emitted["manifest"]["vba"]["modules"]}
    assert len(modules) == 15
    on_disk = {path.stem for path in SRC_VBA.glob("*.bas")}
    assert DIAGNOSTIC_MODULE_NAME not in on_disk
    assert len(on_disk) == 13, f"a production module was added or removed: {sorted(on_disk)}"
    # The harness asserts all three of those things at runtime too.
    source = _executable(SCENARIOS)
    assert "'NO shape has OnAction = PCCM_Calculate'" in source
    assert "'exactly five command buttons persist in the workbook'" in source
    assert "'the manifest declares 15 production modules'" in source
    assert "'the module ' + $name + ' persists in the saved project'" in source


def test_38_the_production_modules_are_asserted_by_name_not_by_count() -> None:
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Result 'P5-M'") - 5000:source.index("Add-Result 'P5-M'")]
    assert "foreach ($name in $expected)" in block, "the modules are not checked by name"
    assert "$present -contains $name" in block
    # And in the other direction: no module outside the manifest may persist.
    assert "'no module outside the manifest persists'" in block


def test_39_no_linux_test_here_executes_windows_or_claims_a_run() -> None:
    """This suite reads source. It starts nothing.

    The forbidden tokens are ASSEMBLED rather than written out, because a literal
    list of them inside this file would match itself and make the sweep either
    permanently red or permanently excused.
    """
    # Lines that are themselves REFUSALS - a token inside a `forbidden` /
    # `for ... in (...)` tuple that another test asserts the ABSENCE of - are not
    # execution paths. They are excluded by their marker, not by their content.
    kept = []
    for line in _text(Path(__file__)).splitlines():
        if "# refusal-list" in line:
            continue
        kept.append(line)
    source = "\n".join(kept)
    forbidden = ("Start-" + "Process", "power" + "shell", "p" + "wsh",
                 "Excel." + "Application", "win32" + "com", "subprocess." + "run")
    for token in forbidden:
        assert token not in source, (
            f"this suite reaches for a Windows execution path ({token})"
        )
    lowered = source.lower()
    for claim in ("gate b pass" + "ed", "vba exec" + "uted", "excel con" + "firmed",
                  "the run pro" + "ved", "phase 5 is acc" + "epted"):
        assert claim not in lowered, f"this suite claims a runtime result: {claim!r}"


def test_40_the_harness_never_writes_to_the_real_build_directory() -> None:
    """The accepted rule: copy the build, drive the copy."""
    harness = _executable(HARNESS)
    assert "$tempRoot = Join-Path ([System.IO.Path]::GetTempPath())" in harness
    assert "Copy-Item -LiteralPath $casesPath -Destination $tempRoot" in harness
    assert "Copy-Item -LiteralPath $inspectPath -Destination $tempRoot" in harness
    scenarios = _executable(SCENARIOS)
    for forbidden in ("Set-Content", "Out-File", "Remove-Item", "New-Item"):
        assert forbidden not in scenarios, (
            f"the Phase-5 scenarios write to disk ({forbidden})"
        )


# ===========================================================================
# 9. NEGATIVE CONTROLS
#
# Each plants the regression the rule exists to prevent, as SYNTHETIC text, and
# asserts the detector above sees it. A detector nobody has watched fail is a
# detector nobody has tested. Nothing here is written to disk and nothing runs.
# ===========================================================================
def _ledger_block(source: str) -> str:
    return source[source.index("function Get-Phase5CoverageLedger"):
                  source.index("function Get-Phase5ScenarioIds")]


def _mapped_ids(block: str) -> set[str]:
    mapped = {match.group(1) for match in re.finditer(r"\$ledger\.Add\('([^']+)'", block)}
    for match in re.finditer(r"foreach \(\$id in ([\d, ]+)\)", block):
        for token in match.group(1).split(","):
            mapped.add(token.strip())
    return mapped


def test_nc_01_an_omitted_plan_case_is_caught() -> None:
    emitted = {str(case["id"]) for case in _emitted()["cases"]["plan_cases"]}
    block = _ledger_block(_executable(SCENARIOS))
    planted = block.replace("$ledger.Add('30', @('P5-AN', 'P5-ID'))", "")
    assert "30" not in _mapped_ids(planted), "the dropped mapping must be visible"
    assert emitted - _mapped_ids(planted) == {"30"}


def test_nc_02_a_case_dropped_from_a_grouped_list_is_caught() -> None:
    block = _ledger_block(_executable(SCENARIOS))
    planted = block.replace("foreach ($id in 14, 15, 16, 17, 18, 20, 23, 24, 29)",
                            "foreach ($id in 14, 15, 16, 17, 18, 20, 23, 24)")
    assert planted != block, "the plant did not apply"
    assert "29" not in _mapped_ids(planted), (
        "a case silently dropped from a filtered list must be visible"
    )


def test_nc_03_a_mapping_to_a_nonexistent_scenario_is_caught() -> None:
    source = _executable(SCENARIOS)
    declared = set(_ps_string_literals(
        source[source.index("function Get-Phase5ScenarioIds"):
               source.index("function Get-Phase5FailpointNames")]
    ))
    assert "P5-NOPE" not in declared
    reported = set(re.findall(r"Add-Result\s+'(P5-[A-Z0-9]+)'", source))
    reported |= set(re.findall(r"-ScenarioId\s+'(P5-[A-Z0-9]+)'", source))
    planted = declared | {"P5-NOPE"}
    assert sorted(name for name in planted if name not in reported) == ["P5-NOPE"], (
        "a declared-but-never-reported scenario must be visible"
    )


def test_nc_04_a_hard_coded_expected_total_is_caught() -> None:
    planted = _synthetic(
        "$null = Add-Check $list 'the total is right' ($total -eq 1234.56)\n"
    )
    offenders = re.findall(r"(?<![\w.\$])(-?\d+\.\d+(?:[eE][-+]?\d+)?)", planted)
    assert offenders == ["1234.56"], "a hand-copied oracle value must be visible"
    assert "$case.expected" not in planted


def test_nc_05_a_hard_coded_reference_digest_is_caught() -> None:
    planted = _synthetic("$null = Add-Check $list 'digest' ($d -eq '50B6EB0E26857EA7')\n")
    assert "50B6EB0E26857EA7" in planted, "the hard-coded digest must be visible"


def test_nc_06_a_removed_phase_4_prerequisite_is_caught() -> None:
    source = _executable(SCENARIOS)
    planted = source.replace("$passed.Count -eq 35", "$passed.Count -ge 0")
    assert "$passed.Count -eq 35" not in planted, "the weakened prerequisite must be visible"
    planted = source.replace("'the Phase-4 matrix has 0 SKIP'", "'skips are fine'")
    assert "'the Phase-4 matrix has 0 SKIP'" not in planted


def test_nc_07_a_prerequisite_failure_reported_as_skip_is_caught() -> None:
    planted = _synthetic(
        "if (-not $phase4Ok) {\n"
        "    Add-Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'SKIP' 'phase 4 not intact'\n"
        "    return\n}\n"
    )
    block = planted[planted.index("if (-not $phase4Ok)"):]
    assert "'SKIP'" in block[:600], "a prerequisite failure hidden as a SKIP must be visible"
    assert "'FAIL'" not in block[:600]


def test_nc_08_a_module_checked_only_by_count_is_caught() -> None:
    planted = _synthetic(
        "$null = Add-Check $list 'fifteen modules' ($present.Count -eq 15)\n"
    )
    assert "foreach ($name in $expected)" not in planted, (
        "a count-only inventory check must be visible"
    )
    assert "$present -contains $name" not in planted


def test_nc_09_a_calculate_button_is_caught() -> None:
    emitted = _emitted()
    buttons = [dict(button) for button in emitted["manifest"]["buttons"]]
    buttons.append({"key": "calculate", "sheet": "Setup", "shape_name": "btnPCCMCalculate",
                    "caption": "Calculate", "entry_point": "PCCM_Calculate",
                    "anchor_cell": "E50", "width": 150.0, "height": 30.0})
    entry_points = {button["entry_point"] for button in buttons}
    assert len(buttons) == 6, "the extra button must be visible"
    assert "PCCM_Calculate" in entry_points, "the Calculate binding must be visible"


def test_nc_10_an_omitted_status_row_is_caught() -> None:
    source = _executable(SCENARIOS)
    planted = source.replace("Add-Result 'P5-S5'", "Add-Result 'P5-SX'")
    assert "Add-Result 'P5-S5'" not in planted, "the missing status row must be visible"
    present = [row for row in STATUS_ROW_IDS if f"Add-Result '{row}'" in planted]
    assert len(present) == 5


def test_nc_11_status_row_6_expecting_current_is_caught() -> None:
    source = _executable(SCENARIOS)
    rollback = source[source.index("function Invoke-Phase5RollbackScenario"):]
    planted = rollback.replace("$status -eq 'STALE'", "$status -eq 'CURRENT'")
    assert "$status -eq 'STALE'" not in planted, (
        "a row-6 expectation of CURRENT instead of STALE must be visible"
    )


def test_nc_12_row_4_comparing_all_of_c13_c20_is_caught() -> None:
    planted = _synthetic(
        "$before = Get-CalcScalarBlock -Workbook $wb -Inspection $i -Block 'calc_state'\n"
        "foreach ($field in $before.Keys) {\n"
        "    $null = Add-Check $list 'unchanged' (Test-CalcValue -Actual $after[$field] "
        "-Expected $before[$field])\n}\n"
    )
    assert "$before.Keys" in planted, (
        "a comparison over the WHOLE eight-cell block must be visible; it would "
        "assert that the refusal was never recorded"
    )
    for changed in ("C17 CHANGED", "C18 CHANGED"):
        assert changed not in planted, "the missing attempt-axis assertions must be visible"


def test_nc_13_a_rollback_comparison_missing_a_table_is_caught() -> None:
    planted = _synthetic(
        "foreach ($key in 'calc_years', 'calc_fx', 'calc_drivers', 'calc_annual') {\n"
        "    $null = Add-Check $List ($Label + ': ' + $key + ' is unchanged') $identical\n}\n"
    )
    assert "calc_inflation_factors" not in planted, "the omitted table must be visible"
    assert "$Before.Tables.Keys" not in planted, (
        "iterating a hand-written table list instead of the projection must be visible"
    )


def test_nc_14_a_rollback_comparison_missing_c23_c32_is_caught() -> None:
    planted = _synthetic(
        "foreach ($field in $SuccessFields) { $null = Add-Check $List 'state' $ok }\n"
        "foreach ($key in $Before.Tables.Keys) { $null = Add-Check $List 'tables' $ok }\n"
    )
    assert "$Before.Totals.Keys" not in planted, "the omitted calc_totals group must be visible"


def test_nc_15_a_missing_failpoint_is_caught() -> None:
    source = _executable(SCENARIOS)
    for dropped, kept in (("SuccessCommit", "AnalyticalWrite"),
                          ("AnalyticalWrite", "SuccessCommit")):
        planted = re.sub(rf"-Failpoint \$failpoints\.{dropped}[^\n]*\n", "", source)
        assert f"-Failpoint $failpoints.{dropped}" not in planted, (
            f"the missing {dropped} injection must be visible"
        )
        assert f"-Failpoint $failpoints.{kept}" in planted


def test_nc_16_the_diagnostic_module_in_the_production_manifest_is_caught() -> None:
    emitted = _emitted()
    modules = [dict(module) for module in emitted["manifest"]["vba"]["modules"]]
    modules.append({"name": DIAGNOSTIC_MODULE_NAME, "generated": False,
                    "responsibility": "diagnostics"})
    names = [module["name"] for module in modules]
    assert DIAGNOSTIC_MODULE_NAME in names, "the smuggled diagnostic module must be visible"
    assert len(names) == 16, "the inventory growth must be visible"


def test_nc_17_the_diagnostic_imported_before_a1_is_caught() -> None:
    planted = _synthetic(
        "$components.Import($source)\n"
        "$excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null\n"
        "Add-Result 'A1' 'VBA automation surface callable' 'PASS' ''\n"
    )
    assert planted.index("$components.Import($source)") < planted.index("Add-Result 'A1'"), (
        "an import that precedes the A1 production compile proof must be visible"
    )


def test_nc_18_a_diagnostic_module_never_removed_is_caught() -> None:
    source = _executable(SCENARIOS)
    planted = source.replace("$components.Remove($target)", "")
    assert "$components.Remove($target)" not in planted, (
        "the missing removal must be visible"
    )


def test_nc_19_a_pccm_endpoint_in_the_diagnostic_module_is_caught() -> None:
    planted = _synthetic(
        'Public Function PCCM_GateBProbe() As String\n'
        '    PCCM_GateBProbe = "OK"\n'
        'End Function\n'
    )
    assert "PCCM_" in planted, "a production endpoint in the diagnostic module must be visible"
    publics = re.findall(r"^Public (?:Sub|Function) (\w+)", planted, re.M)
    assert [name for name in publics if not name.startswith("GBD_")] == ["PCCM_GateBProbe"]


def test_nc_20_only_the_point_separator_tested_is_caught() -> None:
    planted = _synthetic(
        "foreach ($vector in $vectors) {\n"
        "    $reply = [string]$Excel.Run('GBD_CanonicalNumber', [double]$vector.value, '.')\n}\n"
    )
    assert "separator = ','" not in planted and "'.'" in planted, (
        "a point-only separator proof must be visible"
    )
    assert "$vector.comma" not in planted


def test_nc_21_an_omitted_reduction_vector_is_caught() -> None:
    planted = _synthetic(
        "$vectors = @($Cases.fingerprint.reduction_vectors) | Select-Object -First 3\n"
        "foreach ($vector in $vectors) { $null = $Excel.Run('GBD_ReduceDouble') }\n"
    )
    assert "$vectors.Count -eq 4" not in planted, "the missing four-vector count must be visible"
    assert "-First 3" in planted


def test_nc_22_an_omitted_non_bmp_vector_is_caught() -> None:
    source = _executable(SCENARIOS)
    planted = source.replace("$key -eq 'non_bmp'", "$key -eq 'never'")
    assert "$key -eq 'non_bmp'" not in planted, "the dropped non-BMP vector must be visible"
    planted = _synthetic(
        "foreach ($vector in $vectors) { if ($vector.key -eq 'non_bmp') { continue } }\n"
    )
    assert "continue" in planted


def test_nc_23_a_digest_asserted_without_the_unit_count_is_caught() -> None:
    planted = _synthetic(
        "$digest = [string]$Excel.Run('GBD_DigestStream', $stream)\n"
        "$null = Add-Check $list 'digest' ($digest -eq ('OK|' + $reference.digest))\n"
    )
    assert "GBD_StreamLength" not in planted, (
        "a digest asserted alone would agree with itself over a truncated stream"
    )
    assert "$reference.code_units" not in planted


def test_nc_24_a_manual_value_substituted_for_the_corpus_is_caught() -> None:
    planted = _synthetic(
        "$wanted = '1.0000000000000000E+00'\n"
        "$null = Add-Check $list 'canonical' ($reply -eq ('OK|' + $wanted))\n"
    )
    assert "$vector.expected" not in planted, "the substituted expectation must be visible"
    assert "E+00" in planted


def test_nc_25_a_blank_compared_as_zero_is_caught() -> None:
    planted = _synthetic(
        "function Test-CalcValue {\n"
        "    param($Actual, $Expected)\n"
        "    return ([double]$Actual -eq [double]$Expected)\n"
        "}\n"
    )
    assert "Test-CalcBlank" not in planted, (
        "a comparison that coerces both sides to Double must be visible; it would "
        "report a fabricated zero in an N/A field as correct"
    )


def test_nc_26_a_second_injection_mechanism_is_caught() -> None:
    planted = _synthetic(
        "$wb.Names.Item('gAutomationFailAfterStage').RefersToRange.Value2 = 'Commit'\n"
    )
    assert "gAutomationFailAfterStage" in planted, (
        "a hand-rolled injection point must be visible"
    )
    assert "PCCM_AutomationBegin" not in planted


def test_nc_27_a_second_com_lifecycle_is_caught() -> None:
    planted = _synthetic(
        "$excel2 = New-Object -ComObject Excel.Application\n"   # refusal-list
        "$excel2.Visible = $false\n$excel2.Quit()\n"
    )
    assert "New-Object -ComObject Excel.Application" in planted, (   # refusal-list
        "a second Excel instance must be visible"
    )
    assert planted.count("Quit()") == 1


def test_nc_28_an_analytical_check_of_only_the_totals_is_caught() -> None:
    planted = _synthetic(
        "$null = Add-Check $list 'it calculated' ($attempt -eq 'SUCCESS')\n"
        "foreach ($field in 'e_nom', 'e_pv') {\n"
        "    $null = Add-Check $list 'total' (Test-CalcValue -Actual $a -Expected $b)\n}\n"
    )
    for table in ("calc_years", "calc_inflation_factors", "calc_fx", "calc_drivers",
                  "calc_annual"):
        assert table not in planted, f"the unchecked {table} must be visible"


# ===========================================================================
# 10. the inspection projection has exactly ONE authority
# ===========================================================================
# `phase5_gate_b_inspection.json` exists because no other build output projects
# the `_Calc` layout, the Setup input scalars or the Config lookup tables, and
# the harness must not learn an address by parsing VBA. It is a PROJECTION, and
# these tests are what keep it from becoming a second contract: every address in
# it is pinned against the generated modules the same authorities produce.
def test_41_the_inspection_projection_carries_no_expected_value() -> None:
    inspection = _emitted()["inspection"]
    blob = json.dumps(inspection)
    assert "tolerance" not in blob.lower(), "the projection states a tolerance"
    for key in ("expected", "digest", "annual", "totals", "drivers"):
        assert f'"{key}"' not in blob, f"the projection carries an expected value ({key})"
    assert set(inspection.keys()) == {
        "schema_version", "purpose", "provenance", "calc", "inputs", "input_tables"
    }
    # It does not repeat what the manifest already projects.
    assert "registers" not in inspection and "grids" not in inspection


def test_42_every_calc_address_matches_the_generated_module() -> None:
    emitted = _emitted()
    inspection = emitted["inspection"]["calc"]
    module = emitted["calc_module"]

    def constant(name: str) -> str:
        found = re.search(rf'^Public Const {name} As \w+ = "?([^"\n]+?)"?\s*(?:\'.*)?$',
                          module, re.M)
        assert found, f"{name} is not declared in the generated module"
        return found.group(1).strip()

    state = inspection["scalar_blocks"]["calc_state"]
    assert state["value_range"] == constant("CALC_STATE_VALUE_RANGE")
    totals = inspection["scalar_blocks"]["calc_totals"]
    assert totals["value_range"] == constant("CALC_TOTALS_VALUE_RANGE")
    for key, row in state["rows"].items():
        assert str(row) == constant(f"CALC_STATE_ROW_{key.upper()}"), (
            f"calc_state.{key} disagrees with the generated module"
        )
    names = {
        "calc_years": "TBL_CALC_YEARS",
        "calc_inflation_factors": "TBL_CALC_INFLATION_FACTORS",
        "calc_fx": "TBL_CALC_FX",
        "calc_drivers": "TBL_CALC_DRIVERS",
        "calc_annual": "TBL_CALC_ANNUAL",
    }
    assert set(inspection["tables"]) == set(names)
    for key, prefix in names.items():
        table = inspection["tables"][key]
        assert table["table_name"] == constant(prefix)
        assert str(table["header_row"]) == constant(f"{prefix}_HEADER_ROW")
        assert table["first_column"] == constant(f"{prefix}_FIRST_COLUMN")
        assert table["last_column"] == constant(f"{prefix}_LAST_COLUMN")
        assert str(table["column_count"]) == constant(f"{prefix}_COLUMN_COUNT")
        assert str(table["first_body_row"]) == constant(f"{prefix}_FIRST_BODY_ROW")


def test_43_every_input_identity_matches_the_generated_constants() -> None:
    emitted = _emitted()
    inspection = emitted["inspection"]
    constants = emitted["constants"]
    for key, spec in inspection["inputs"].items():
        assert f'"{spec["defined_name"]}"' in constants, (
            f"the input {key} names a defined name the constants module does not declare"
        )
    for key, table in inspection["input_tables"].items():
        assert f'"{table["table_name"]}"' in constants, (
            f"the input table {key} names a table the constants module does not declare"
        )
    # The two the fixtures actually need most.
    assert inspection["inputs"]["discount_rate"]["defined_name"] == "inpDiscountRate"
    assert inspection["input_tables"]["fx_rates"]["table_name"] == "tblFXRates"


def test_44_the_projection_is_emitted_by_the_stage_a_build() -> None:
    """Not written by hand, and not left to be forgotten."""
    builder = _text(PCCM_ROOT / "builder" / "build_stage_a.py")
    assert "emit_inspection(" in builder, "the Stage-A build does not emit the projection"
    assert "phase5_gate_b_inspection.json" not in builder, (
        "the build names the artifact itself instead of letting the emitter own it"
    )
    # Identities only: the emitter must not IMPORT the case corpus or an oracle.
    # Read from executable text, so the module's own explanation of why it is not
    # an expected-value authority does not trip the check that it is not one.
    module = PCCM_ROOT / "builder" / "pccm_builder" / "gate_b_inspection.py"
    code = "\n".join(
        line for line in _text(module).splitlines()
        if not line.lstrip().startswith("#")
    )
    code = re.sub(r'"""(?:.|\n)*?"""', "", code)
    for forbidden in ("calc_cases", "calc_oracle", "calc_numeric", "expected"):
        assert forbidden not in code, f"the projection emitter reaches for {forbidden}"


def test_45_the_harness_never_parses_vba_to_find_an_address() -> None:
    """The reason the projection exists at all."""
    for path in (SCENARIOS, HARNESS):
        source = _executable(path)
        for forbidden in ("modCalcContract.bas", "modConstants.bas",
                          "Public Const", "CALC_STATE_VALUE_RANGE", "TBL_CALC_"):
            assert forbidden not in source, (
                f"{path.name} parses generated VBA for an address ({forbidden}); "
                "that is a second reader of the same authority"
            )
    source = _executable(SCENARIOS)
    assert "$Inspection.calc" in source, "the projection is never consulted"
    assert "$Inspection.inputs" in source


# ===========================================================================
# 11. this suite makes no runtime claim
# ===========================================================================
def test_46_the_harness_source_states_that_no_run_has_happened() -> None:
    for path in (SCENARIOS, DIAGNOSTIC):
        assert "NOT" in _text(path) or "not been" in _text(path)
    assert "NO PHASE-5 GATE-B RUN HAS BEEN MADE" in _text(HARNESS), (
        "the harness does not state that its Phase-5 extension is unrun"
    )


def test_47_the_transient_module_is_never_persisted() -> None:
    source = _executable(SCENARIOS)
    removal = source.index("Add-Result 'P5-D8'")
    for forbidden in (".Save()", ".SaveAs(", "SaveCopyAs", "xlOpenXMLWorkbookMacroEnabled"):
        assert forbidden not in source, f"the Phase-5 scenarios persist a workbook ({forbidden})"
    harness = _executable(HARNESS)
    # The accepted harness closes without saving, and that is unchanged.
    assert "$wb.Close($false)" in harness, "the workbook is no longer closed without saving"
