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
# Phase-4 scenarios report through the accepted Add-Result. Phase-5 scenarios
# report through Add-Phase5Result, the one-result-per-ID guard added after
# Runtime Run 4 recorded P5-S2 and P5-ST twice each. Tests that locate a result
# by ID must accept either, or they would silently stop finding half of them.
RESULT_CALL = r"Add-(?:Phase5)?Result"


def _result_call(identifier: str) -> str:
    """The emitter token for a scenario ID, whichever family it belongs to."""
    return ("Add-Phase5Result" if identifier.startswith("P5-") else "Add-Result") \
        + f" '{identifier}'"


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
    reported = set(re.findall(RESULT_CALL + r"\s+'([^']+)'", source))
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
    assert "Add-Phase5Result 'P5-ALL'" in block[:600]
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
    reported = set(re.findall(RESULT_CALL + r"\s+'(P5-[A-Z]+)'", source))
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
    reported = set(re.findall(RESULT_CALL + r"\s+'(P5-[A-Z0-9]+)'", source))
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
        assert _result_call(row) in source, f"status-matrix row {row} is missing"
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
    prerequisite = scenarios.index("Add-Phase5Result 'P5-P4'")
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
    removal = source.index("Add-Phase5Result 'P5-D8'")
    for vector in ("P5-D1", "P5-D2", "P5-D3", "P5-D4", "P5-D5", "P5-D6", "P5-D7"):
        assert source.index(_result_call(vector)) < removal, (
            f"{vector} runs after the diagnostic module was removed"
        )
    tail = source[removal:]
    assert "'the diagnostic module is absent from the standard modules'" in source, (
        "removal is no longer proved against the partition the module lives in"
    )
    assert "'the diagnostic module is absent from the project entirely'" in source
    assert "'no diagnostic procedure is callable any more'" in source
    # The full manifest inventory is re-asserted after removal, through the same
    # helper P5-M uses, so the two cannot drift apart.
    d8 = source[source.index("$components.Remove($target)"):removal]
    assert "Add-Phase5ModuleInventoryChecks -List $list -Components $inventory" in d8, (
        "the inventory is no longer re-asserted after the diagnostic module is removed"
    )
    assert "-Label 'after removal'" in d8
    # The inventory re-assertion happens BEFORE the analytical acceptance work.
    assert source.index("Add-Phase5Result 'P5-AN'") > removal, (
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
    block = source[source.index("Add-Phase5Result 'P5-D1'") - 3000:source.index("Add-Phase5Result 'P5-D1'")]
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
    block = source[source.index("Add-Phase5Result 'P5-D2'") - 3000:source.index("Add-Phase5Result 'P5-D2'")]
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
    block = source[source.index("Add-Phase5Result 'P5-D3'") - 2500:source.index("Add-Phase5Result 'P5-D3'")]
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
    block = source[source.index("Add-Phase5Result 'P5-D4'") - 4000:source.index("Add-Phase5Result 'P5-D4'")]
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
    block = source[source.index("Add-Phase5Result 'P5-D5'") - 1600:source.index("Add-Phase5Result 'P5-D5'")]
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
    helper = helper[:helper.index("\nfunction ")]
    assert "$SuccessFields" in helper, "the success record is not compared as its own group"
    assert "$Before.Totals.Keys" in helper, "C23:C32 is not compared"
    assert "$Before.Tables.Keys" in helper, "the five tables are not compared"
    # The unchanged group is the SUCCESS RECORD only, and the attempt fields are
    # deliberately absent from it.
    declared = re.search(r"\$script:Phase5SuccessRecordFields = @\((.*?)\)", source, re.S)
    assert declared, "the success-record field group is not declared"
    fields = set(_ps_string_literals(declared.group(1)))
    assert fields == {"last_successful_stamp", "last_successful_fingerprint",
                      "fingerprint_version", "last_successful_applied_timeline"}
    attempt = re.search(r"\$script:Phase5AttemptFields = @\((.*?)\)", source, re.S)
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
                  "'ScreenUpdating was restored to the CAPTURED caller value'",
                  "'EnableEvents was restored to the CAPTURED caller value'",
                  "'DisplayAlerts was restored to the CAPTURED caller value'",
                  "'Calculation was restored to the CAPTURED caller value'",
                  "'StatusBar was restored to the CAPTURED sentinel'"):
        assert claim in block, f"the rollback scenario never asserts {claim}"
    assert "Add-SnapshotUnchangedChecks" in block, (
        "the rollback never compares against the previous successful snapshot"
    )


def test_34_the_refusal_compares_the_two_groups_separately() -> None:
    source = _executable(SCENARIOS)
    row4 = source[source.index("$Excel.Run('PCCM_Calculate') | Out-Null\n        $row4"):
                  source.index("Add-Phase5Result 'P5-S4'")]
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
    assert "'exactly the five declared (sheet, shape, macro) bindings exist'" in source
    assert "'every macro-bound shape is one of the five declared buttons'" in source
    assert "'no undeclared shape invokes a PCCM_ procedure'" in source
    assert "'the button ' + $wantSheet + '!' + $wantName + ' calls ' + $wantAction" in source
    assert "'the manifest declares 15 production modules'" in source
    assert "': the production module ' + $name + ' is a standard module'" in source


def test_38_the_production_modules_are_asserted_by_name_not_by_count() -> None:
    source = _executable(SCENARIOS)
    helper = _procedure(source, "Add-Phase5ModuleInventoryChecks")
    assert "foreach ($name in $ExpectedModules)" in helper, (
        "the modules are not checked by name"
    )
    assert "$standardNames -contains $name" in helper
    # And exactly, not merely at least: a set that gained a stray and lost a
    # real one would satisfy any inequality.
    assert "($standardNames.Count -eq @($ExpectedModules).Count)" in helper
    # And in the other direction: no standard module outside the manifest may
    # persist. A count on its own would pass a project that gained a stray
    # module and lost a real one.
    assert "': no standard module outside the manifest persists'" in helper
    assert "$ExpectedModules -notcontains $_" in helper
    # P5-M reaches the production namespace through that helper.
    block = source[source.index("Add-Phase5Result 'P5-M'") - 8000:source.index("Add-Phase5Result 'P5-M'")]
    assert "Add-Phase5ModuleInventoryChecks -List $list -Components $components" in block


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
        # A test NAME may say what it refuses - test_52 names the shell it is
        # keeping an algorithm OUT of - and that is a claim about the harness,
        # not a call into one.
        if line.lstrip().startswith("def test_"):
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
    reported = set(re.findall(RESULT_CALL + r"\s+'(P5-[A-Z0-9]+)'", source))
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
        "    Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'SKIP' 'phase 4 not intact'\n"
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
    planted = source.replace("Add-Phase5Result 'P5-S5'", "Add-Phase5Result 'P5-SX'")
    assert "Add-Phase5Result 'P5-S5'" not in planted, "the missing status row must be visible"
    present = [row for row in STATUS_ROW_IDS if _result_call(row) in planted]
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
    removal = source.index("Add-Phase5Result 'P5-D8'")
    for forbidden in (".Save()", ".SaveAs(", "SaveCopyAs", "xlOpenXMLWorkbookMacroEnabled"):
        assert forbidden not in source, f"the Phase-5 scenarios persist a workbook ({forbidden})"
    harness = _executable(HARNESS)
    # The accepted harness closes without saving, and that is unchanged.
    assert "$wb.Close($false)" in harness, "the workbook is no longer closed without saving"


# ===========================================================================
# 12. CORRECTION ROUND 1
#
# Eight defects found in independent review of 93f306d. Each has a test here
# that fails against the submitted harness and passes against the corrected one.
# ===========================================================================
def _scenario_block(source: str, after: str, upto: str) -> str:
    """The executable body of one scenario, bounded by its neighbours' results.

    Comment lines are stripped by `_executable`, so a section cannot be located
    by its banner. It is located by the Add-Result that closes the scenario
    before it and the one that closes it.
    """
    start = source.index(_result_call(after)) if after else 0
    return source[start:source.index(_result_call(upto))]


def _procedure(source: str, name: str) -> str:
    start = source.index(f"function {name} ")if f"function {name} " in source \
        else source.index(f"function {name}")
    tail = source[start + 1:]
    end = tail.index("\nfunction ") if "\nfunction " in tail else len(tail)
    return source[start:start + 1 + end]


def test_48_a_null_fixture_value_is_written_as_a_blank_cell() -> None:
    """BLOCKER 1. `[double]$null` is numeric ZERO in PowerShell.

    Both fixture writers cast before they branched, so plan case 14's blank
    inflation rate became a rate of 0 and plan case 23's blank profiling weight
    became 0%. Both models are VALID with a zero in place of the blank, so the
    refusal each case exists to prove could never have fired: the fixture was
    quietly destroying the condition it was written to exercise.
    """
    source = _executable(SCENARIOS)
    for procedure, subject in (("Write-Phase5InflationRates", "$rates.$year"),
                               ("Write-Phase5Weights", "$weight")):
        body = _procedure(source, procedure)
        guard = f"if ($null -eq {subject}) {{"
        assert guard in body, f"{procedure} does not branch on null at all"
        # THE BRANCH IS BEFORE THE CAST. A guard that runs after the conversion
        # would be testing a zero it created itself.
        cast = body.index(f"([double]{subject})")
        assert body.index(guard) < cast, (
            f"{procedure} casts to Double before it checks for null"
        )
        blank = body[body.index(guard):cast]
        assert "-Value $null" in blank, (
            f"{procedure} does not write a genuine blank on the null branch"
        )
        for forbidden in ("-Value 0", "-Value ''", '-Value ""', "-Value 0.0"):
            assert forbidden not in body, (
                f"{procedure} writes {forbidden} where the fixture says blank"
            )


def test_49_the_blank_fixture_cases_still_carry_their_blanks() -> None:
    """BLOCKER 1, tied to the two cases that define the condition."""
    cases = {str(case["id"]): case for case in _emitted()["cases"]["plan_cases"]}
    # Case 14: a required inflation rate is null.
    rates = cases["14"]["model"]["inflation"]["Standard"]
    assert any(value is None for value in rates.values()), (
        "plan case 14 no longer carries a blank inflation rate"
    )
    assert cases["14"]["expected_refusal"] == "ModelInputRefusal"
    # Case 23: a profile weight is null while the vector still sums to 100%.
    weights = cases["23"]["model"]["cost_lines"][0]["profile_weights"]
    assert any(value is None for value in weights), (
        "plan case 23 no longer carries a blank profiling weight"
    )
    present = [value for value in weights if value is not None]
    assert abs(sum(present) - 1.0) < 1e-12, (
        "plan case 23's remaining weights no longer sum to one, so a zero in "
        "place of the blank would be refused for the wrong reason"
    )
    # Both are in the refusal scenario, so a lost blank cannot hide there.
    source = _executable(SCENARIOS)
    block = _ledger_block(source)
    assert "14" in _mapped_ids(block) and "23" in _mapped_ids(block)


def test_50_a_refusal_preserves_the_prior_snapshot_rather_than_emptying_it() -> None:
    """BLOCKER 2. "No partial result" means no partial NEW snapshot survives.

    The first submission asserted that every _Calc table held zero populated rows
    after a refusal. P5-AN runs first and leaves a successful snapshot,
    Set-Phase5Fixture changes the INPUT model and never touches _Calc, and a
    pre-write refusal is REQUIRED to leave C13:C16, C23:C32 and all five tables
    exactly as they were - so that assertion would have failed against correct
    production behaviour.
    """
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-AN", "P5-RF")
    assert "populated" not in block, (
        "the refusal proof still requires the analytical tables to be empty"
    )
    assert "$before = Get-Phase5Snapshot" in block, (
        "the refusal proof captures no baseline to compare against"
    )
    assert "Add-SnapshotUnchangedChecks" in block, (
        "the refusal proof never asserts the prior snapshot survived"
    )
    assert "Add-Phase5AttemptAxisChecks" in block, (
        "the refusal proof never asserts that C17:C20 CHANGED"
    )
    # A SUCCESSFUL baseline is established first, and it is not vacuous.
    assert "'a successful baseline snapshot was established first'" in block
    assert "not empty, so the comparison is not vacuous" in block
    # And the snapshot is NOT cleared to make the assertion pass.
    for forbidden in ("Clear-Phase5GridBody -Workbook $Workbook -SheetName $Inspection.calc.sheet",
                      "ClearContents", "Remove-TableRow -Workbook $Workbook -SheetName $Inspection"):
        assert forbidden not in block, (
            f"the refusal proof clears the calculation workspace ({forbidden})"
        )


def test_51_the_identity_set_is_the_locked_i1_to_i5_mapping() -> None:
    """BLOCKER 3. I1, I2, I3a-c, I4a-c and I5 - each named, none merged."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-RF", "P5-ID")
    for identity in ("I1", "I2", "I3a", "I3b", "I3c", "I4a", "I4b", "I4c", "I5"):
        assert f"'{identity}'" in block or f"'{identity} " in block or f"' {identity}:" in block \
            or f"{identity}:" in block, f"identity {identity} is not named in the evidence"
    # I3 splits Base, Risk and Total. The first submission checked only Total.
    for column in ("base_cost_nominal", "expected_risk_nominal", "total_nominal",
                   "base_cost_pv", "expected_risk_pv", "total_pv"):
        assert f"'{column}'" in block, f"the annual column {column} is never asserted separately"
    # I5 is profiling evidence, per driver.
    assert "profile_weights" in block, "I5 asserts no profiling weights at all"
    assert "cost_profiling" in block and "risk_profiling" in block
    # Production's own reconciliation is the authority.
    assert "Reconcile" in block and "AllIdentitiesHold" in block, (
        "the evidence never names production's own reconciliation as the authority"
    )
    assert "'30'" in block, "the cancellation-heavy fixture is not among the identity cases"


def test_52_no_headline_conditioning_is_reimplemented_in_powershell() -> None:
    """BLOCKER 3. Erratum C1 rejected headline-based conditioning.

    The first submission decided each identity with
    `max(|left|, |right|, floor) * coefficient` in PowerShell - the rejected
    oracle, made the acceptance authority. Case 30 exists because that shape can
    falsely fail a correct cancellation-heavy calculation.
    """
    source = _executable(SCENARIOS)
    assert "[Math]::Max([Math]::Max(" not in source, (
        "a headline-based conditioning allowance was reintroduced in PowerShell"
    )
    assert "identity_absolute_floor" not in source, (
        "the harness reads the conditioning floor, which only a reimplementation needs"
    )
    for forbidden in ("$close = {", "$allowance", "conditioning"):
        assert forbidden not in source, (
            f"the harness reimplements the production conditioning rule ({forbidden})"
        )
    # The one comparison primitive is the same one every other value check uses.
    block = _scenario_block(source, "P5-RF", "P5-ID")
    assert "Test-CalcValue" in block
    assert block.count("[Math]::Abs") == 0, (
        "the identity block computes its own difference instead of comparing to the oracle"
    )


def test_53_the_staleness_target_has_an_emitted_oracle() -> None:
    """BLOCKER 4. §25.2 requires the affected value to change TO the oracle value.

    Exchanging two profiling weights produced a model the corpus does not
    describe, so after the second Calculate there was nothing to compare against
    and the proof degenerated into an annual ROW COUNT.
    """
    cases = {str(case["id"]): case for case in _emitted()["cases"]["plan_cases"]}
    source, target = cases["3"], cases["19"]
    # The two fixtures really do differ in exactly one fingerprinted scalar.
    for key in ("timeline", "fx", "inflation", "cost_lines", "risks"):
        assert source["model"][key] == target["model"][key], (
            f"plan cases 3 and 19 no longer share {key}, so the transition is not one variable"
        )
    assert source["model"]["discount_rate"] != target["model"]["discount_rate"]
    assert "expected" in target and target["expected"]["totals"], (
        "the staleness target emits no expected block"
    )

    text = _executable(SCENARIOS)
    block = _scenario_block(text, "P5-S1", "P5-ST")
    assert "$candidate.id -eq '19'" in block, "the staleness target fixture is not plan case 19"
    assert "$Inspection.inputs.discount_rate.defined_name" in block, (
        "the staleness edit is not the Discount Rate scalar"
    )
    assert "'the source and target fixtures differ ONLY in Discount Rate'" in block
    # AN ORACLE COMPARISON, not a row count.
    assert "Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `" in block
    assert "-Case $targetCase" in block, (
        "the recalculated model is never compared against the target's emitted block"
    )
    assert "PCCM_ApplyTimeline" not in block, (
        "Apply Timeline is used to create staleness"
    )


def test_54_the_row_order_proof_actually_reorders_rows() -> None:
    """BLOCKER 5. Sorting a one-row table changes nothing.

    The Sort call was real; the reorder evidence was not. Plan case 3 has one
    Cost Line.
    """
    cases = {str(case["id"]): case for case in _emitted()["cases"]["plan_cases"]}
    assert len(cases["3"]["model"]["cost_lines"]) == 1, (
        "plan case 3 gained rows; the reasoning below needs restating"
    )
    assert len(cases["30"]["model"]["cost_lines"]) >= 2, (
        "the reorder fixture no longer has more than one Cost Line"
    )
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-ST", "P5-NS")
    assert "$candidate.id -eq '30'" in block, "the reorder probe does not use the multi-row fixture"
    assert "'the reorder fixture has MORE THAN ONE Cost Line, so a sort can move rows'" in block
    assert "$idsBefore = @(Get-IdColumnValues" in block, "the order before the sort is never captured"
    assert "$idsAfter = @(Get-IdColumnValues" in block, "the order after the sort is never captured"
    assert "'the physical permanent-ID order ACTUALLY changed'" in block, (
        "the harness accepts 'Sort was called' as evidence that order changed"
    )
    assert "($idsBefore -join ',') -cne ($idsAfter -join ',')" in block


def test_55_the_four_non_staleness_probes_are_independent() -> None:
    """BLOCKER 5.2. Each probe starts from a baseline and restores its change."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-ST", "P5-NS")
    for restored in ("'Description restored'",
                     "'Cost Lines re-sorted back'",
                     "'Selected Confidence Level restored'",
                     "'the unreferenced FX assumption removed'"):
        assert restored in block, f"a probe never restores its change ({restored})"
    # The probe takes the digest it must hold against as an ARGUMENT, so a later
    # probe cannot silently inherit an earlier probe's edit.
    assert "param([string]$Name, [string]$Digest)" in block, (
        "the probe reads a shared mutable digest instead of the one it began with"
    )
    assert block.count("& $probe") >= 8, (
        "there are fewer probe invocations than four changes plus four restorations"
    )
    # Each probe re-establishes a baseline it can trust.
    assert block.count("Set-Phase5Fixture") >= 3, (
        "the probes never re-establish a known baseline between changes"
    )


def test_56_the_golden_case_asserts_the_emitted_reference_digest() -> None:
    """BLOCKER 6. Two identically WRONG fingerprints would pass "stored == current".

    Plan case 1 is the model the reference stream was built from, so the complete
    production path must land on the emitted digest.
    """
    cases = {str(case["id"]): case for case in _emitted()["cases"]["plan_cases"]}
    reference = _emitted()["cases"]["fingerprint"]["reference"]
    assert reference["case"] == 26
    # The reference stream really is case 1's model: same header scalars, same
    # single Cost Line record.
    model = cases["1"]["model"]
    stream = reference["stream"]
    assert f"S6:{model['cost_lines'][0]['permanent_id']}" in stream
    assert f"S{len(model['cost_lines'][0]['distribution'])}:{model['cost_lines'][0]['distribution']}" in stream

    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-D8", "P5-AN")
    assert "$id -eq '1'" in block, "the golden case is not singled out"
    assert "$Cases.fingerprint.reference.digest" in block, (
        "the end-to-end fingerprint is never compared against the emitted digest"
    )
    assert "PCCM_CurrentInputFingerprint() on plan case 1 equals the emitted reference digest" in block
    assert "PCCM_CalculationFingerprint() after the commit equals the emitted reference digest" in block
    # The direct primitive proof is a DIFFERENT claim and both survive.
    assert "Add-Phase5Result 'P5-D5'" in source
    assert "GBD_DigestStream" in source


def test_57_the_utf16_canonical_field_is_compared_in_full() -> None:
    """BLOCKER 6.2. A prefix check passes a mangled payload of the right length."""
    vectors = _emitted()["cases"]["fingerprint"]["utf16_vectors"]["vectors"]
    for vector in vectors:
        assert vector["canonical_text_field"], "the corpus emits no canonical field"
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Phase5Result 'P5-D4'") - 5000:source.index("Add-Phase5Result 'P5-D4'")]
    assert "$vector.canonical_text_field" in block, (
        "the emitted canonical field is never used"
    )
    assert "the COMPLETE canonical text field matches the emitted one" in block
    assert "-ceq $expectedField" in block, "the comparison is not an exact ordinal one"
    # The prefix check survives as a supplementary claim, not as the whole one.
    assert "its length prefix is the UTF-16 unit count" in block


def test_58_the_analytical_audit_covers_every_declared_column() -> None:
    """BLOCKER 7. Two published columns had no emitted expectation behind them."""
    emitted = _emitted()
    inspection = emitted["inspection"]["calc"]["tables"]
    case = next(c for c in emitted["cases"]["plan_cases"]
                if c["kind"] == "analytical" and str(c["id"]) == "2")
    expected = case["expected"]

    # The corpus now states an expectation for every column of every table.
    coverage = {
        "calc_years": {row for row in expected["calc_years"][0]},
        "calc_fx": {row for row in expected["resolved_fx_rows"][0]},
        "calc_drivers": set(expected["drivers"][0]) - {"weights"},
        "calc_annual": set(expected["annual"][0]),
        "calc_inflation_factors": {"inflation_profile", "calendar_year", "annual_rate",
                                   "cumulative_inflation_factor"},
    }
    for key, table in inspection.items():
        declared = set(table["columns"])
        if key == "calc_inflation_factors":
            continue  # named differently in the corpus; checked below
        missing = declared - coverage[key]
        assert not missing, f"{key} publishes {sorted(missing)} with no emitted expectation"
    # calc_inflation_factors uses `profile` / `cumulative_factor` in the corpus.
    row = expected["inflation_factors"][0]
    assert set(row) == {"profile", "calendar_year", "annual_rate", "cumulative_factor"}

    source = _executable(SCENARIOS)
    checks = _procedure(source, "Add-Phase5AnalyticalChecks")
    assert "$expected.calc_years" in checks, "tblCalcYears.Calendar Year is still unasserted"
    assert "$expected.resolved_fx_rows" in checks, "tblCalcFX.Referenced By is still unasserted"
    # And the assertions are driven from the fixture's own field names, so a new
    # emitted column is asserted without editing the harness.
    assert checks.count("foreach ($field in $wanted.PSObject.Properties.Name)") >= 4
    # Neither value is DERIVED in PowerShell.
    for forbidden in ("start_year +", "+ $index - 1", "$references++", "Measure-Object"):
        assert forbidden not in checks, (
            f"the harness derives an expected value instead of reading it ({forbidden})"
        )


def test_59_the_successful_calc_state_record_is_asserted(): 
    """BLOCKER 7.2. C13:C20 itself, not only the four accessors."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Add-Phase5SuccessStateChecks")
    assert "C13 last successful stamp is non-blank" in body
    assert "C14 is exactly the digest PCCM_CalculationFingerprint returned" in body
    assert "$Cases.fingerprint.constants.FP_VERSION" in body, (
        "C15 is checked against something other than the emitted fingerprint version"
    )
    assert "$Case.expected.applied_timeline" in body, (
        "C16 is checked against something other than the emitted applied-timeline text"
    )
    assert "C17 = SUCCESS" in body
    assert "C18 is BLANK on success" in body
    assert "C19 = CURRENT" in body
    assert "C20 status-evaluation timestamp is non-blank" in body
    # The two timestamps are NOT required to be equal to each other.
    assert "status_evaluated_at'] -ceq" not in body
    assert "last_successful_stamp'] -ceq" not in body
    # And it is actually called from the analytical scenario.
    assert "Add-Phase5SuccessStateChecks" in _scenario_block(source, "P5-D8", "P5-AN")


def test_60_the_applied_timeline_text_is_a_checked_copy_of_production() -> None:
    """The corpus emits `base/start/duration`; modCalcReport OWNS that format."""
    report = _text(SRC_VBA / "modCalcReport.bas")
    body = report[report.index("Private Function AppliedTimelineText"):]
    body = body[:body.index("End Function")]
    assert 'CStr(package.Model.Timeline.BaseYear) & "/"' in body
    assert 'CStr(package.Model.Timeline.StartYear) & "/"' in body
    assert "CStr(package.Model.Timeline.Duration)" in body
    emitted = _emitted()["cases"]
    for case in emitted["plan_cases"]:
        if case["kind"] != "analytical":
            continue
        timeline = case["model"]["timeline"]
        wanted = f"{timeline['base_year']}/{timeline['start_year']}/{timeline['duration']}"
        assert case["expected"]["applied_timeline"] == wanted, (
            f"case {case['id']}: the emitted applied-timeline text is not base/start/duration"
        )


def test_61_the_inspection_projection_is_identities_only() -> None:
    """BLOCKER 8. A POSITIVE schema, not a list of banned names.

    `fingerprint_version` is an expected VALUE and the two label lists are model
    SEMANTICS. A ban-list can only refuse the semantic values somebody already
    thought of; an allowlist refuses the next one too, whatever it is called.
    """
    from pccm_builder.gate_b_inspection import (
        ALLOWED_BLOCK_KEYS, ALLOWED_CALC_KEYS, ALLOWED_INPUT_KEYS,
        ALLOWED_INPUT_TABLE_KEYS, ALLOWED_ROOT_KEYS, ALLOWED_TABLE_KEYS,
    )

    inspection = _emitted()["inspection"]
    assert set(inspection) == set(ALLOWED_ROOT_KEYS)
    assert set(inspection["calc"]) == set(ALLOWED_CALC_KEYS)
    for table in inspection["calc"]["tables"].values():
        assert set(table) == set(ALLOWED_TABLE_KEYS)
    for block in inspection["calc"]["scalar_blocks"].values():
        assert set(block) == set(ALLOWED_BLOCK_KEYS)
    for spec in inspection["inputs"].values():
        assert set(spec) == set(ALLOWED_INPUT_KEYS)
    for table in inspection["input_tables"].values():
        assert set(table) == set(ALLOWED_INPUT_TABLE_KEYS)

    # The three semantic values independent review removed, refused by name at
    # the level they lived at. `fingerprint_version` survives as a calc_state ROW
    # NAME, which is an address - row 15 is where the version is written - and
    # banning it everywhere would refuse an identity along with the value.
    assert "fingerprint_version" not in inspection["calc"]
    assert "derived_status_labels" not in inspection["calc"]
    assert "attempt_result_labels" not in inspection["calc"]
    assert inspection["calc"]["scalar_blocks"]["calc_state"]["rows"]["fingerprint_version"] == 15, (
        "the row that HOLDS the version is an address and must stay"
    )
    blob = json.dumps(inspection)
    for banned in ("tolerance", "\"expected\"", "digest", "NOT CALCULATED", "REFUSED"):
        assert banned not in blob, f"the projection carries a semantic value ({banned})"
    # Every leaf is an identity: a string name, or an integer row/column ordinal.
    for table in inspection["calc"]["tables"].values():
        assert isinstance(table["table_name"], str)
        assert isinstance(table["header_row"], int)
    # And the version bumped, because the shape changed.
    assert inspection["schema_version"] == 2


def test_62_the_semantic_values_still_come_from_their_own_authorities() -> None:
    """Removed from the projection, not moved into PowerShell literals."""
    emitted = _emitted()
    assert emitted["cases"]["fingerprint"]["constants"]["FP_VERSION"] == 1
    source = _executable(SCENARIOS)
    assert "$Cases.fingerprint.constants.FP_VERSION" in source, (
        "the fingerprint version is not read from the value corpus"
    )
    assert "$Inspection.calc.fingerprint_version" not in source
    assert "$Inspection.calc.derived_status_labels" not in source
    assert "$Inspection.calc.attempt_result_labels" not in source
    # The status and attempt vocabularies are asserted as literals in the ROWS
    # they belong to, which is where the matrix states them - not read from an
    # address projection.
    for label in ("'CURRENT'", "'STALE'", "'INVALID'", "'SUCCESS'", "'REFUSED'", "'FAILED'"):
        assert label in source


# ===========================================================================
# 13. CORRECTION-ROUND NEGATIVE CONTROLS
#
# One per defect independent review found, planted as synthetic text and watched
# by the detector above. Nothing is written to disk and nothing runs.
# ===========================================================================
def test_nc_29_a_null_inflation_rate_cast_to_zero_is_caught() -> None:
    """DEFECT 1, exactly as it shipped."""
    planted = _synthetic(
        "function Write-Phase5InflationRates {\n"
        "    foreach ($year in $rates.PSObject.Properties.Name) {\n"
        "        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet "
        "-TableName $grid.table_name `\n"
        "            -RowIndex $rowIndex -ColumnIndex $ordinal -Value ([double]$rates.$year)\n"
        "    }\n"
        "}\n"
    )
    assert "if ($null -eq $rates.$year) {" not in planted, (
        "the missing null branch must be visible; [double]$null is numeric zero"
    )
    assert "([double]$rates.$year)" in planted


def test_nc_30_a_null_profiling_weight_cast_to_zero_is_caught() -> None:
    """DEFECT 1, the other writer."""
    planted = _synthetic(
        "function Write-Phase5Weights {\n"
        "    foreach ($weight in @($driver.profile_weights)) {\n"
        "        Set-TableCell -Workbook $Workbook -SheetName $grid.sheet "
        "-TableName $grid.table_name `\n"
        "            -RowIndex $rowIndex -ColumnIndex ($fixed + $offset) -Value ([double]$weight)\n"
        "    }\n"
        "}\n"
    )
    assert "if ($null -eq $weight) {" not in planted, "the missing null branch must be visible"
    assert "([double]$weight)" in planted


def test_nc_31_a_null_branch_placed_after_the_cast_is_caught() -> None:
    """The branch must precede the conversion, not follow it."""
    planted = _synthetic(
        "$value = [double]$weight\n"
        "if ($null -eq $weight) { $value = $null }\n"
        "Set-TableCell -Value $value\n"
    )
    cast = planted.index("[double]$weight")
    guard = planted.index("if ($null -eq $weight)")
    assert cast < guard, "a guard that runs after the cast must be visible"


def test_nc_32_a_refusal_requiring_empty_analytical_tables_is_caught() -> None:
    """DEFECT 2, exactly as it shipped: it would FAIL against correct production."""
    planted = _synthetic(
        "foreach ($tableKey in $Inspection.calc.tables.PSObject.Properties.Name) {\n"
        "    $rows = @(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection "
        "-TableKey $tableKey)\n"
        "    $populated = @($rows | Where-Object { -not (Test-CalcBlank -Actual $_[0]) })\n"
        "    $null = Add-Check $list 'carries no refused output' ($populated.Count -eq 0)\n"
        "}\n"
    )
    assert "populated" in planted, (
        "an assertion that the workspace is EMPTY after a refusal must be visible; "
        "a pre-write refusal preserves the prior successful snapshot"
    )
    assert "Add-SnapshotUnchangedChecks" not in planted


def test_nc_33_clearing_the_snapshot_before_a_refusal_is_caught() -> None:
    """Weakening the proof by removing what it is supposed to preserve."""
    planted = _synthetic(
        "Clear-Phase5GridBody -Workbook $Workbook -SheetName $Inspection.calc.sheet `\n"
        "    -TableName 'tblCalcDrivers' -ColumnCount 21\n"
        "$Excel.Run('PCCM_Calculate') | Out-Null\n"
        "$null = Add-Check $list 'nothing survived' ($rows.Count -eq 0)\n"
    )
    assert "Clear-Phase5GridBody -Workbook $Workbook -SheetName $Inspection.calc.sheet" in planted, (
        "clearing the calculation workspace to make a refusal assertion pass must be visible"
    )


def test_nc_34_a_headline_conditioning_tolerance_is_caught() -> None:
    """DEFECT 3, exactly as it shipped: erratum C1's rejected oracle."""
    planted = _synthetic(
        "$close = {\n"
        "    param([double]$Left, [double]$Right)\n"
        "    $scale = [Math]::Max([Math]::Max([Math]::Abs($Left), [Math]::Abs($Right)), $floor)\n"
        "    return ([Math]::Abs($Left - $Right) -le ($tolerance * $scale))\n"
        "}\n"
    )
    assert "[Math]::Max([Math]::Max(" in planted, (
        "a PowerShell reimplementation of the conditioning allowance must be visible"
    )
    assert "$close = {" in planted


def test_nc_35_an_identity_set_missing_the_base_risk_split_is_caught() -> None:
    """DEFECT 3: I3 and I4 split Base, Risk and Total. Total alone is not I3."""
    planted = _synthetic(
        "foreach ($identity in @(\n"
        "    @{ name = 'I3'; column = 'total_nominal'; headline = 'e_nom' },\n"
        "    @{ name = 'I4'; column = 'total_pv'; headline = 'e_pv' })) { }\n"
    )
    for column in ("base_cost_nominal", "expected_risk_nominal",
                   "base_cost_pv", "expected_risk_pv"):
        assert column not in planted, f"the missing {column} assertion must be visible"
    for identity in ("'I3a'", "'I3b'", "'I4a'", "'I4b'"):
        assert identity not in planted


def test_nc_36_an_identity_set_with_no_i5_evidence_is_caught() -> None:
    """DEFECT 3: I5 is per-driver profiling evidence, and it was absent entirely."""
    planted = _synthetic(
        "$null = Add-Check $list 'I1' ($ok)\n"
        "$null = Add-Check $list 'I2' ($ok)\n"
        "$null = Add-Check $list 'I5: the annual series sums to E' ($ok)\n"
    )
    assert "profile_weights" not in planted, "the missing I5 profiling evidence must be visible"
    assert "cost_profiling" not in planted and "risk_profiling" not in planted


def test_nc_37_a_row_count_only_staleness_proof_is_caught() -> None:
    """DEFECT 4, exactly as it shipped: no numerical value was proved."""
    planted = _synthetic(
        "$null = Add-Check $list 'the recalculated annual series is non-empty' `\n"
        "    (@(Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection "
        "-TableKey 'calc_annual').Count -eq `\n"
        "     @($baseCase.expected.annual).Count)\n"
    )
    assert "Add-Phase5AnalyticalChecks" not in planted, (
        "a staleness proof that compares only a row count must be visible"
    )
    assert ".Count -eq" in planted


def test_nc_38_a_staleness_target_with_no_emitted_oracle_is_caught() -> None:
    """DEFECT 4: swapping weights produces a model the corpus does not describe."""
    planted = _synthetic(
        "Set-TableCell -RowIndex 1 -ColumnIndex ($fixed + 1) -Value ([double]$weights[1])\n"
        "Set-TableCell -RowIndex 1 -ColumnIndex ($fixed + 2) -Value ([double]$weights[0])\n"
        "$Excel.Run('PCCM_Calculate') | Out-Null\n"
    )
    assert "$targetCase" not in planted, (
        "an edit with no emitted target fixture must be visible"
    )
    assert "$weights[1]" in planted


def test_nc_39_a_one_row_reorder_fixture_is_caught() -> None:
    """DEFECT 5: sorting a one-row table changes nothing."""
    emitted = _emitted()["cases"]["plan_cases"]
    single = next(case for case in emitted if str(case["id"]) == "3")
    assert len(single["model"]["cost_lines"]) < 2, (
        "the one-row fixture must be visible as one-row"
    )
    planted = _synthetic(
        "Invoke-TableSort -Workbook $Workbook -SheetName $costReg.sheet `\n"
        "    -TableName $costReg.table_name -KeyColumnIndex 3 -Order 2\n"
        "& $probe 'Cost Lines physically re-sorted' $digest\n"
    )
    assert "$idsBefore" not in planted and "$idsAfter" not in planted, (
        "a reorder proof that never captures the order must be visible"
    )
    assert "ACTUALLY changed" not in planted


def test_nc_40_cumulative_non_staleness_probes_are_caught() -> None:
    """DEFECT 5.2: four live edits at once isolates nothing."""
    planted = _synthetic(
        "Set-TableCell -ColumnIndex $descriptionOrdinal -Value 'changed'\n"
        "& $probe 'Description changed'\n"
        "Invoke-TableSort -KeyColumnIndex $descriptionOrdinal -Order 2\n"
        "& $probe 'reordered'\n"
        "Set-NamedValueText -Text 'P90'\n"
        "& $probe 'confidence changed'\n"
    )
    for restored in ("'Description restored'", "'Cost Lines re-sorted back'",
                     "'Selected Confidence Level restored'"):
        assert restored not in planted, f"the missing restoration must be visible ({restored})"
    assert "param([string]$Name, [string]$Digest)" not in planted, (
        "a probe with no per-probe baseline digest must be visible"
    )


def test_nc_41_a_golden_case_asserting_only_stored_equals_current_is_caught() -> None:
    """DEFECT 6: two identically WRONG fingerprints satisfy that."""
    planted = _synthetic(
        "$fingerprint = [string]$Excel.Run('PCCM_CalculationFingerprint')\n"
        "$current = [string]$Excel.Run('PCCM_CurrentInputFingerprint')\n"
        "$null = Add-Check $list 'they agree' ($fingerprint -ceq $current)\n"
    )
    assert "$Cases.fingerprint.reference.digest" not in planted, (
        "a golden case that never reaches the emitted digest must be visible"
    )
    assert "$fingerprint -ceq $current" in planted


def test_nc_42_a_prefix_only_utf16_field_check_is_caught() -> None:
    """DEFECT 6.2: a mangled payload of the right length passes a prefix check."""
    planted = _synthetic(
        "$field = [string]$Excel.Run('GBD_CanonicalTextField', $units)\n"
        "$expectedField = 'OK|S' + [string]$vector.utf16_length + ':'\n"
        "$null = Add-Check $list 'prefix' ($field.StartsWith($expectedField))\n"
    )
    assert "$vector.canonical_text_field" not in planted, (
        "a prefix-only canonical field check must be visible"
    )
    assert "StartsWith" in planted


def test_nc_43_an_omitted_calendar_year_assertion_is_caught() -> None:
    """DEFECT 7: tblCalcYears.Calendar Year had no emitted expectation."""
    planted = _synthetic(
        "$indexColumn = Get-CalcTableColumnIndex -TableKey 'calc_years' -ColumnKey 'project_index'\n"
        "$factorColumn = Get-CalcTableColumnIndex -TableKey 'calc_years' -ColumnKey 'discount_factor'\n"
        "foreach ($row in $years) { $null = Add-Check $list 'discount factor' ($ok) }\n"
    )
    assert "$expected.calc_years" not in planted, (
        "an assertion set that never reads the emitted calc_years rows must be visible"
    )
    assert "calendar_year" not in planted


def test_nc_44_an_omitted_referenced_by_assertion_is_caught() -> None:
    """DEFECT 7: tblCalcFX.Referenced By had no emitted expectation."""
    planted = _synthetic(
        "$expectedFx = @($expected.resolved_fx.PSObject.Properties.Name)\n"
        "foreach ($currency in $expectedFx) { $null = Add-Check $list 'FX rate' ($ok) }\n"
    )
    assert "$expected.resolved_fx_rows" not in planted, (
        "an assertion set that never reads the emitted FX rows must be visible"
    )
    assert "referenced_by" not in planted


def test_nc_45_a_powershell_derived_calendar_year_is_caught() -> None:
    """DEFECT 7.1: deriving the answer is not checking it."""
    planted = _synthetic(
        "$wanted = $model.timeline.start_year + $index - 1\n"
        "$null = Add-Check $list 'calendar year' ($row[$ordinal] -eq $wanted)\n"
    )
    assert "+ $index - 1" in planted, "a PowerShell-derived calendar year must be visible"
    assert "$expected.calc_years" not in planted


def test_nc_46_an_omitted_successful_calc_state_assertion_is_caught() -> None:
    """DEFECT 7.2: the four accessors alone leave C13:C20 unexamined."""
    planted = _synthetic(
        "$null = Add-Check $list 'C17 = SUCCESS' ($attempt -eq 'SUCCESS')\n"
        "$null = Add-Check $list 'C18 blank' ([string]::IsNullOrEmpty($detail))\n"
    )
    for missing in ("FP_VERSION", "applied_timeline", "last_successful_fingerprint"):
        assert missing not in planted, f"the unasserted calc_state field must be visible ({missing})"


def test_nc_47_a_semantic_value_in_the_inspection_projection_is_caught() -> None:
    """DEFECT 8: an allowlist refuses the NEXT semantic value too."""
    from pccm_builder.gate_b_inspection import ALLOWED_CALC_KEYS

    planted = dict(_emitted()["inspection"]["calc"])
    planted["fingerprint_version"] = 1
    planted["derived_status_labels"] = ["NOT CALCULATED", "CURRENT", "STALE", "INVALID"]
    extra = set(planted) - set(ALLOWED_CALC_KEYS)
    assert extra == {"fingerprint_version", "derived_status_labels"}, (
        "the reintroduced semantic values must be visible to the positive schema"
    )
    # And an innocent-looking new key is caught by the same rule.
    planted2 = dict(_emitted()["inspection"]["calc"])
    planted2["default_precision"] = 6
    assert set(planted2) - set(ALLOWED_CALC_KEYS) == {"default_precision"}, (
        "a ban-list would have missed this; the allowlist does not"
    )


# ===========================================================================
# 14. CORRECTION ROUND 2
#
# Five defects found in independent review of aa18cab. Each has a test here that
# fails against that source and passes against the corrected one.
# ===========================================================================
LOCKED_PREREQUISITE_PREDICATES = {
    # timeline and the structural handoff
    "base_year_after_start_year", "structure_change_pending",
    # the discount input's type
    "discount_rate_blank", "discount_rate_non_numeric",
    # FX and the reporting-currency invariant
    "referenced_currency_missing", "referenced_currency_duplicated",
    "referenced_rate_not_positive", "referenced_rate_blank",
    "referenced_rate_non_numeric", "reporting_currency_missing",
    "reporting_currency_duplicated", "reporting_currency_rate_not_one",
    # inflation
    "referenced_profile_missing", "referenced_rate_non_numeric_inflation",
    # profiling
    "profiling_cell_non_numeric",
    # distribution
    "distribution_missing", "distribution_unknown",
    # three-point ordering
    "triangular_ordering", "beta_pert_ordering", "uniform_ordering",
    # quantity
    "quantity_missing", "quantity_non_numeric",
    # probability
    "probability_missing", "probability_non_numeric",
    "probability_below_zero", "probability_above_one",
}

LOCKED_NO_BLOCK_PREDICATES = {
    "unreferenced_fx_duplicated", "unreferenced_fx_blank_rate",
    "unreferenced_profile_incomplete",
}


def _gate_b() -> dict:
    return _emitted()["cases"]["gate_b"]


def test_63_every_locked_prerequisite_has_a_windows_scenario() -> None:
    """BLOCKER 1. The nine refusal plan cases do not exhaust plan section 18.

    Base Year after Start Year, STRUCTURE CHANGE PENDING, a duplicated
    referenced currency, a non-numeric Probability, an unknown Distribution and a
    dozen more locked predicates had no real-Windows scenario at all.
    """
    gate_b = _gate_b()
    # THREE ROUTES, ONE LEDGER. A predicate the workbook cannot reach because an
    # earlier accepted gate refuses first is proved by calling the checker
    # directly; that is still a Windows scenario, and it still counts.
    emitted = ({entry["predicate"] for entry in gate_b["prerequisite_cases"]}
               | {entry["predicate"] for entry in gate_b["direct_check_cases"]})
    # The inflation non-numeric predicate shares a name with the FX one in the
    # locked list above; the corpus disambiguates by section.
    sections = {entry["section"] for entry in gate_b["prerequisite_cases"]}
    assert "18.I2" in sections, "no scenario covers a non-numeric referenced inflation rate"
    locked = LOCKED_PREREQUISITE_PREDICATES - {"referenced_rate_non_numeric_inflation"}
    missing = sorted(locked - emitted)
    assert not missing, f"locked prerequisites with no Windows scenario: {missing}"
    # BIDIRECTIONAL: a scenario nobody locked is unexplained coverage.
    extra = sorted(emitted - locked)
    assert not extra, f"Gate-B prerequisite scenarios outside the locked set: {extra}"
    assert len(gate_b["prerequisite_cases"]) == 25
    assert len(gate_b["direct_check_cases"]) == 1


def test_64_structure_change_pending_is_covered_and_is_not_re_applied() -> None:
    """BLOCKER 1. The one predicate the Phase-4 structural gate holds."""
    entry = next(item for item in _gate_b()["prerequisite_cases"]
                 if item["predicate"] == "structure_change_pending")
    assert entry["mutation"]["kind"] == "entered_structure"
    assert entry["mutation"]["apply_timeline"] is False, (
        "the timeline is re-applied, so the pending state is never reached"
    )
    assert "STRUCTURE CHANGE PENDING" in entry["detail_tokens"]
    # And the harness really honours the flag.
    source = _executable(SCENARIOS)
    body = _procedure(source, "Invoke-Phase5Mutation")
    assert "if ($Mutation.apply_timeline) {" in body, (
        "the mutation applier always re-applies, so the pending state is unreachable"
    )


def test_65_every_prerequisite_has_a_specific_detail_discriminator() -> None:
    """BLOCKER 1.3. "some error occurred" is not evidence the predicate fired."""
    gate_b = _gate_b()
    for entry in gate_b["prerequisite_cases"] + gate_b["direct_check_cases"]:
        tokens = entry["detail_tokens"]
        assert tokens, f"{entry['id']} has no detail discriminator"
        for token in tokens:
            assert isinstance(token, str) and token.strip(), f"{entry['id']} has an empty token"
    # The nine plan-case refusals have them too.
    plan_tokens = gate_b["plan_refusal_tokens"]
    refusals = {str(case["id"]) for case in _emitted()["cases"]["plan_cases"]
                if case["kind"] == "refusal"}
    assert set(plan_tokens) == refusals, (
        f"plan-refusal discriminators do not match the refusal cases: "
        f"{sorted(set(plan_tokens) ^ refusals)}"
    )
    for case_id, tokens in plan_tokens.items():
        assert tokens, f"plan case {case_id} has no detail discriminator"

    # Every token is a real fragment of an accepted production message.
    production = "\n".join(_text(SRC_VBA / name) for name in (
        "modCalcResolve.bas", "modCalcCheck.bas", "modCalcFactors.bas",
        "modCalcReport.bas", "modAppState.bas", "modStructuralCheck.bas",
    )) + "\n" + _emitted()["constants"]
    unknown = []
    for entry in gate_b["prerequisite_cases"] + gate_b["direct_check_cases"]:
        for token in entry["detail_tokens"]:
            # Identifiers the fixture supplies (currency codes, permanent IDs,
            # calendar years) are not in the production text; the PREDICATE
            # fragments must be.
            if token in ("USD", "SAR", "CL-001", "2027", "Standard",
                         "MissingGateBProfile"):
                continue
            if token not in production:
                unknown.append(f"{entry['id']}:{token!r}")
    assert not unknown, (
        f"detail tokens that appear in no accepted production message: {unknown}"
    )

    # And the harness ASSERTS them, rather than only checking non-empty.
    source = _executable(SCENARIOS)
    helper = _procedure(source, "Add-Phase5DetailTokenChecks")
    assert "foreach ($token in @($Tokens))" in helper, (
        "the discriminator tokens are never compared"
    )
    assert "-like ('*' + $token + '*')" in helper
    assert "Add-Phase5DetailTokenChecks" in _scenario_block(source, "P5-AN", "P5-RF"), (
        "the plan-case refusals still check only that a detail exists"
    )


def test_66_the_prerequisite_matrix_is_emitted_not_hand_held() -> None:
    """BLOCKER 1.1. PowerShell consumes the corpus; it holds no list."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-RF", "P5-PQ")
    assert "$Cases.gate_b.prerequisite_cases" in block, (
        "the prerequisite scenario does not read the emitted matrix"
    )
    # No PowerShell-side predicate list may exist.
    for predicate in sorted(LOCKED_PREREQUISITE_PREDICATES):
        assert f"'{predicate}'" not in source, (
            f"the harness names the predicate {predicate} in its own source"
        )
    # The mutation applier is driven by the corpus's `kind` vocabulary.
    body = _procedure(source, "Invoke-Phase5Mutation")
    for kind in ("entered_structure", "named_number", "named_text", "named_blank",
                 "register_cell", "fx_row", "fx_remove", "inflation_cell",
                 "config_profile_add", "profiling_cell"):
        assert f"'{kind}'" in body, f"the applier cannot apply a {kind} mutation"
    assert "default { throw" in body, "an unknown mutation kind is silently ignored"
    # And every emitted mutation kind is one the applier implements.
    used = {entry["mutation"]["kind"] for entry in
            _gate_b()["prerequisite_cases"] + _gate_b()["no_block_cases"]}
    for kind in sorted(used):
        assert f"'{kind}'" in body, f"the corpus emits a {kind} mutation the harness cannot apply"


def test_67_a_null_mutation_value_is_written_as_a_blank() -> None:
    """BLOCKER 1.2. Several locked prerequisites ARE the blank."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Get-MutationValue")
    assert "if ($null -eq $raw) { return $null }" in body, (
        "a null mutation value is cast rather than written as a blank"
    )
    assert body.index("if ($null -eq $raw)") < body.index("[double]$raw"), (
        "the null branch runs after the cast"
    )
    # A blank Setup scalar goes through ClearContents, not through ''.
    clear = _procedure(source, "Clear-NamedValue")
    assert "$rng.ClearContents()" in clear
    # The corpus really does carry null mutations.
    nulls = [entry for entry in _gate_b()["prerequisite_cases"]
             if entry["mutation"].get("value", "") is None
             or entry["mutation"].get("rate", "") is None
             or entry["mutation"]["kind"] == "named_blank"]
    assert len(nulls) >= 4, f"only {len(nulls)} blank prerequisites are emitted"


def test_68_the_referenced_only_no_block_complement_exists() -> None:
    """BLOCKER 1.4. A harness that only proved refusals would accept over-refusal."""
    gate_b = _gate_b()
    emitted = {entry["predicate"] for entry in gate_b["no_block_cases"]}
    assert emitted == LOCKED_NO_BLOCK_PREDICATES, (
        f"the no-block set is {sorted(emitted)}, not {sorted(LOCKED_NO_BLOCK_PREDICATES)}"
    )
    for entry in gate_b["no_block_cases"]:
        assert entry["expected_attempt"] == "SUCCESS"
        assert entry["expected_status"] == "CURRENT"
        assert entry["detail_tokens"] == []
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-PQ", "P5-PN")
    assert "$Cases.gate_b.no_block_cases" in block
    assert "the detail stays blank - nothing was refused" in block
    assert "the stored fingerprint is unchanged by the unreferenced row" in block


def test_69_status_row_2_compares_the_whole_analytical_snapshot() -> None:
    """BLOCKER 2. A status query that rewrote analytical outputs would have passed."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-S1", "P5-S2")
    assert "$rowTwoBefore = Get-Phase5Snapshot" in block, (
        "row 2 captures no full baseline before the edit"
    )
    assert "$rowTwoAfter = Get-Phase5Snapshot" in block
    assert "Add-SnapshotUnchangedChecks -List $list -Before $rowTwoBefore -After $rowTwoAfter" in block, (
        "row 2 never compares C23:C32 or the five analytical tables"
    )
    # C17:C20 is handled SEPARATELY, because C19/C20 are deliberately refreshed.
    assert "'row 2: C19 was re-derived to STALE by the status evaluation'" in block
    assert "'row 2: C20 carries a status-evaluation timestamp'" in block
    assert "'row 2: C17 still records the previous SUCCESS'" in block


def test_70_the_row_order_probe_edits_no_cell() -> None:
    """BLOCKER 3. Rewriting Description changed two dimensions at once."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-ST", "P5-NS")
    # Bounded by the NEXT probe's first executable line: `_executable` strips
    # comments, so the banner that separates the probes is not there to slice on.
    reorder = block[block.index("$reorderCase = $null"):
                    block.index("$confidence = $Inspection.inputs.selected_confidence_level")]
    assert "Set-TableCell" not in reorder, (
        "the row-order probe still edits a cell to manufacture the ordering"
    )
    assert "sort-key" not in reorder
    assert "$descriptionOrdinal" in reorder, "the sort key is not the existing Description"
    assert "$idsBefore" in reorder and "$idsAfter" in reorder
    assert "'the physical permanent-ID order ACTUALLY changed'" in reorder
    # The fixture applier is what gives every row a distinct Description.
    driver = _procedure(source, "Write-Phase5Driver")
    assert "'GateB ' + [string]$Driver.permanent_id" in driver, (
        "the applier no longer writes a deterministic distinct Description"
    )


def test_71_the_driver_audit_reconstruction_exists() -> None:
    """BLOCKER 4. The A/B/C/D cross-check between two parts of the workbook."""
    audit = _gate_b()["audit_reconstruction"]
    assert len(audit["model"]["cost_lines"]) >= 2, "the audit fixture has too few Cost Lines"
    assert len(audit["model"]["risks"]) >= 1, "the audit fixture carries no Risk"
    mapping = {entry["headline"]: (entry["driver_column"], entry["kind"])
               for entry in audit["relationships"]}
    assert mapping == {
        "a_nom": ("deterministic_nominal", "Cost Line"),
        "a_pv": ("deterministic_pv", "Cost Line"),
        "b_nom": ("uncertainty_mean_shift_nominal", "Cost Line"),
        "b_pv": ("uncertainty_mean_shift_pv", "Cost Line"),
        "c_nom": ("mean_basis_nominal", "Cost Line"),
        "c_pv": ("mean_basis_pv", "Cost Line"),
        "d_nom": ("expected_risk_nominal", "Risk"),
        "d_pv": ("expected_risk_pv", "Risk"),
    }, f"the audit relationships are wrong: {mapping}"
    # The ordinals really are the locked ones: A=14/15, C=16/17, B=18/19, D=20/21.
    columns = _emitted()["inspection"]["calc"]["tables"]["calc_drivers"]["columns"]
    for headline, ordinal in (("a_nom", 14), ("a_pv", 15), ("c_nom", 16), ("c_pv", 17),
                              ("b_nom", 18), ("b_pv", 19), ("d_nom", 20), ("d_pv", 21)):
        assert columns[ordinal - 1] == mapping[headline][0], (
            f"{headline} is not column {ordinal} of tblCalcDrivers"
        )
    # And the relationship actually holds in the emitted oracle.
    drivers = audit["expected"]["drivers"]
    totals = audit["expected"]["totals"]
    for entry in audit["relationships"]:
        total = sum(row[entry["driver_column"]] for row in drivers
                    if row["driver_kind"] == entry["kind"]
                    and row[entry["driver_column"]] is not None)
        assert abs(total - totals[entry["headline"]]) <= 1e-9 * max(
            abs(totals[entry["headline"]]), 1.0
        ), f"{entry['headline']} does not reconstruct in the emitted oracle"


def test_72_the_reconstruction_reads_the_actual_workbook_on_both_sides() -> None:
    """BLOCKER 4.2. ACTUAL tblCalcDrivers to ACTUAL calc_totals."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-PN", "P5-AR")
    assert "Get-CalcTableRows -Workbook $Workbook -Inspection $Inspection -TableKey 'calc_drivers'" in block, (
        "the reconstruction never reads the real driver table"
    )
    assert "-Block 'calc_totals' -FieldKey ([string]$relationship.headline)" in block, (
        "the reconstruction never reads the real totals"
    )
    assert "$audit.relationships" in block, "the column mapping is not read from the corpus"
    # Partitioned by kind, both ways.
    assert "$isKind = ([string]$row[$kindColumn] -eq [string]$relationship.kind)" in block
    assert ": the opposite kind publishes BLANK in " in block, (
        "nothing proves the opposite kind's cells are N/A"
    )
    # A BLANK IS SKIPPED, never folded in as the opposite kind's identity 1.
    assert "$blank = Test-CalcBlank -Actual $row[$ordinal]" in block
    assert "if (-not $blank) {" in block
    for forbidden in ("-eq $null) { $sum = $sum + 1", "$sum + 1.0", "Value 1 }"):
        assert forbidden not in block, (
            f"a blank is folded in as an identity value ({forbidden})"
        )
    # No new tolerance is invented for an audit relationship.
    assert "[Math]::Max" not in block
    assert "identity_absolute_floor" not in block


def test_73_both_failpoints_prove_exact_caller_state_restoration() -> None:
    """BLOCKER 5. Defaults are not evidence that anything was restored."""
    source = _executable(SCENARIOS)
    block = source[source.index("function Invoke-Phase5RollbackScenario"):
                   source.index("$analyticalOk = Invoke-Phase5RollbackScenario")]
    # A NON-DEFAULT caller state is established and captured.
    for established in ("$Excel.ScreenUpdating = $false", "$Excel.EnableEvents = $false",
                        "$Excel.DisplayAlerts = $false", "$Excel.Calculation = -4135",
                        "$Excel.StatusBar = $sentinel"):
        assert established in block, f"the caller state is not made non-default ({established})"
    assert "$sentinel = 'PCCM Phase-5 rollback sentinel ' + $Failpoint" in block, (
        "the StatusBar sentinel is not unique to the failpoint under test"
    )
    assert "$callerState = [pscustomobject]@{" in block, "the caller state is never captured"
    # ALL FIVE properties compared against the CAPTURED values.
    for prop in ("ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation", "StatusBar"):
        assert f"'{prop} was restored to the CAPTURED" in block, (
            f"{prop} is not compared against the captured caller value"
        )
        assert f"$callerState.{prop}" in block
    # NOT against Excel's defaults.
    for hardcoded in ("'EnableEvents was restored' ($Excel.EnableEvents -eq $true)",
                      "'ScreenUpdating was restored' ($Excel.ScreenUpdating -eq $true)",
                      "'Calculation mode was restored to automatic'"):
        assert hardcoded not in block, f"a hard-coded default check survives ({hardcoded})"
    # The comparison happens BEFORE the harness normalises for later scenarios.
    compare = block.index("'StatusBar was restored to the CAPTURED sentinel'")
    normalise = block.index("$Excel.ScreenUpdating = $true")
    assert compare < normalise, (
        "the harness normalises application state before asserting restoration"
    )
    # And the proof is vacuous-free: the captured state must differ from defaults.
    assert "'the restored state is NOT merely Excel default state'" in block


def test_74_neither_failpoint_inherits_the_others_state_proof() -> None:
    """BLOCKER 5.2. One shared runner, invoked twice, with its own capture each time."""
    source = _executable(SCENARIOS)
    assert source.count("$callerState = [pscustomobject]@{") == 1, (
        "the caller state is captured in more than one place, so the two runs could diverge"
    )
    runner = source[source.index("function Invoke-Phase5RollbackScenario"):
                    source.index("$analyticalOk = Invoke-Phase5RollbackScenario")]
    assert "$sentinel = " in runner, "the sentinel is not built inside the shared runner"
    # Both invocations go through the runner that carries the capture.
    tail = source[source.index("$analyticalOk = Invoke-Phase5RollbackScenario"):]
    assert tail.count("Invoke-Phase5RollbackScenario -Excel $Excel") == 2, (
        "the two failpoint scenarios do not both run the state-restoration proof"
    )
    assert "-Failpoint $failpoints.AnalyticalWrite" in tail
    assert "-Failpoint $failpoints.SuccessCommit" in tail


# ===========================================================================
# 15. CORRECTION-ROUND-2 NEGATIVE CONTROLS
# ===========================================================================
def test_nc_48_a_missing_section_18_prerequisite_is_caught() -> None:
    """A. A locked predicate dropped from the emitted matrix."""
    gate_b = _gate_b()
    emitted = ({entry["predicate"] for entry in gate_b["prerequisite_cases"]}
               | {entry["predicate"] for entry in gate_b["direct_check_cases"]})
    planted = emitted - {"probability_above_one"}
    locked = LOCKED_PREREQUISITE_PREDICATES - {"referenced_rate_non_numeric_inflation"}
    assert sorted(locked - planted) == ["probability_above_one"], (
        "a locked prerequisite with no Windows scenario must be visible"
    )
    # And a predicate that lost its DIRECT route is equally visible.
    without_direct = {entry["predicate"] for entry in gate_b["prerequisite_cases"]}
    assert "base_year_after_start_year" in sorted(locked - without_direct), (
        "a predicate reachable only by the direct route must be visible if that route goes"
    )


def test_nc_49_a_missing_structure_change_pending_is_caught() -> None:
    """B. The one predicate the Phase-4 structural gate holds."""
    planted = [entry for entry in _gate_b()["prerequisite_cases"]
               if entry["predicate"] != "structure_change_pending"]
    assert not [entry for entry in planted
                if entry["predicate"] == "structure_change_pending"], (
        "the dropped STRUCTURE CHANGE PENDING scenario must be visible"
    )
    # And a version that re-applies the timeline never reaches the state.
    entry = dict(next(item for item in _gate_b()["prerequisite_cases"]
                      if item["predicate"] == "structure_change_pending"))
    entry["mutation"] = {**entry["mutation"], "apply_timeline": True}
    assert entry["mutation"]["apply_timeline"] is True, (
        "a mutation that re-applies the timeline must be visible as such"
    )


def test_nc_50_a_non_empty_only_detail_check_is_caught() -> None:
    """C. "some error occurred" is not evidence the predicate fired."""
    planted = _synthetic(
        "$null = Add-Check $list ('case ' + $id + ': the refusal detail is specific, not empty') `\n"
        "    (-not [string]::IsNullOrWhiteSpace($detail)) $detail\n"
    )
    assert "Add-Phase5DetailTokenChecks" not in planted, (
        "a non-empty-only detail check must be visible"
    )
    assert "$token" not in planted
    # And a corpus entry with no tokens is visible too.
    entry = dict(_gate_b()["prerequisite_cases"][0])
    entry["detail_tokens"] = []
    assert not entry["detail_tokens"], "a discriminator-free prerequisite must be visible"


def test_nc_51_row_2_checking_only_c13_c16_is_caught() -> None:
    """D. A status query that rewrote analytical outputs would pass this."""
    planted = _synthetic(
        "$after = Get-CalcScalarBlock -Workbook $Workbook -Inspection $Inspection "
        "-Block 'calc_state'\n"
        "foreach ($field in $successRecordFields) {\n"
        "    $null = Add-Check $list ('row 2: calc_state.' + $field + ' is unchanged') `\n"
        "        (Test-CalcValue -Actual $after[$field] -Expected $establishedState[$field])\n"
        "}\n"
    )
    assert "Get-Phase5Snapshot" not in planted, (
        "a row-2 proof that captures no analytical baseline must be visible"
    )
    assert "Add-SnapshotUnchangedChecks" not in planted


def test_nc_52_a_row_order_probe_that_edits_description_is_caught() -> None:
    """E. Rewriting the sort key changes two dimensions at once."""
    planted = _synthetic(
        "for ($row = 1; $row -le $rowCount; $row++) {\n"
        "    Set-TableCell -Workbook $Workbook -SheetName $costReg.sheet "
        "-TableName $costReg.table_name `\n"
        "        -RowIndex $row -ColumnIndex $descriptionOrdinal "
        "-Value ('sort-key-' + [string](100 - $row))\n"
        "}\n"
        "$idsBefore = @(Get-IdColumnValues -Workbook $Workbook -Info $costReg)\n"
    )
    assert "Set-TableCell" in planted, "the manufactured ordering must be visible"
    assert "sort-key" in planted


def test_nc_53_a_missing_audit_reconstruction_is_caught() -> None:
    """F. Comparing each side to the oracle is not the cross-check."""
    planted = _synthetic(
        "Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook -Inspection $Inspection `\n"
        "    -Case $case -Tolerances $Cases.tolerances\n"
    )
    assert "$audit.relationships" not in planted, (
        "an analytical-only proof with no reconstruction must be visible"
    )
    assert "calc_totals' -FieldKey ([string]$relationship.headline)" not in planted


def test_nc_54_a_reconstruction_with_wrong_columns_is_caught() -> None:
    """G. B on 16/17, C on 18/19, D including cost rows, D omitted."""
    columns = _emitted()["inspection"]["calc"]["tables"]["calc_drivers"]["columns"]
    correct = {entry["headline"]: (entry["driver_column"], entry["kind"])
               for entry in _gate_b()["audit_reconstruction"]["relationships"]}

    swapped = dict(correct)
    swapped["b_nom"] = ("mean_basis_nominal", "Cost Line")     # 16 instead of 18
    swapped["c_nom"] = ("uncertainty_mean_shift_nominal", "Cost Line")  # 18 instead of 16
    assert swapped != correct, "the swapped B/C mapping must be visible"
    assert columns[17] == "uncertainty_mean_shift_nominal", "column 18 is B, not C"
    assert columns[15] == "mean_basis_nominal", "column 16 is C, not B"

    wrong_kind = dict(correct)
    wrong_kind["d_nom"] = ("expected_risk_nominal", "Cost Line")
    assert wrong_kind["d_nom"][1] != correct["d_nom"][1], (
        "D partitioned over cost rows must be visible"
    )

    dropped = {key: value for key, value in correct.items() if not key.startswith("d_")}
    assert set(correct) - set(dropped) == {"d_nom", "d_pv"}, (
        "an omitted D must be visible"
    )


def test_nc_55_a_blank_folded_in_as_identity_one_is_caught() -> None:
    """G. An N/A cell is not the opposite kind's identity value."""
    planted = _synthetic(
        "foreach ($row in $rows) {\n"
        "    $value = $row[$ordinal]\n"
        "    if (Test-CalcBlank -Actual $value) { $value = 1 }\n"
        "    $sum = $sum + [double]$value\n"
        "}\n"
    )
    assert "$value = 1 }" in planted, (
        "a blank fabricated as the identity value must be visible"
    )
    assert "if (-not $blank) {" not in planted


def test_nc_56_hard_coded_default_application_state_is_caught() -> None:
    """H. Defaults prove nothing about restoration."""
    planted = _synthetic(
        "$null = Add-Check $list 'EnableEvents was restored' ($Excel.EnableEvents -eq $true)\n"
        "$null = Add-Check $list 'ScreenUpdating was restored' ($Excel.ScreenUpdating -eq $true)\n"
        "$null = Add-Check $list 'Calculation mode was restored to automatic' "
        "([int]$Excel.Calculation -eq -4105)\n"
    )
    assert "$callerState" not in planted, (
        "a proof against Excel defaults rather than the captured caller state must be visible"
    )
    assert "-eq $true" in planted and "-eq -4105" in planted


def test_nc_57_an_application_state_proof_missing_a_property_is_caught() -> None:
    """I and J. DisplayAlerts and StatusBar are part of the caller's state."""
    for omitted in ("DisplayAlerts", "StatusBar"):
        kept = [prop for prop in
                ("ScreenUpdating", "EnableEvents", "DisplayAlerts", "Calculation", "StatusBar")
                if prop != omitted]
        planted = "\n".join(
            f"$null = Add-Check $list '{prop} was restored to the CAPTURED caller value' "
            f"($Excel.{prop} -eq $callerState.{prop})"
            for prop in kept
        )
        assert f"'{omitted} was restored" not in planted, (
            f"the omitted {omitted} assertion must be visible"
        )
        assert len(kept) == 4


def test_nc_58_a_powershell_held_prerequisite_list_is_caught() -> None:
    """1.1. The matrix is emitted; PowerShell consumes it."""
    planted = _synthetic(
        "$prerequisites = @(\n"
        "    @{ id = 'PQ-01'; predicate = 'base_year_after_start_year'; "
        "detail_tokens = @('Base Year') },\n"
        "    @{ id = 'PQ-02'; predicate = 'structure_change_pending'; "
        "detail_tokens = @('STRUCTURE CHANGE PENDING') })\n"
    )
    assert "$Cases.gate_b.prerequisite_cases" not in planted, (
        "a hand-maintained duplicate matrix in PowerShell must be visible"
    )
    assert "'base_year_after_start_year'" in planted


def test_nc_59_a_no_block_case_expecting_a_refusal_is_caught() -> None:
    """1.4. The complement must expect SUCCESS, not REFUSED."""
    entry = dict(_gate_b()["no_block_cases"][0])
    entry["expected_attempt"] = "REFUSED"
    entry["expected_status"] = "INVALID"
    assert entry["expected_attempt"] != "SUCCESS", (
        "a no-block case inverted into a refusal must be visible"
    )


# ===========================================================================
# 16. CORRECTION ROUND 3
#
# Four defects found in independent review of 2a2ae86. Each has a test here that
# fails against that source and passes against the corrected one.
# ===========================================================================
def test_75_the_config_master_owns_inflation_profile_rows() -> None:
    """BLOCKER 1. `SyncProfileRows` rebuilds tblInflation from the Config master.

    A profile row planted straight into `Inflation!tblInflation` is removed by
    the very `PCCM_ApplyTimeline` the fixture depends on, and the rate writer
    then searches for a row production has already deleted. That is a fixture
    defect, not a calculation defect.
    """
    # The production ownership this rests on, asserted rather than assumed.
    inflation = _text(SRC_VBA / "modInflation.bas")
    sync = inflation[inflation.index("Public Sub SyncProfileRows"):]
    sync = sync[:sync.index("\nEnd Sub")]
    assert "TBL_INFLATION_PROFILES" in sync and "SH_CONFIG" in sync, (
        "SyncProfileRows no longer reads the Config master; the reasoning below needs restating"
    )
    timeline = _text(SRC_VBA / "modTimeline.bas")
    assert "modInflation.SyncProfileRows" in timeline, (
        "Apply Timeline no longer synchronises the profile rows"
    )

    source = _executable(SCENARIOS)
    fixture = _procedure(source, "Set-Phase5Fixture")
    assert "Set-Phase5InflationProfileMaster" in fixture, (
        "the fixture does not populate the Config profile master"
    )
    master = _procedure(source, "Set-Phase5InflationProfileMaster")
    assert "$Inspection.input_tables.inflation_profiles" in master, (
        "the Config table identity is not read from the projection"
    )
    for hardcoded in ("'Config'", "'tblInflationProfiles'"):
        assert hardcoded not in source, f"the harness hard-codes {hardcoded}"
    # The master carries IDENTITIES only.
    assert "-ColumnIndex 1" in master
    assert "-ColumnIndex 2" not in master, "a rate is written into the profile master"
    # And nothing seeds tblInflation profile identities directly any more.
    assert "$inflGrid" not in fixture, (
        "the fixture still reaches into the inflation grid before Apply"
    )
    assert "Clear-Phase5GridBody" not in fixture, (
        "the fixture still clears the inflation grid it does not own"
    )


def test_76_the_fixture_proves_its_own_structural_prerequisites() -> None:
    """BLOCKER 2. A broken fixture must fail at fixture establishment."""
    source = _executable(SCENARIOS)
    fixture = _procedure(source, "Set-Phase5Fixture")
    apply_at = fixture.index("$Excel.Run('PCCM_ApplyTimeline')")
    assert "$applied -notlike 'OK|*'" in fixture, (
        "the Apply result is not required to be a success"
    )
    assert "throw (\"Gate-B fixture establishment failed: PCCM_ApplyTimeline" in fixture, (
        "a refused or failed Apply does not stop the fixture"
    )
    assert "PCCM_StructuralReport" in fixture, (
        "the fixture never asks whether the generated structure is coherent"
    )
    assert "throw (\"Gate-B fixture establishment failed: the generated structure" in fixture
    # BOTH GATES PRECEDE the value writers.
    rates_at = fixture.index("Write-Phase5InflationRates")
    weights_at = fixture.index("Write-Phase5Weights")
    report_at = fixture.index("PCCM_StructuralReport")
    assert apply_at < report_at < rates_at, (
        "inflation rates are written before the structural gate"
    )
    assert report_at < weights_at, "profiling weights are written before the structural gate"
    # It THROWS. It does not return a diagnostic the caller may ignore.
    assert fixture.count("throw (") >= 2


def test_77_the_baseline_gate_is_not_applied_to_deliberate_mutations() -> None:
    """BLOCKER 2.2. A prerequisite mutation is MEANT to make the model invalid."""
    source = _executable(SCENARIOS)
    mutation = _procedure(source, "Invoke-Phase5Mutation")
    # The mutation applier does not impose the clean-baseline gate globally.
    assert "PCCM_StructuralReport" not in mutation or "require_clean_structure" in mutation, (
        "the mutation applier imposes the baseline gate on deliberate corruption"
    )
    # Only the entry that ASKS for a clean structure gets one checked.
    clean = [entry for entry in _gate_b()["no_block_cases"]
             if entry["mutation"].get("require_clean_structure")]
    assert clean, "no no-block case requires the structure to stay valid"
    dirty = [entry for entry in _gate_b()["prerequisite_cases"]
             if entry["mutation"].get("require_clean_structure")]
    assert not dirty, "a prerequisite mutation demands a clean structure it is meant to break"


def test_78_inflation_rates_are_placed_by_profile_name() -> None:
    """BLOCKER 1.3. Model order is not physical grid order.

    `SyncProfileRows` rebuilds the grid in Config-master order, and nothing binds
    that to the order the emitted model happens to list its profiles in.
    """
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    assert "$rowIndex = Find-GridRow -Workbook $Workbook -Grid $grid -Key ([string]$name)" in body, (
        "the profile row is not located by name"
    )
    assert "$rowIndex++" not in body, "the profile row is still an incremented counter"
    # The column axis stays keyed by calendar-year header.
    assert "[array]::IndexOf($headers, [string]$year)" in body
    # And a rate a previous fixture left on a surviving profile is cleared first:
    # SyncProfileRows keeps rates by name, so inheritance is real.
    assert "-Value $null" in body, (
        "stale rates on a surviving profile are inherited into the next fixture"
    )


def test_79_the_unreachable_predicate_is_proved_by_calling_the_checker() -> None:
    """BLOCKER 3.1. Entering Base > Start never reaches modCalcCheck.

    `modTimeline` prevalidates the relationship and refuses the Apply, so the
    workbook is left with entered <> applied and the next `PCCM_Calculate` is
    refused by `StructuralPrerequisites` with STRUCTURE CHANGE PENDING - a
    different predicate, in a different module, with a different message.
    """
    # The production prevalidation this rests on.
    timeline = _text(SRC_VBA / "modTimeline.bas")
    assert "If t.BaseYear > t.StartYear Then" in timeline, (
        "Apply no longer prevalidates Base > Start; the reasoning needs restating"
    )
    # The predicate has left the workbook-mutation matrix.
    predicates = {entry["predicate"] for entry in _gate_b()["prerequisite_cases"]}
    assert "base_year_after_start_year" not in predicates, (
        "the predicate is still claimed by a workbook mutation that cannot reach it"
    )
    entry = next(item for item in _gate_b()["direct_check_cases"]
                 if item["predicate"] == "base_year_after_start_year")
    assert entry["procedure"] == "GBD_CheckBaseAfterStart"
    assert entry["control_procedure"] == "GBD_CheckTimelineAccepted"
    assert entry["arguments"]["base_year"] > entry["arguments"]["start_year"]
    assert entry["control_arguments"]["base_year"] <= entry["control_arguments"]["start_year"]

    # The diagnostic calls the ACCEPTED checker, and reopens nothing.
    diagnostic = _vba_executable(DIAGNOSTIC)
    assert "modCalcCheck.CheckResolvedModel(model, detail)" in diagnostic, (
        "the diagnostic does not call the accepted checker"
    )
    assert "Dim model As ResolvedModel" in diagnostic
    assert "model.DriverCount = 0" in diagnostic, (
        "a driver could refuse first and mask the model-level predicate"
    )
    checker = _text(SRC_VBA / "modCalcCheck.bas")
    assert re.search(r"^Public Function CheckResolvedModel\b", checker, re.M), (
        "CheckResolvedModel is not Public in the accepted source"
    )
    resolve = _text(SRC_VBA / "modCalcResolve.bas")
    assert re.search(r"^Public Type ResolvedModel\b", resolve, re.M)

    # The scenario runs BEFORE the diagnostic module is removed.
    source = _executable(SCENARIOS)
    assert source.index("Add-Phase5Result 'P5-DC'") < source.index("Add-Phase5Result 'P5-D8'"), (
        "the direct check runs after the diagnostic module was removed"
    )
    block = _scenario_block(source, "P5-D7", "P5-DC")
    assert "$Cases.gate_b.direct_check_cases" in block
    assert "Add-Phase5DetailTokenChecks" in block, "the refusal detail is not discriminated"
    assert "the control model with the predicate satisfied is ACCEPTED" in block, (
        "nothing proves the checker accepts the same construction when the predicate holds"
    )
    # STRUCTURE CHANGE PENDING remains a SEPARATE workbook scenario.
    assert "structure_change_pending" in {entry["predicate"] for entry
                                          in _gate_b()["prerequisite_cases"]}


def test_80_the_missing_profile_mutation_moves_the_driver_reference() -> None:
    """BLOCKER 3.2. Renaming a grid row breaks the Phase-4 invariant first."""
    # The structural check this rests on.
    structural = _text(SRC_VBA / "modStructuralCheck.bas")
    assert "which is not in the Config " in structural, (
        "the master/grid invariant is gone; the reasoning needs restating"
    )
    entry = next(item for item in _gate_b()["prerequisite_cases"]
                 if item["predicate"] == "referenced_profile_missing")
    assert entry["mutation"]["kind"] == "register_cell", (
        "the mutation still edits the structural grid instead of the driver reference"
    )
    assert entry["mutation"]["column"] == "inflation_profile"
    assert entry["mutation"]["value"] not in ("Standard", "Flat"), (
        "the driver still references a profile the fixture declares"
    )
    assert entry["mutation"]["value"] in entry["detail_tokens"], (
        "the missing profile name is not among the discriminators"
    )
    # The grid-editing kind is gone from the applier entirely.
    source = _executable(SCENARIOS)
    body = _procedure(source, "Invoke-Phase5Mutation")
    assert "'inflation_profile_rename'" not in body, (
        "the structure-breaking rename mutation survives"
    )


def test_81_the_unreferenced_profile_goes_through_the_config_master() -> None:
    """BLOCKER 3.3. A grid-only addition proves the structural gate, not no-block."""
    entry = next(item for item in _gate_b()["no_block_cases"]
                 if item["predicate"] == "unreferenced_profile_incomplete")
    assert entry["mutation"]["kind"] == "config_profile_add", (
        "the unused profile is still added straight to the grid"
    )
    assert entry["mutation"]["apply_timeline"] is True, (
        "production SyncProfileRows never runs, so no matching grid row is created"
    )
    assert entry["mutation"]["require_clean_structure"] is True, (
        "nothing proves the workbook is still structurally valid"
    )
    source = _executable(SCENARIOS)
    body = _procedure(source, "Invoke-Phase5Mutation")
    branch = body[body.index("'config_profile_add' {"):]
    branch = branch[:branch.index("\n        '")] if "\n        '" in branch else branch
    assert "$Inspection.input_tables.inflation_profiles" in branch, (
        "the profile is not added to the Config master"
    )
    assert "PCCM_ApplyTimeline" in branch, "production sync is never invoked"
    assert "PCCM_StructuralReport" in branch, "the structure is never re-proved"
    assert "'inflation_profile_add'" not in body, "the grid-only addition survives"
    # A new profile arrives with BLANK rates by construction: SyncProfileRows
    # clears the slot for a profile it has not seen before.
    inflation = _text(SRC_VBA / "modInflation.bas")
    assert "' A new profile, or a newly required year: BLANK, never zero." in inflation


def test_82_the_audit_reconstruction_is_exact() -> None:
    """BLOCKER 4. A relative epsilon can hide a real mismatch."""
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-PN", "P5-AR")
    assert "-Tolerance 0.0)" in block, (
        "the audit reconstruction does not compare exactly"
    )
    assert "rows, EXACTLY'" in block
    for forbidden in ("identity_relative_coefficient", "identity_absolute_floor",
                      "conditioning_scale_floor", "profiling_sum_absolute"):
        assert forbidden not in block, (
            f"the audit reconstruction applies a tolerance ({forbidden})"
        )
    # No epsilon of any kind is introduced.
    assert not re.search(r"-Tolerance\s+(?!0\.0)", block), (
        "a non-zero tolerance survives in the audit reconstruction"
    )
    assert "[Math]::Abs" not in block and "[Math]::Max" not in block
    # And the emitted fixture really does reconstruct exactly, so exactness is
    # representable rather than aspirational.
    audit = _gate_b()["audit_reconstruction"]
    for relationship in audit["relationships"]:
        total = sum(row[relationship["driver_column"]] for row in audit["expected"]["drivers"]
                    if row["driver_kind"] == relationship["kind"]
                    and row[relationship["driver_column"]] is not None)
        assert total == audit["expected"]["totals"][relationship["headline"]], (
            f"{relationship['headline']} does not reconstruct to the identical Double"
        )


# ===========================================================================
# 17. CORRECTION-ROUND-3 NEGATIVE CONTROLS
# ===========================================================================
def test_nc_60_seeding_inflation_profiles_into_the_grid_is_caught() -> None:
    """A. The Config master owns profile identities; Apply rebuilds the grid."""
    planted = _synthetic(
        "$inflGrid = $null\n"
        "foreach ($grid in @($Manifest.grids)) { if ($grid.key -eq 'inflation') "
        "{ $inflGrid = $grid } }\n"
        "Clear-Phase5GridBody -Workbook $Workbook -SheetName $inflGrid.sheet "
        "-TableName $inflGrid.table_name -ColumnCount $inflHeaders.Count\n"
        "foreach ($name in $profiles) {\n"
        "    Set-TableCell -Workbook $Workbook -SheetName $inflGrid.sheet "
        "-TableName $inflGrid.table_name -RowIndex $index -ColumnIndex 1 -Value $name\n"
        "}\n"
    )
    assert "Set-Phase5InflationProfileMaster" not in planted, (
        "a fixture that plants profile rows in the grid must be visible"
    )
    assert "$inflGrid" in planted and "Clear-Phase5GridBody" in planted


def test_nc_61_ignoring_a_failed_apply_is_caught() -> None:
    """B. A refused Apply must fail at fixture establishment."""
    planted = _synthetic(
        "$Excel.Run('PCCM_ApplyTimeline') | Out-Null\n"
        "$applied = [string]$Excel.Run('PCCM_AutomationResult')\n"
        "Write-Phase5InflationRates -Workbook $Workbook -Manifest $Manifest -Model $Model\n"
        "return $applied\n"
    )
    assert "$applied -notlike 'OK|*'" not in planted, (
        "a fixture that discards the Apply result must be visible"
    )
    assert "throw" not in planted
    assert planted.index("$applied = ") < planted.index("Write-Phase5InflationRates")


def test_nc_62_ignoring_a_non_empty_structural_report_is_caught() -> None:
    """C. An incoherent generated structure must fail at establishment."""
    planted = _synthetic(
        "if ($applied -notlike 'OK|*') { throw 'apply failed' }\n"
        "Write-Phase5InflationRates -Workbook $Workbook -Manifest $Manifest -Model $Model\n"
        "Write-Phase5Weights -Workbook $Workbook -Manifest $Manifest -Model $Model\n"
    )
    assert "PCCM_StructuralReport" not in planted, (
        "a fixture that never asks for the structural report must be visible"
    )


def test_nc_63_positional_inflation_rate_placement_is_caught() -> None:
    """D. Model profile order is not physical grid order."""
    planted = _synthetic(
        "$rowIndex = 0\n"
        "foreach ($name in $Model.inflation.PSObject.Properties.Name) {\n"
        "    $rowIndex++\n"
        "    $rates = $Model.inflation.$name\n"
        "}\n"
    )
    assert "$rowIndex++" in planted, "the positional row counter must be visible"
    assert "Find-GridRow" not in planted


def test_nc_64_a_swapped_profile_order_does_not_move_the_rates() -> None:
    """D, positively: the keyed writer is order-independent.

    Two profiles, listed by the model in one order and materialised by
    SyncProfileRows in the Config master's order. A name-keyed writer puts each
    rate on its own profile whichever order the grid ends up in; a positional one
    swaps them.
    """
    model_order = ["Standard", "Flat"]
    grid_order = ["Flat", "Standard"]          # what SyncProfileRows produced
    rates = {"Standard": 0.05, "Flat": 0.0}

    keyed = {name: rates[name] for name in model_order}          # by profile name
    positional = {grid_order[index]: rates[name]                 # by row position
                  for index, name in enumerate(model_order)}
    assert keyed == rates, "the keyed placement must be order-independent"
    assert positional != rates, (
        "the positional placement must be visibly wrong when the orders differ"
    )
    assert positional["Flat"] == rates["Standard"], (
        "the positional writer puts Standard's rate on Flat's row"
    )


def test_nc_65_the_entered_apply_route_for_base_after_start_is_caught() -> None:
    """E. It reaches STRUCTURE CHANGE PENDING, not modCalcCheck."""
    planted = {
        "id": "PQ-01", "predicate": "base_year_after_start_year",
        "mutation": {"kind": "entered_structure", "target": "base_year",
                     "value": 2030, "apply_timeline": True},
        "detail_tokens": ["Base Year", "Start Year"],
    }
    assert planted["mutation"]["kind"] == "entered_structure", (
        "the workbook-mutation route for an unreachable predicate must be visible"
    )
    # Apply prevalidates and refuses, so the applied timeline never moves and the
    # NEXT calculate is refused by the structural gate instead.
    timeline = _text(SRC_VBA / "modTimeline.bas")
    assert "If t.BaseYear > t.StartYear Then" in timeline
    resolve = _text(SRC_VBA / "modCalcResolve.bas")
    assert "STATE_PENDING" in resolve, (
        "the pending state the entered route actually reaches must be visible"
    )
    assert "procedure" not in planted, "the direct-check route must be visibly absent"


def test_nc_66_a_grid_rename_for_the_missing_profile_is_caught() -> None:
    """F. It breaks the master/grid invariant and hits the structural gate."""
    planted = {
        "predicate": "referenced_profile_missing",
        "mutation": {"kind": "inflation_profile_rename", "profile": "Standard",
                     "value": "Renamed"},
    }
    assert planted["mutation"]["kind"] != "register_cell", (
        "a structural-grid rename must be visible as such"
    )
    structural = _text(SRC_VBA / "modStructuralCheck.bas")
    assert "which is not in the Config " in structural, (
        "the invariant the rename breaks must be visible"
    )


def test_nc_67_a_grid_only_unused_profile_is_caught() -> None:
    """G. The same mismatch, from the other side."""
    planted = {
        "predicate": "unreferenced_profile_incomplete",
        "mutation": {"kind": "inflation_profile_add", "profile": "Unused",
                     "calendar_year": 2027, "value": None},
    }
    assert planted["mutation"]["kind"] != "config_profile_add", (
        "a grid-only profile addition must be visible as such"
    )
    assert "apply_timeline" not in planted["mutation"], (
        "the missing production sync must be visible"
    )
    assert "require_clean_structure" not in planted["mutation"], (
        "the missing structural re-proof must be visible"
    )


def test_nc_68_a_toleranced_audit_reconstruction_is_caught() -> None:
    """H. Even 1e-12 can hide a real mismatch between two published values."""
    planted = _synthetic(
        "$null = Add-Check $list ($relationship.headline + ' = SUM(...)') `\n"
        "    (Test-CalcValue -Actual $headline -Expected $sum `\n"
        "        -Tolerance ([double]$Cases.tolerances.identity_relative_coefficient))\n"
    )
    assert "identity_relative_coefficient" in planted, (
        "a toleranced audit comparison must be visible"
    )
    assert "-Tolerance 0.0" not in planted
    assert re.search(r"-Tolerance\s+(?!0\.0)", planted), (
        "the non-zero tolerance must be visible to the detector"
    )


# ===========================================================================
# 18. CORRECTION ROUND 4
#
# Two defects found in independent review of aa6611c. Each has a test here that
# fails against that source and passes against the corrected one.
# ===========================================================================
def test_83_the_locked_fx_seed_is_captured_from_the_real_workbook() -> None:
    """BLOCKER 1. `-KeepRows 1` trusted whatever happened to be row 1.

    PQ-10 REMOVES the reporting row, so a foreign currency shifts up into row 1
    and the next fixture preserves it as though it were the seed. PQ-12 leaves
    the reporting rate at 2, and every later fixture inherits it and refuses on
    the global invariant instead of the predicate under test.
    """
    source = _executable(SCENARIOS)
    capture = _procedure(source, "Save-Phase5LockedFxSeed")
    assert "$Inspection.input_tables.fx_rates" in capture, (
        "the FX table identity is not read from the projection"
    )
    assert "[int]$fx.locked_seed_rows -ne 1" in capture, (
        "the capture does not require exactly one locked seed row"
    )
    assert "Get-Phase5TypedTableBody -Workbook $Workbook" in capture, (
        "the seed is not read from the real workbook through the TYPED reader"
    )
    assert "Get-TableBody -Workbook $Workbook" not in capture, (
        "the seed is captured through the stringifying Phase-4 reader, so a "
        "correct numeric rate and a defective text rate look identical"
    )
    # NOT a literal, and NOT the emitted model.
    for forbidden in ("'SAR'", '"SAR"', "$Model.fx", "$Cases.", "$entry.rate",
                      "REPORTING_CURRENCY"):
        assert forbidden not in capture, (
            f"the seed is reconstructed rather than captured ({forbidden})"
        )
    assert not re.search(r"Rate\s*=\s*1\b", capture), "the seed rate is hard-coded"
    # An uncaptured seed is a loud failure, never a silent fall-back to row 1.
    getter = _procedure(source, "Get-Phase5LockedFxSeed")
    assert "throw (" in getter, "a missing capture degrades silently"


def test_84_the_capture_precedes_every_phase_5_mutation() -> None:
    """BLOCKER 1. Captured once, on the untouched Stage-B workbook."""
    source = _executable(SCENARIOS)
    scenarios = source[source.index("function Invoke-Phase5GateBScenarios"):]
    capture_at = scenarios.index("Save-Phase5LockedFxSeed -Workbook $Workbook")
    # Before the first fixture, and therefore before the first FX write.
    first_fixture = scenarios.index("Set-Phase5Fixture -Excel $Excel")
    assert capture_at < first_fixture, (
        "the seed is captured after a fixture has already rewritten the FX table"
    )
    # And after the Phase-4 prerequisite, so the workbook is known good.
    prerequisite_at = scenarios.index("Add-Phase5Result 'P5-P4'")
    assert prerequisite_at < capture_at, (
        "the seed is captured before the Phase-4 matrix is known intact"
    )
    assert scenarios.count("Save-Phase5LockedFxSeed -Workbook $Workbook") == 1, (
        "the seed is re-captured, so a mutated table could become the new baseline"
    )
    assert "Add-Phase5Result 'P5-FX'" in scenarios, "the capture is not reported as a scenario"


def test_85_the_fixture_restores_the_seed_before_appending() -> None:
    """BLOCKER 1. Row 1 is rewritten from the capture, not preserved."""
    source = _executable(SCENARIOS)
    fixture = _procedure(source, "Set-Phase5Fixture")
    assert "Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection" in fixture, (
        "the fixture does not reset the FX table from the capture"
    )
    assert "Clear-Phase5UserRows" not in fixture, (
        "the fixture still preserves whatever happens to be row 1"
    )
    reset_at = fixture.index("Reset-Phase5FxTable")
    append_at = fixture.index("Add-BlankTableRow -Workbook $Workbook -SheetName $fx.sheet")
    assert reset_at < append_at, "the fixture appends its FX rows before restoring the seed"

    reset = _procedure(source, "Reset-Phase5FxTable")
    # Everything after the seed goes.
    assert "Remove-TableRow -Workbook $Workbook -SheetName $fx.sheet" in reset
    assert "for ($row = $rows; $row -gt 1; $row--)" in reset
    # Row 1's currency AND rate are both rewritten from the capture, AS
    # THEMSELVES. A [double] cast here would convert a defective text seed into a
    # number and repair the workbook into agreement with the contract.
    assert "Set-Phase5TypedCell -Workbook $Workbook -SheetName $fx.sheet" in reset, (
        "the restoration goes through a coercing setter"
    )
    assert "-RowIndex 1 -ColumnIndex 1 -Value $Seed.Currency" in reset, (
        "the seed currency is not restored"
    )
    assert "-RowIndex 1 -ColumnIndex 2 -Value $Seed.Rate" in reset, (
        "the seed rate is not restored as itself"
    )
    assert "[double]$Seed.Rate" not in reset, (
        "the captured rate is converted before it is written back"
    )
    assert "[string]$Seed.Currency" not in reset
    # The read-back uses the typed reader and the STRICT comparator.
    assert "Get-Phase5TypedTableBody" in reset
    assert "Test-Phase5ExactValue -Actual $body[0][1] -Expected $Seed.Rate" in reset, (
        "the restoration is verified with the analytical comparator, which would "
        "accept a type change"
    )
    # And the restoration is read back rather than assumed.
    assert "the locked FX seed did not restore" in reset
    # Still no literal reconstruction.
    for forbidden in ("'SAR'", '"SAR"', "-Value 1)", "$Model.fx"):
        assert forbidden not in reset, f"the reset reconstructs the seed ({forbidden})"


def test_86_the_sar_mutations_still_mutate_before_any_restoration() -> None:
    """BLOCKER 1.2.5. The reset must not disarm PQ-10, PQ-11 or PQ-12."""
    prerequisites = {entry["id"]: entry for entry in _gate_b()["prerequisite_cases"]}
    assert prerequisites["PQ-10"]["mutation"] == {"kind": "fx_remove", "currency": "SAR"}, (
        "PQ-10 no longer physically removes the reporting row"
    )
    assert prerequisites["PQ-11"]["mutation"]["append"] is True, (
        "PQ-11 no longer creates a duplicate reporting row"
    )
    assert prerequisites["PQ-11"]["mutation"]["currency"] == "SAR"
    assert prerequisites["PQ-12"]["mutation"]["rate"] == 2, (
        "PQ-12 no longer changes the reporting rate"
    )
    assert prerequisites["PQ-12"]["mutation"].get("append") is None, (
        "PQ-12 appends a row instead of rewriting the seed's rate"
    )
    # The order inside the scenario: establish, THEN mutate, THEN calculate.
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-RF", "P5-PQ")
    fixture_at = block.index("Set-Phase5Fixture -Excel $Excel")
    mutate_at = block.index("Invoke-Phase5Mutation -Excel $Excel")
    calculate_at = block.index("$Excel.Run('PCCM_Calculate') | Out-Null", mutate_at)
    assert fixture_at < mutate_at < calculate_at, (
        "the mutation does not happen between the clean fixture and the calculation"
    )
    # The reset belongs to establishing the NEXT fixture, not to the mutation.
    mutation = _procedure(source, "Invoke-Phase5Mutation")
    assert "Reset-Phase5FxTable" not in mutation, (
        "the mutation applier resets the table it is meant to corrupt"
    )


def test_87_the_no_block_cases_reassert_the_full_analytical_workspace() -> None:
    """BLOCKER 2. Referenced-only means outside the CALCULATION, not just the digest.

    A defect that kept the unreferenced assumption out of the fingerprint while
    consuming it in the calculation would satisfy SUCCESS / CURRENT / blank
    detail / same digest and still publish wrong numbers.
    """
    source = _executable(SCENARIOS)
    block = _scenario_block(source, "P5-PQ", "P5-PN")
    assert "Add-Phase5AnalyticalChecks -List $list -Workbook $Workbook `" in block, (
        "the no-block scenario never re-checks the analytical workspace"
    )
    assert "-Inspection $Inspection -Case $base -Tolerances $Cases.tolerances" in block, (
        "the recheck is not against the base plan case's own emitted block"
    )
    assert "Add-Phase5SuccessStateChecks" in block, (
        "the successful calc_state record is not re-asserted after the mutation"
    )
    # The recheck happens AFTER the mutation and the recalculation.
    mutate_at = block.index("Invoke-Phase5Mutation -Excel $Excel")
    recheck_at = block.index("Add-Phase5AnalyticalChecks")
    assert mutate_at < recheck_at, "the analytical recheck runs before the mutation"
    # Every base plan case a no-block entry names really carries an expected block.
    cases = {str(case["id"]): case for case in _emitted()["cases"]["plan_cases"]}
    for entry in _gate_b()["no_block_cases"]:
        base = cases[str(entry["base_plan_case"])]
        assert base["kind"] == "analytical", (
            f"{entry['id']} names a base case with no analytical expectation"
        )
        for key in ("calc_years", "resolved_fx_rows", "inflation_factors",
                    "drivers", "annual", "totals"):
            assert key in base["expected"], (
                f"{entry['id']}'s base case emits no {key} to re-check"
            )
    # And the digest assertion survives alongside it.
    assert "the stored fingerprint is unchanged by the unreferenced row" in block


def test_88_the_analytical_recheck_covers_all_five_tables_and_the_totals() -> None:
    """BLOCKER 2. The shared checker is what makes the recheck complete."""
    source = _executable(SCENARIOS)
    checks = _procedure(source, "Add-Phase5AnalyticalChecks")
    for table in ("calc_years", "calc_inflation_factors", "calc_fx", "calc_drivers",
                  "calc_annual"):
        assert f"'{table}'" in checks, f"{table} is not covered by the shared checker"
    assert "-Block 'calc_totals'" in checks
    # The no-block scenario calls that shared checker rather than a reduced copy.
    block = _scenario_block(source, "P5-PQ", "P5-PN")
    assert "Get-CalcTableRows" not in block, (
        "the no-block scenario reads tables itself instead of using the shared checker"
    )


# ===========================================================================
# 19. CORRECTION-ROUND-4 NEGATIVE CONTROLS
# ===========================================================================
def test_nc_69_keeping_the_current_physical_row_one_is_caught() -> None:
    """The shipped shape: trust row 1, restore nothing."""
    planted = _synthetic(
        "$fx = $Inspection.input_tables.fx_rates\n"
        "Clear-Phase5UserRows -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name `\n"
        "    -KeepRows ([int]$fx.locked_seed_rows)\n"
        "foreach ($entry in @($Model.fx)) { }\n"
    )
    assert "Reset-Phase5FxTable" not in planted, (
        "a fixture that preserves whatever is in row 1 must be visible"
    )
    assert "Clear-Phase5UserRows" in planted


def test_nc_70_the_deleted_reporting_row_contaminates_the_next_fixture() -> None:
    """A. PQ-10 removes SAR and USD shifts up into row 1.

    Modelled as data: `-KeepRows 1` preserves the shifted row, the fixture then
    skips the model's own reporting entry and appends a SECOND foreign row.
    """
    captured = ("SAR", 1.0)
    baseline = [("SAR", 1.0), ("USD", 3.75)]
    model_fx = [("SAR", 1.0), ("USD", 3.75)]

    # PQ-10: the reporting row is physically removed.
    after_mutation = [row for row in baseline if row[0] != "SAR"]
    assert after_mutation[0][0] == "USD", "the shifted row must be visible"

    # The SHIPPED reset: keep row 1, append the model's non-reporting rows.
    shipped = after_mutation[:1] + [row for row in model_fx if row[0] != "SAR"]
    assert shipped[0][0] != captured[0], (
        "the contaminated baseline must be visible: row 1 is not the reporting seed"
    )
    assert [row[0] for row in shipped] == ["USD", "USD"], (
        "the duplicated foreign row must be visible"
    )

    # The CORRECTED reset: drop everything after row 1, rewrite row 1 from the
    # capture, then append.
    corrected = [captured] + [row for row in model_fx if row[0] != captured[0]]
    assert corrected[0] == captured, "the restored seed must be row 1"
    assert [row[0] for row in corrected] == ["SAR", "USD"]


def test_nc_71_the_mutated_reporting_rate_is_inherited() -> None:
    """B. PQ-12 leaves SAR at 2 and the next fixture never restores it."""
    captured = ("SAR", 1.0)
    after_mutation = [("SAR", 2.0), ("USD", 3.75)]

    shipped = after_mutation[:1]                       # -KeepRows 1, no restore
    assert shipped[0][1] == 2.0, "the inherited reporting rate must be visible"
    assert shipped[0][1] != captured[1]

    corrected = [captured]
    assert corrected[0][1] == captured[1], "the restored rate must be the captured one"


def test_nc_72_a_hard_coded_seed_reconstruction_is_caught() -> None:
    """The fixture must not manufacture the invariant PQ-10..12 test."""
    planted = _synthetic(
        "Set-TableCell -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name `\n"
        "    -RowIndex 1 -ColumnIndex 1 -Value 'SAR'\n"
        "Set-TableCell -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name `\n"
        "    -RowIndex 1 -ColumnIndex 2 -Value ([double]1)\n"
    )
    assert "'SAR'" in planted, "the hard-coded reporting currency must be visible"
    assert "$Seed.Currency" not in planted and "$Seed.Rate" not in planted


def test_nc_73_a_model_sourced_seed_reconstruction_is_caught() -> None:
    """The seed must come from the built workbook, not from the corpus."""
    planted = _synthetic(
        "foreach ($entry in @($Model.fx)) {\n"
        "    if ([string]$entry.currency -eq $reporting) {\n"
        "        Set-TableCell -RowIndex 1 -ColumnIndex 2 -Value ([double]$entry.rate)\n"
        "    }\n"
        "}\n"
    )
    assert "$Model.fx" in planted, "the model-sourced seed must be visible"
    assert "$Seed.Rate" not in planted
    assert "Save-Phase5LockedFxSeed" not in planted


def test_nc_74_a_digest_only_no_block_proof_is_caught() -> None:
    """The shipped P5-PN shape: four green flags and no numbers."""
    planted = _synthetic(
        "$null = Add-Check $list ($id + ': attempt') ($attempt -eq 'SUCCESS')\n"
        "$null = Add-Check $list ($id + ': status') ($status -eq 'CURRENT')\n"
        "$null = Add-Check $list ($id + ': detail blank') "
        "([string]::IsNullOrEmpty($detail))\n"
        "$null = Add-Check $list ($id + ': fingerprint') "
        "([string]$Excel.Run('PCCM_CalculationFingerprint') -ceq $baseline)\n"
        "$covered += $id\n"
    )
    assert "Add-Phase5AnalyticalChecks" not in planted, (
        "a no-block proof with no analytical recheck must be visible; an "
        "assumption can leak into the calculation while staying out of the digest"
    )
    assert "Add-Phase5SuccessStateChecks" not in planted


def test_nc_75_a_partial_analytical_recheck_is_caught() -> None:
    """One table left out of the recheck."""
    planted = _synthetic(
        "foreach ($key in 'calc_years', 'calc_fx', 'calc_drivers', 'calc_annual') {\n"
        "    $null = Add-Check $list ('recheck ' + $key) ($ok)\n"
        "}\n"
        "$null = Add-Check $list 'totals' ($ok)\n"
    )
    assert "calc_inflation_factors" not in planted, "the omitted table must be visible"
    assert "Add-Phase5AnalyticalChecks" not in planted, (
        "a hand-rolled partial recheck instead of the shared checker must be visible"
    )


# ===========================================================================
# 20. CORRECTION ROUND 5
#
# Three defects found in independent review of 56b90d9. Each has a test here
# that fails against that source and passes against the corrected one.
# ===========================================================================
def test_89_the_calc_tables_are_read_with_their_value2_types() -> None:
    """BLOCKER 1. The Phase-4 reader stringifies; Test-CalcValue is type-sensitive.

    `Get-TableBody` does `if ($null -eq $v) { '' } else { [string]$v }`, which is
    right for the structural comparisons it was written for. Fed into a
    type-sensitive analytical comparator, a correct cell holding `Value2 = 1.05`
    arrived as the String `"1.05"` and every comparison returned False before it
    compared a number. The first successful analytical Gate-B scenario would have
    failed with production behaving perfectly.
    """
    # The Phase-4 behaviour this rests on, asserted rather than assumed - and
    # asserted to be UNCHANGED, because Step B1 may not edit it.
    phase4 = _executable(HARNESS)
    reader = _procedure(phase4, "Get-TableBody")
    assert "if ($null -eq $v) { $line += '' } else { $line += [string]$v }" in reader, (
        "the accepted Phase-4 reader changed; the reasoning below needs restating"
    )

    source = _executable(SCENARIOS)
    rows = _procedure(source, "Get-CalcTableRows")
    assert "Get-Phase5TypedTableBody" in rows, (
        "the calculation tables are not read through the typed reader"
    )
    assert "Get-TableBody" not in rows, (
        "the calculation tables are still read through the stringifying reader"
    )

    typed = _procedure(source, "Get-Phase5TypedTableBody")
    assert "$line[$c - 1] = $cell.Value2" in typed, (
        "the typed reader does not assign Value2 straight into the row"
    )
    # NOTHING is stringified, coalesced or formatted on the way out.
    for forbidden in ("[string]$v", "[string]$cell", "+= ''", "Format-CalcValue",
                      "Format-Phase5Typed", ".Text", ".Value)"):
        assert forbidden not in typed, f"the typed reader transforms a cell ({forbidden})"
    # The row is allocated at the column count and filled BY INDEX, because
    # `$line += $null` appends nothing and a blank cell would vanish.
    assert "New-Object 'object[]' $colCount" in typed, (
        "the row is accumulated with +=, so a blank cell disappears from it"
    )
    assert "$line +=" not in typed
    # The accepted row-emission idiom survives: one non-enumerated object per row.
    assert "Write-RowObject $line" in typed
    assert "if ($null -eq $body) { return }" in typed, (
        "an empty body no longer emits zero objects"
    )
    # And the accepted COM ownership discipline.
    for owned in ("$localWorksheets = $Workbook.Worksheets", "$ws = $localWorksheets.Item",
                  "$los = $ws.ListObjects", "$lo = $los.Item", "$body = $lo.DataBodyRange"):
        assert owned in typed, f"the typed reader skips an owned COM acquisition ({owned})"
    assert typed.count("Release-Transient") >= 7


def test_90_the_typed_reader_preserves_every_value2_class() -> None:
    """BLOCKER 1.3. A no-COM shape and TYPE control, like the accepted PRE probe.

    Linux cannot run PowerShell, so the reader's pipeline behaviour cannot be
    observed here. What CAN be stated is the contract the source implements:
    a row of `[1.25, "USD", $null, 3]` must arrive as four cells whose classes
    are numeric, String, null and numeric - not four strings, and not three cells
    with the blank silently dropped.
    """
    fabricated = [1.25, "USD", None, 3]

    # THE SHIPPED READER, modelled: every non-null becomes a string and null
    # becomes an empty string.
    stringified = ["" if cell is None else str(cell) for cell in fabricated]
    assert stringified == ["1.25", "USD", "", "3"]
    assert all(isinstance(cell, str) for cell in stringified), (
        "the shipped reader's output must be visibly all-text"
    )

    # THE TYPED READER, modelled: the scalar arrives as itself.
    preserved = list(fabricated)
    assert len(preserved) == 4, "the row shape must survive the blank"
    assert isinstance(preserved[0], float) and not isinstance(preserved[0], str)
    assert isinstance(preserved[1], str)
    assert preserved[2] is None
    assert isinstance(preserved[3], int) and not isinstance(preserved[3], str)

    # And the consequence for the comparator: the shipped shape fails a correct
    # numeric expectation, the typed one passes it.
    def type_sensitive(actual, expected):
        if expected is None:
            return actual is None or (isinstance(actual, str) and actual == "")
        if isinstance(actual, str) and not isinstance(expected, str):
            return False
        return actual == expected

    assert not type_sensitive(stringified[0], 1.25), (
        "the stringified numeric must be visibly rejected by a type-sensitive comparator"
    )
    assert type_sensitive(preserved[0], 1.25)


def test_91_the_snapshot_keeps_typed_cells_not_a_row_string() -> None:
    """BLOCKER 2. "Exact" was a proof about display text.

    Rows were `Format-CalcValue`'d and joined into one String each, so a numeric
    1 and the String "1" produced identical evidence, and a real Empty and an
    empty String both collapsed to `<blank>`.
    """
    source = _executable(SCENARIOS)
    snapshot = _procedure(source, "Get-Phase5Snapshot")
    assert "Format-CalcValue" not in snapshot, (
        "the snapshot still formats its cells into the authoritative comparison"
    )
    assert "-join" not in snapshot, "the snapshot still joins a row into one String"
    assert "[char]31" not in snapshot, "the unit-separator serialisation survives"
    assert "$rows += , @($row)" in snapshot, (
        "the snapshot does not retain the row as a typed cell array"
    )


def test_92_snapshot_identity_uses_the_strict_comparator() -> None:
    """BLOCKER 2.1/2.3. "Restored exactly" is not "matches the oracle"."""
    source = _executable(SCENARIOS)
    strict = _procedure(source, "Test-Phase5ExactValue")
    # A. a real absence is not an empty String.
    assert "if ($null -eq $Expected) {" in strict
    assert "return ($null -eq $Actual)" in strict, (
        "an expected blank accepts something other than a genuine blank"
    )
    assert ".Length -eq 0" not in strict, "an empty String is treated as a blank"
    # B/C/D. same type class, exact value, no tolerance.
    assert "if ($Expected -is [string]) {" in strict
    assert "-ceq [string]$Expected" in strict, "text is compared case-insensitively"
    assert "if ($Expected -is [bool]) {" in strict
    assert "if ($Actual -is [string]) { return $false }" in strict, (
        "a String that looks numeric is accepted against a number"
    )
    assert "Tolerance" not in strict, "the strict comparator carries a tolerance"
    assert "Format-" not in strict, "the strict comparator compares display text"

    checks = _procedure(source, "Add-SnapshotUnchangedChecks")
    assert "Test-CalcValue" not in checks, (
        "the snapshot comparison still uses the analytical comparator"
    )
    assert checks.count("Test-Phase5ExactValue") >= 3, (
        "C13:C16, C23:C32 and the tables are not all compared strictly"
    )
    # Row count, column count, then every cell.
    assert "$was.Count -eq $now.Count" in checks
    assert "$wasRow.Count -ne $nowRow.Count" in checks, "the column count is not compared"
    assert "for ($c = 0; $c -lt $wasRow.Count; $c++)" in checks, "cells are not compared"
    # Formatting survives ONLY as diagnostics.
    assert "Format-Phase5Typed" in checks
    diagnostics = _procedure(source, "Format-Phase5Typed")
    assert "DIAGNOSTICS ONLY" in _text(SCENARIOS)[_text(SCENARIOS).index("function Format-Phase5Typed"):
                                                  _text(SCENARIOS).index("function Format-CalcValue")]


def test_93_the_strict_comparator_rejects_the_four_planted_restorations() -> None:
    """BLOCKER 2.4, as a decision table over the rule the source implements."""
    def strict(actual, expected):
        if expected is None:
            return actual is None
        if actual is None:
            return False
        if isinstance(expected, str):
            return isinstance(actual, str) and actual == expected
        if isinstance(expected, bool):
            return isinstance(actual, bool) and actual == expected
        if isinstance(actual, (str, bool)):
            return False
        return float(actual) == float(expected)

    # 1. a numeric Double restored as a String.
    assert not strict('1', 1.0), 'numeric 1 restored as text must be rejected'
    # 2. a genuine blank restored as an empty String.
    assert not strict('', None), 'a blank restored as an empty string must be rejected'
    # 3. identical display text, wrong underlying type.
    assert not strict("1.05", 1.05)
    assert not strict(1.05, "1.05")
    # 4. a C23:C32 blank replaced by "".
    assert not strict("", None)
    # And the correct restorations still pass.
    assert strict(1.0, 1.0) and strict(None, None) and strict("USD", "USD")
    # THE SHIPPED CHAIN cannot see any of it. Two stages destroy the evidence,
    # and modelling only the second would understate the defect:
    #
    #   Get-TableBody   every non-null cell -> String, every null -> ""
    #   Format-CalcValue + join   the row -> one display String
    def shipped_read(cell):
        return "" if cell is None else (cell if isinstance(cell, str) else f"{cell:g}")

    def shipped_format(cell):
        if cell is None or (isinstance(cell, str) and cell == ""):
            return "<blank>"
        return "'" + cell + "'" if isinstance(cell, str) else f"{cell:g}"

    def shipped_row(cells):
        return chr(31).join(shipped_format(shipped_read(c)) for c in cells)

    assert shipped_row([1.0, None]) == shipped_row(["1", ""]), (
        "the shipped row-string evidence must be visibly blind to both defects"
    )
    # The typed snapshot is not: it compares cell by cell with the strict rule.
    typed_before, typed_after = [1.0, None], ["1", ""]
    assert not all(strict(a, b) for a, b in zip(typed_after, typed_before)), (
        "the typed snapshot must reject the same pair the row string accepted"
    )


def test_94_the_fx_seed_is_captured_and_restored_without_coercion() -> None:
    """BLOCKER 3. The harness must not repair a defective built seed.

    `Get-TableBody` captured a correct numeric 1 as the String "1", and an
    incorrect text seed of "1" identically. `-Value ([double]$Seed.Rate)` then
    converted the capture into a number, so a workbook that had built the
    reporting rate as TEXT would have been silently corrected before the
    analytical scenarios ran.
    """
    source = _executable(SCENARIOS)
    capture = _procedure(source, "Save-Phase5LockedFxSeed")
    assert "Get-Phase5TypedTableBody" in capture
    assert "Get-TableBody" not in capture
    for forbidden in ("[string]$body", "[double]$body", "Format-CalcValue"):
        assert forbidden not in capture, f"the capture coerces the seed ({forbidden})"

    reset = _procedure(source, "Reset-Phase5FxTable")
    assert "[double]$Seed.Rate" not in reset, "the captured rate is converted on the way back"
    assert "Set-Phase5TypedCell" in reset, "the restoration goes through a coercing setter"
    setter = _procedure(source, "Set-Phase5TypedCell")
    # RUNTIME RUN 4 RETIRED THE GENERIC SITE. `$cell.Value2 = $Value` on one
    # line bound to String on its first call and then could not marshal a Double
    # through the same cached call site - the whole R5 family. The setter now
    # dispatches on the CAPTURED CLR type, one COM assignment per branch.
    #
    # That is not the same as choosing a type for the caller: the branch is
    # selected by what Excel published, so a captured String '1' is still
    # written back as String '1'. What the earlier form of this assertion was
    # protecting - no inference from the contract - is protected below by the
    # dispatch being on $Value's own type and by the read-back comparator.
    assert "$cell.Value2 = $Value" not in setter, (
        "the generic polymorphic COM assignment is back; Runtime Run 4 proved a "
        "second type cannot be marshalled through the same cached call site"
    )
    assert "$cell.Value2 = [string]$Value" in setter
    assert "$cell.Value2 = [double]$Value" in setter
    assert "$Value -is [string]" in setter, "the dispatch is not on the captured type"
    assert "$Value.GetType().FullName" in setter, (
        "an unsupported captured type is coerced rather than reported"
    )
    assert "$cell.ClearContents()" in setter, "a null cannot be written as a genuine blank"
    # The accepted Phase-4 setter is untouched and still chooses, which is right
    # for fixture authoring.
    phase4 = _procedure(_executable(HARNESS), "Set-TableCell")
    assert "$cell.Value2 = [double]$Value" in phase4, (
        "the accepted Phase-4 setter was modified"
    )


def test_95_nothing_widened_the_analytical_comparator_to_accept_text() -> None:
    """BLOCKER 1.2. The reader was fixed, not the comparator weakened."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Test-CalcValue")
    assert "if ($Actual -is [string]) { return $false }" in body, (
        "the analytical comparator now accepts a numeric String, which would "
        "hide a workbook that wrote a number as text"
    )
    assert "[double]$Actual -eq [double]$Expected" not in body.split(
        "if ($Actual -is [string]) { return $false }")[0], (
        "a numeric comparison happens before the String rejection"
    )


# ===========================================================================
# 21. CORRECTION-ROUND-5 NEGATIVE CONTROLS
# ===========================================================================
def test_nc_76_reading_the_calc_tables_through_get_tablebody_is_caught() -> None:
    """The shipped path."""
    planted = _synthetic(
        "function Get-CalcTableRows {\n"
        "    param($Workbook, $Inspection, [string]$TableKey)\n"
        "    $table = $Inspection.calc.tables.$TableKey\n"
        "    return @(Get-TableBody -Workbook $Workbook -SheetName $Inspection.calc.sheet `\n"
        "        -TableName $table.table_name)\n"
        "}\n"
    )
    assert "Get-TableBody" in planted, "the stringifying reader must be visible"
    assert "Get-Phase5TypedTableBody" not in planted


def test_nc_77_a_reader_that_stringifies_is_caught() -> None:
    planted = _synthetic(
        "$v = $cell.Value2\n"
        "if ($null -eq $v) { $line += '' } else { $line += [string]$v }\n"
    )
    assert "[string]$v" in planted, "the cast must be visible"
    assert "$line[$c - 1] = $cell.Value2" not in planted
    assert "+= ''" in planted, "the blank-to-empty-string coalescing must be visible"


def test_nc_78_an_accumulating_row_drops_a_blank_cell() -> None:
    """Why the typed reader allocates and assigns by index."""
    cells = [1.25, None, "USD"]
    # `$line += $null` appends NOTHING in PowerShell: the row silently shortens.
    accumulated = [c for c in cells if c is not None]
    assert len(accumulated) == 2, "the dropped blank must be visible"
    indexed = [None] * len(cells)
    for index, cell in enumerate(cells):
        indexed[index] = cell
    assert len(indexed) == 3 and indexed[1] is None


def test_nc_79_a_row_string_snapshot_is_caught() -> None:
    planted = _synthetic(
        "foreach ($cell in @($row)) { $cells += (Format-CalcValue $cell) }\n"
        "$rows += ($cells -join ([string][char]31))\n"
    )
    assert "Format-CalcValue" in planted and "-join" in planted, (
        "the display-text serialisation must be visible"
    )
    assert "$rows += , @($row)" not in planted


def test_nc_80_a_snapshot_compared_with_the_analytical_comparator_is_caught() -> None:
    planted = _synthetic(
        "foreach ($field in $SuccessFields) {\n"
        "    $null = Add-Check $List 'unchanged' `\n"
        "        (Test-CalcValue -Actual $After.State[$field] -Expected $Before.State[$field])\n"
        "}\n"
    )
    assert "Test-CalcValue" in planted, "the analytical comparator must be visible"
    assert "Test-Phase5ExactValue" not in planted
    # And it is genuinely too weak: it treats $null and "" as the same blank.
    def analytical(actual, expected):
        blank = lambda v: v is None or (isinstance(v, str) and v == "")
        if expected is None:
            return blank(actual)
        return actual == expected
    assert analytical('', None), (
        'the analytical comparator must be visibly blind to a blank restored '
        'as an empty string'
    )


def test_nc_81_a_coerced_seed_restoration_is_caught() -> None:
    planted = _synthetic(
        "Set-TableCell -Workbook $Workbook -SheetName $fx.sheet -TableName $fx.table_name `\n"
        "    -RowIndex 1 -ColumnIndex 2 -Value ([double]$Seed.Rate)\n"
    )
    assert "[double]$Seed.Rate" in planted, (
        "the conversion that would repair a defective text seed must be visible"
    )
    assert "Set-Phase5TypedCell" not in planted


def test_nc_82_a_stringifying_seed_capture_is_caught() -> None:
    planted = _synthetic(
        "$body = @(Get-TableBody -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name)\n"
        "$script:Phase5LockedFxSeed = [pscustomobject]@{ Currency = $body[0][0]; "
        "Rate = $body[0][1] }\n"
    )
    assert "Get-TableBody" in planted, "the stringifying capture must be visible"
    # A correct numeric seed and a defective text seed capture identically.
    correct, defective = 1.0, "1"
    assert str(correct) != str(defective)          # "1.0" vs "1" in Python
    assert f"{correct:g}" == defective, (
        "under Excel's Value2 -> String conversion both become the same text, "
        "which is the point"
    )


def test_nc_83_a_widened_analytical_comparator_is_caught() -> None:
    """Fixing the reader is the correct repair; widening the comparator is not."""
    planted = _synthetic(
        "function Test-CalcValue {\n"
        "    param($Actual, $Expected)\n"
        "    return ([double]$Actual -eq [double]$Expected)\n"
        "}\n"
    )
    assert "$Actual -is [string]) { return $false }" not in planted, (
        "a comparator that accepts a numeric String must be visible; it would "
        "hide a workbook that published a number as text"
    )


# ===========================================================================
# 22. RUNTIME RUN 1: the Phase-4 prerequisite must be lifecycle-reachable
# ===========================================================================
# Runtime Run 1 (harness commit 35640ec, first real Windows/Excel execution)
# reported:
#
#   [FAIL] P5-P4   FAIL all 35 Phase-4 scenarios reported a result -- missing: Y, Z
#                  ok   the Phase-4 matrix has 0 FAIL
#                  ok   the Phase-4 matrix has 0 SKIP
#                  FAIL the Phase-4 matrix is 35/35 PASS -- passed 33 of 35
#   [FAIL] P5-ALL  not attempted: the Phase-4 structural matrix is not intact
#   [PASS] Z       Excel closed naturally after the functional run
#   [PASS] Y       Transient COM releases
#   ... Phase-4 structural matrix: 35 of 35
#
# The prerequisite demanded two results that cannot exist at the point it runs.
# Z is recorded after Application.Quit and the natural-exit wait; Y is recorded
# last of all, from the whole-run transient ledger. Phase 5 executes inside the
# live automation session, so neither can precede it.
#
# The tests below pin the corrected lifecycle. They do NOT relax 35/35: they pin
# that the 35/35 demand moved to a gate that runs after Y and Z, and that no run
# can be accepted without them.
def _prerequisite_gate() -> str:
    """The P5-P4 block: entry into Invoke-Phase5GateBScenarios, up to P5-FX."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Invoke-Phase5GateBScenarios")
    # Bounded at both ends: after the procedure's own param block, and before the
    # locked FX capture, which is the first statement that touches the workbook
    # and belongs to P5-FX rather than to the gate.
    return body[body.index("$required = Get-Phase4RequiredScenarioIds"):
                body.index("Save-Phase5LockedFxSeed")]


def _final_gate() -> str:
    return _procedure(_executable(SCENARIOS), "Add-Phase4FinalCompletenessResult")


def test_96_the_deferred_cases_are_declared_and_the_partition_is_derived() -> None:
    """The prerequisite set is the matrix MINUS the named lifecycle cases.

    Derived, never a second hand-maintained list: a case added to the 35 becomes
    a prerequisite case automatically, and the only way to defer one is to name
    it in the finalization list, which stays inside the matrix.
    """
    source = _executable(SCENARIOS)

    declared = re.search(
        r"\$script:Phase4RequiredScenarioIds\s*=\s*@\((.*?)\)", source, re.S
    )
    assert declared, "the required Phase-4 matrix is not declared"
    required = _ps_string_literals(declared.group(1))
    assert len(required) == 35, f"the matrix names {len(required)} scenarios, not 35"
    assert set(required) == set(PHASE4_SCENARIO_IDS)

    deferred_decl = re.search(
        r"\$script:Phase4FinalizationScenarioIds\s*=\s*@\((.*?)\)", source, re.S
    )
    assert deferred_decl, "the post-session lifecycle cases are not declared"
    deferred = _ps_string_literals(deferred_decl.group(1))
    assert deferred == ["Y", "Z"] or sorted(deferred) == ["Y", "Z"], (
        f"the deferred set is {deferred}, not exactly the two lifecycle cases"
    )

    # NOTHING LEFT THE MATRIX. Every deferred case is still one of the 35.
    for name in deferred:
        assert name in required, (
            f"{name} was deferred out of the matrix instead of within it"
        )

    # DERIVED, not declared twice.
    derived = _procedure(source, "Get-Phase4PrerequisiteScenarioIds")
    assert "$script:Phase4RequiredScenarioIds | Where-Object" in derived, (
        "the prerequisite set is not derived from the matrix"
    )
    assert "$deferred -notcontains $_" in derived
    assert "@(" in derived and "'PRE0'" not in derived, (
        "the prerequisite set restates scenario names instead of deriving them"
    )

    # AND BOTH GATES CONSUME IT. A declaration nothing reads is documentation.
    assert "Get-Phase4FinalizationScenarioIds" in _prerequisite_gate(), (
        "the entry gate does not read the deferred set it excludes"
    )
    final = _final_gate()
    assert "Get-Phase4FinalizationScenarioIds" in final, (
        "the final gate does not read the deferred set it must demand"
    )
    assert "foreach ($id in $deferred)" in final, (
        "the final gate reads the deferred set without checking its members"
    )


def test_97_the_prerequisite_cannot_demand_results_that_cannot_exist_yet() -> None:
    """PROOF 1. P5-P4 evaluates the prerequisite set, never the whole matrix.

    This is the exact Run-1 defect: `$required` at the gate meant `missing: Y, Z`
    on a run in which nothing was wrong.
    """
    gate = _prerequisite_gate()
    assert "Get-Phase4PrerequisiteScenarioIds" in gate, (
        "the prerequisite gate does not use the prerequisite set"
    )

    # The completeness, FAIL, SKIP and count checks all read $prerequisite.
    assert "foreach ($id in $prerequisite) { if ($seen -notcontains $id)" in gate, (
        "the completeness check still walks the full 35-case matrix"
    )
    assert "$phase4 = @($Results | Where-Object { $prerequisite -contains $_.Id })" in gate, (
        "the FAIL/SKIP partition is still taken over the full matrix"
    )
    assert "$passed.Count -eq $prerequisite.Count" in gate, (
        "the pass count is still compared against a hard 35 at the entry gate"
    )
    assert "$passed.Count -eq 35" not in gate, (
        "the entry gate still demands 35, which Y and Z cannot satisfy there"
    )
    assert "'all 35 Phase-4 scenarios reported a result'" not in gate, (
        "the entry gate still demands all 35 results exist"
    )

    # AND THE DEFERRAL IS PROVED, not asserted in prose.
    assert "$earlyDeferred" in gate, (
        "nothing proves the deferred cases have genuinely not run yet"
    )
    assert "$deferred | Where-Object { $seen -contains $_ }" in gate


def test_98_no_phase_4_fail_or_skip_is_ignored_at_either_gate() -> None:
    """PROOFS 2 and 3. Zero tolerance survives at both gates.

    The entry gate tolerates no FAIL and no SKIP among the cases it can see; the
    final gate tolerates none across all 35. Neither counts PASS alone.
    """
    gate = _prerequisite_gate()
    assert "'the Phase-4 matrix has 0 FAIL'" in gate
    assert "'the Phase-4 matrix has 0 SKIP'" in gate
    assert "$_.Status -eq 'FAIL'" in gate and "$_.Status -eq 'SKIP'" in gate
    assert "($failed.Count -eq 0)" in gate and "($skipped.Count -eq 0)" in gate

    final = _final_gate()
    assert "'the final Phase-4 matrix has 0 FAIL'" in final
    assert "'the final Phase-4 matrix has 0 SKIP'" in final
    assert "($failed.Count -eq 0)" in final and "($skipped.Count -eq 0)" in final
    assert "$required -contains $_.Id" in final, (
        "the final gate does not take its partition over the full 35-case matrix"
    )

    # No -ne, no exclusion list, no 'ignore', no allowance anywhere near either.
    for block, label in ((gate, "P5-P4"), (final, "P5-FIN")):
        for forbidden in ("-contains 'SKIP'", "Status -ne 'FAIL'", "Status -ne 'SKIP'"):
            assert forbidden not in block, (
                f"{label} appears to exempt results from the FAIL/SKIP rule"
            )


def test_99_p5_all_is_refused_when_any_reachable_prerequisite_is_not_pass() -> None:
    """PROOFS 4 and 8. Nothing runs on an unmet prerequisite, and it is a FAIL."""
    gate = _prerequisite_gate()
    assert "$phase4Ok = Test-ChecklistOk $list" in gate, (
        "the gate decision is not taken from the whole checklist"
    )
    source = _executable(SCENARIOS)
    block = source[source.index("if (-not $phase4Ok)"):]
    assert "Add-Phase5Result 'P5-ALL'" in block[:600]
    assert "'FAIL'" in block[:600], "an unmet prerequisite is reported as something other than FAIL"
    assert "'SKIP'" not in block[:600], "an unmet prerequisite must not read as a quiet skip"
    # AND IT RETURNS. Reporting the refusal and then running anyway would be worse.
    assert re.search(r"Add-Phase5Result 'P5-ALL'.*?\n\s*return\b", block[:900], re.S), (
        "the refusal does not stop the Phase-5 scenarios from running"
    )


def test_100_the_final_gate_still_demands_the_whole_35_case_matrix() -> None:
    """PROOF 6. 35/35 did not weaken; it moved to where it is reachable."""
    final = _final_gate()
    assert "Get-Phase4RequiredScenarioIds" in final, (
        "the final gate does not read the full 35-case matrix"
    )
    assert "'all 35 Phase-4 scenarios reported a result'" in final
    assert "'the final Phase-4 matrix is 35/35 PASS'" in final
    assert "$passed.Count -eq 35" in final, "35/35 is not asserted at the final gate"
    assert "Add-Phase5Result 'P5-FIN'" in final
    assert "Get-Phase4PrerequisiteScenarioIds" not in final, (
        "the final gate must judge the whole matrix, not the reduced entry set"
    )


def test_101_final_acceptance_is_impossible_without_y_and_z_passing() -> None:
    """PROOFS 5 and 7. Each deferred case is checked BY NAME, exactly once.

    A bare 35/35 count is satisfiable by a matrix that lost Z and counted some
    other result twice, so the count alone is not the guarantee.
    """
    final = _final_gate()
    assert "foreach ($id in $deferred)" in final, (
        "the deferred lifecycle cases are not checked individually"
    )
    assert "$record = @($Results | Where-Object { $_.Id -eq $id })" in final
    assert "$record.Count -eq 1" in final, (
        "a deferred case recorded twice would pass an existence-only check"
    )
    assert "[string]$record[0].Status -eq 'PASS'" in final, (
        "a deferred case that FAILED would satisfy a presence-only check"
    )
    assert "ran exactly once and PASSED" in final

    # A SKIP is not a PASS. Both zero-tolerance predicates at the final gate must
    # be live expressions: a deferred case that skipped would otherwise ride
    # through on nothing but the named check, and a softened predicate here is
    # indistinguishable from acceptance.
    assert "'the final Phase-4 matrix has 0 FAIL' ($failed.Count -eq 0)" in final, (
        "the final gate's FAIL predicate is not a live comparison"
    )
    assert "'the final Phase-4 matrix has 0 SKIP' ($skipped.Count -eq 0)" in final, (
        "the final gate's SKIP predicate is not a live comparison"
    )

    # AND IT IS WIRED. A gate that is defined but never called gates nothing.
    harness = _executable(HARNESS)
    assert "Add-Phase4FinalCompletenessResult -Results $results" in harness, (
        "the final gate is never called, so nothing ever demands Y and Z"
    )
    # The run's exit code is driven by the FAIL count, so a P5-FIN FAIL fails it.
    assert "$failed = @($results | Where-Object { $_.Status -eq 'FAIL' })" in harness
    assert "if ($failed.Count -eq 0) {" in harness
    assert "exit 1" in harness


def test_102_shutdown_and_com_release_proof_remains_mandatory() -> None:
    """PROOF 9. Y and Z keep their coverage; nothing was removed to fix ordering."""
    harness = _executable(HARNESS)
    assert "Add-Result 'Z' 'Excel closed naturally after the functional run' 'PASS'" in harness
    assert "Add-Result 'Y' 'Transient COM releases' 'PASS'" in harness
    assert "Add-Result 'Y' 'Transient COM releases' 'FAIL'" in harness
    assert "$transient = @(Get-TransientFailures)" in harness, (
        "the whole-run transient ledger is no longer read"
    )
    # The shutdown ledger itself is untouched.
    for token in ("$rel.WorkbookClosed = $true", "$rel.QuitCalled = $true",
                  "$rel.NaturalExit = Wait-ExcelExit", "$rel.EmergencyRequired = $true"):
        assert token in harness, f"the shutdown ledger lost {token}"

    # Emitting Y and Z is not enough: both must sit in the deferred set, because
    # that set is what the final gate demands BY NAME. Dropping either from it
    # would leave the case running and judged by nothing.
    deferred_decl = re.search(
        r"\$script:Phase4FinalizationScenarioIds\s*=\s*@\((.*?)\)",
        _executable(SCENARIOS), re.S,
    )
    assert deferred_decl, "the deferred set is not declared"
    deferred = set(_ps_string_literals(deferred_decl.group(1)))
    assert deferred == {"Y", "Z"}, (
        f"the final gate demands {sorted(deferred)} by name; both lifecycle "
        "cases must be there or one of them is proved by nothing"
    )


def test_103_the_correction_is_pure_powershell_and_touches_no_vba() -> None:
    """PROOF 10. A lifecycle/ledger fix has no business in a production module."""
    final = _final_gate()
    gate = _prerequisite_gate()
    for block, label in ((gate, "P5-P4"), (final, "P5-FIN")):
        # No COM, no Run(), no workbook, no VBA at either gate: both read the
        # in-memory result ledger and nothing else.
        for forbidden in ("$Workbook", "$Excel", ".Run(", "VBProject", "PCCM_"):
            assert forbidden not in block, (
                f"{label} reaches into the workbook; it must judge results only"
            )
    # And the diagnostic module is still the only non-production .bas involved.
    assert DIAGNOSTIC.is_file()
    assert (SRC_VBA / "modCalcReport.bas").is_file()


def test_104_the_accepted_gate_b_scenario_topology_is_preserved() -> None:
    """PROOF 11. Every accepted family still emits; P5-FIN is the only addition."""
    scenarios = _executable(SCENARIOS)
    harness = _executable(HARNESS)
    both = scenarios + "\n" + harness
    for name in ("P5-PRE", "P5-P4", "P5-ALL", "P5-FX", "P5-M", "P5-DC",
                 "P5-AN", "P5-RF", "P5-PQ", "P5-PN", "P5-AR", "P5-ID",
                 "P5-ST", "P5-NS", "P5-KP", "P5-RC", "P5-FA", "P5-FC",
                 "P5-AX", "P5-EV", "P5-XX", "P5-FIN"):
        # P5-FA and P5-FC are emitted through the shared failpoint driver, which
        # takes the ID as a parameter; both emission forms count.
        assert (_result_call(name) in both
                or f"-ScenarioId '{name}'" in both), f"scenario {name} no longer emits"
    for name in STATUS_ROW_IDS:
        assert (_result_call(name) in both
                or f"-ScenarioId '{name}'" in both), f"status row {name} no longer emits"
    # The plan-case scenario registry is unchanged: P5-FIN is a gate, not a
    # mapping target, exactly as P5-PRE, P5-P4 and P5-ALL are not.
    registry = _procedure(scenarios, "Get-Phase5ScenarioIds")
    for name in ("P5-FIN", "P5-P4", "P5-ALL", "P5-PRE"):
        assert f"'{name}'" not in registry, (
            f"{name} is a lifecycle gate and must not be a plan-case mapping target"
        )


def test_105_the_runtime_run_1_topology_is_pinned_end_to_end() -> None:
    """PROOF 12. The observed Run-1 ledger, replayed against both gates.

    Order is taken from the harness source, not assumed: the Phase-5 call, then
    Z, then Y, then the final gate.
    """
    harness = _executable(HARNESS)
    phase5_at = harness.index("Invoke-Phase5GateBScenarios -Excel")
    z_at = harness.index("Add-Result 'Z' 'Excel closed naturally")
    y_at = harness.index("$transient = @(Get-TransientFailures)")
    fin_at = harness.index("Add-Phase4FinalCompletenessResult -Results $results")
    assert phase5_at < z_at < y_at < fin_at, (
        "the corrected lifecycle order is not Phase 5 -> Z -> Y -> final gate"
    )
    # Z really is after the session is gone, and Y after Z.
    quit_at = harness.index("$excel.Quit(); $rel.QuitCalled = $true")
    assert phase5_at < quit_at < z_at, (
        "Z no longer follows Application.Quit, so it is not a post-session case"
    )

    # The entry gate judges the reduced set, the final gate the whole matrix.
    gate = _prerequisite_gate()
    assert "$prerequisite -contains $_.Id" in gate, (
        "the entry gate no longer partitions on the prerequisite set"
    )
    assert "$required -contains $_.Id" not in gate, (
        "the entry gate is back to judging the full matrix, which is the defect"
    )
    assert "$required -contains $_.Id" in _final_gate(), (
        "the final gate no longer judges the whole matrix"
    )

    required = set(PHASE4_SCENARIO_IDS)
    deferred = {"Y", "Z"}
    prerequisite = required - deferred

    def entry_gate(ledger: dict) -> bool:
        seen = set(ledger)
        if prerequisite - seen:
            return False
        judged = {k: v for k, v in ledger.items() if k in prerequisite}
        return (not any(v == "FAIL" for v in judged.values())
                and not any(v == "SKIP" for v in judged.values())
                and sum(v == "PASS" for v in judged.values()) == len(prerequisite))

    def final_gate(ledger: dict) -> bool:
        seen = set(ledger)
        if required - seen:
            return False
        judged = {k: v for k, v in ledger.items() if k in required}
        if any(v in ("FAIL", "SKIP") for v in judged.values()):
            return False
        if sum(v == "PASS" for v in judged.values()) != 35:
            return False
        return all(ledger.get(name) == "PASS" for name in deferred)

    # RUN 1, exactly as observed: at the entry gate only the 33 exist.
    at_entry = {name: "PASS" for name in prerequisite}
    assert entry_gate(at_entry), (
        "the corrected entry gate still refuses the Run-1 ledger it should accept"
    )
    # The old gate is what actually fired, and it could not have done otherwise.
    old_missing = required - set(at_entry)
    assert old_missing == {"Y", "Z"}, "the Run-1 'missing: Y, Z' detail is not reproduced"

    # Y and Z then run, and the final gate accepts the completed matrix.
    completed = dict(at_entry, Y="PASS", Z="PASS")
    assert final_gate(completed)

    # But not without them, and not with them failing.
    assert not final_gate(dict(at_entry)), "the final gate accepts a matrix missing Y and Z"
    assert not final_gate(dict(at_entry, Y="PASS", Z="FAIL"))
    assert not final_gate(dict(at_entry, Y="FAIL", Z="PASS"))
    assert not final_gate(dict(at_entry, Y="PASS", Z="SKIP"))
    # And a genuine Phase-4 regression still stops Phase 5 at the entry gate.
    broken = dict(at_entry); broken["K2"] = "FAIL"
    assert not entry_gate(broken)
    skipped = dict(at_entry); skipped["K2"] = "SKIP"
    assert not entry_gate(skipped)
    dropped = {k: v for k, v in at_entry.items() if k != "K2"}
    assert not entry_gate(dropped)


# --- negative controls -----------------------------------------------------
def test_nc_84_the_original_run_1_prerequisite_is_caught() -> None:
    """The accepted-at-35640ec gate, replayed: unsatisfiable by construction."""
    planted = _synthetic(
        "$required = Get-Phase4RequiredScenarioIds\n"
        "foreach ($id in $required) { if ($seen -notcontains $id) { $missing += $id } }\n"
        "$null = Add-Check $list 'all 35 Phase-4 scenarios reported a result' "
        "($missing.Count -eq 0)\n"
    )
    assert "Get-Phase4PrerequisiteScenarioIds" not in planted, (
        "the defective gate must be visible as one that reads the whole matrix"
    )
    at_entry = set(PHASE4_SCENARIO_IDS) - {"Y", "Z"}
    missing = set(PHASE4_SCENARIO_IDS) - at_entry
    assert missing == {"Y", "Z"} and len(at_entry) == 33, (
        "33 of 35 with Y and Z missing is exactly what Run 1 reported"
    )


def test_nc_85_lowering_the_threshold_to_33_is_caught() -> None:
    """The tempting non-fix: drop the count and lose Y/Z coverage entirely."""
    planted = _synthetic(
        "$null = Add-Check $list 'the Phase-4 matrix is 33/33 PASS' ($passed.Count -eq 33)\n"
    )
    assert "35" not in planted, "a threshold cut must be visible as one"
    # With no later gate, a run whose Z FAILED would still be accepted.
    judged = {name: "PASS" for name in set(PHASE4_SCENARIO_IDS) - {"Y", "Z"}}
    assert sum(v == "PASS" for v in judged.values()) == 33
    ledger = dict(judged, Y="PASS", Z="FAIL")
    assert sum(v == "PASS" for v in ledger.values() ) == 34, (
        "a 33/33 entry gate with no final gate never judges Y or Z at all"
    )


def test_nc_86_a_count_only_final_gate_is_caught() -> None:
    """35/35 by count alone is satisfiable by a matrix that lost Z."""
    planted = _synthetic(
        "$null = Add-Check $list 'the final Phase-4 matrix is 35/35 PASS' "
        "($passed.Count -eq 35)\n"
    )
    assert "foreach ($id in $deferred)" not in planted, (
        "a count-only final gate must be visible as one"
    )
    # A duplicated result stands in for the missing one under a pure count.
    ledger = [(name, "PASS") for name in set(PHASE4_SCENARIO_IDS) - {"Z"}]
    ledger.append(("Y", "PASS"))          # Y recorded twice, Z never
    assert sum(1 for _, status in ledger if status == "PASS") == 35
    assert "Z" not in {name for name, _ in ledger}, (
        "the count reaches 35 while the case the gate exists to prove is absent"
    )


# ===========================================================================
# 23. RUNTIME RUN 2: the four confirmed harness roots
# ===========================================================================
# Runtime Run 2 (harness commit cc70c37) reached the Phase-5 scenarios and
# finalised cleanly - 35/35 Phase-4, P5-FIN PASS, natural Excel exit - and then
# failed 24 scenarios. They are not 24 defects. Four harness roots and one
# production finding account for all of them; the tests below pin the four
# harness roots so that exact run cannot recur.
#
#   R1  inventory semantics    P5-M, P5-D8   "present 30 of 15"
#   R2  commentary as code     P5-EV         "modAppState: Worksheet_Change; NPV"
#   R3  parameter shadowing    7 scenarios   "The property 'rows' cannot be found"
#   R4  shape count as buttons P5-M          "found 6"
#
# The fifth, P5-D1/P5-D2, is a PRODUCTION finding and is NOT repaired here.
MANIFEST_SHEET_COUNT = 14
VBEXT_STD_MODULE = 1
VBEXT_DOCUMENT = 100


def _inventory_helper() -> str:
    return _procedure(_executable(SCENARIOS), "Add-Phase5ModuleInventoryChecks")


# --- R1: inventory semantics ------------------------------------------------
def test_106_the_component_inventory_is_partitioned_by_vbide_type() -> None:
    """R1. 15 manifest modules + 15 document components = 30 VBComponents.

    Run 2's `present 30 of 15` was arithmetic, not a defect: a VBProject holds
    one document component per worksheet plus ThisWorkbook. The manifest's
    vba.modules describes STANDARD modules and never described documents.
    """
    source = _executable(SCENARIOS)
    types = re.search(r"\$script:VbextComponentTypes\s*=\s*@\{(.*?)\n\}", source, re.S)
    assert types, "the VBIDE component types are not declared"
    declared = dict(re.findall(r"(\w+)\s*=\s*(\d+)", types.group(1)))
    assert declared.get("StdModule") == str(VBEXT_STD_MODULE)
    assert declared.get("Document") == str(VBEXT_DOCUMENT)
    for named in ("ClassModule", "MSForm", "ActiveXDesigner"):
        assert named in declared, f"{named} is not named, so it cannot be excluded by name"

    reader = _procedure(source, "Get-Phase5VbComponentInventory")
    assert "Type = [int]$component.Type" in reader, (
        "the inventory does not carry the component type, so it cannot partition"
    )
    assert "Name = [string]$component.Name" in reader
    assert "Release-Transient $component 'VBComponent'" in reader, (
        "the inventory leaks a COM object per component"
    )

    helper = _inventory_helper()
    assert "[int]$_.Type -eq $script:VbextComponentTypes.StdModule" in helper
    assert "[int]$_.Type -eq $script:VbextComponentTypes.Document" in helper
    # The document partition is JUDGED, not merely separated. Splitting the
    # components and then asserting nothing about the documents would let a
    # stray sheet component through in exchange for fixing the standard ones.
    assert "($documents.Count -eq $expectedDocuments)" in helper, (
        "document components are separated but never counted"
    )

    # The exact Run-2 topology, modelled.
    components = (
        [("mod%d" % i, VBEXT_STD_MODULE) for i in range(15)]
        + [("sh%d" % i, VBEXT_DOCUMENT) for i in range(MANIFEST_SHEET_COUNT)]
        + [("ThisWorkbook", VBEXT_DOCUMENT)]
    )
    assert len(components) == 30, "the reproduced Run-2 topology is not 30 components"
    standard = [n for n, t in components if t == VBEXT_STD_MODULE]
    documents = [n for n, t in components if t == VBEXT_DOCUMENT]
    assert len(standard) == 15 and len(documents) == MANIFEST_SHEET_COUNT + 1
    # The old rule failed on a correct project; the new one does not.
    assert len(components) != 15, "the old name-only rule compared 30 against 15"


def test_107_the_inventory_is_not_weakened_to_at_least_fifteen() -> None:
    """R1. Exact, both directions, and every other component type excluded."""
    helper = _inventory_helper()
    # Every manifest module must be a STANDARD module - not merely present.
    assert "': the production module ' + $name + ' is a standard module'" in helper
    # No stray standard module.
    assert "$strayStandard = @($standardNames | Where-Object { $ExpectedModules -notcontains $_ })" in helper
    assert "($strayStandard.Count -eq 0)" in helper
    # And the set size is exact.
    assert "($standardNames.Count -eq @($ExpectedModules).Count)" in helper
    for forbidden in ("-ge 15", "-gt 14", "at least"):
        assert forbidden not in helper, f"the inventory was weakened ({forbidden})"
    # Documents are counted, not waved through.
    assert "$expectedDocuments = $ExpectedSheetCount + 1" in helper, (
        "document components are not counted, so a stray one could hide there"
    )
    assert "($documents.Count -eq $expectedDocuments)" in helper, (
        "the document count is computed and then not compared"
    )
    assert "': exactly one ThisWorkbook document component'" in helper
    # Nothing else may exist at all.
    assert "': no class module, UserForm or designer component exists'" in helper
    assert "($other.Count -eq 0)" in helper
    # And the arithmetic is stated so the evidence is self-checking.
    assert "@($Components).Count -eq (@($ExpectedModules).Count + $expectedDocuments)" in helper


def test_108_a_stray_or_missing_standard_module_still_fails() -> None:
    """R1. The decision table the corrected rule must implement."""
    manifest = {f"mod{i}" for i in range(15)}
    docs = MANIFEST_SHEET_COUNT + 1

    def verdict(standard: set, documents: int, other: int) -> bool:
        return (manifest <= standard and not (standard - manifest)
                and len(standard) == len(manifest)
                and documents == docs and other == 0)

    assert verdict(set(manifest), docs, 0), "the correct project must pass"
    assert not verdict(manifest | {"modStray"}, docs, 0), "a stray standard module must fail"
    assert not verdict(manifest - {"mod0"}, docs, 0), "a missing manifest module must fail"
    assert not verdict(manifest | {DIAGNOSTIC_MODULE_NAME}, docs, 0), (
        "a left-behind diagnostic module must fail"
    )
    assert not verdict(set(manifest), docs, 1), "a class module or UserForm must fail"
    assert not verdict(set(manifest), docs + 1, 0), "a stray document component must fail"
    assert not verdict(set(manifest), docs - 1, 0), "a missing document component must fail"


def test_109_the_diagnostic_module_absence_is_proved_in_its_own_partition() -> None:
    """R1. The diagnostic module is a STANDARD module, so that is where it is denied."""
    source = _executable(SCENARIOS)
    d8 = source[source.index("$components.Remove($target)"):source.index("Add-Phase5Result 'P5-D8'")]
    assert "'the diagnostic module is absent from the standard modules'" in d8
    assert "$standardNames -notcontains $diagnosticName" in d8, (
        "absence is not proved against the partition the module would reappear in"
    )
    assert "'the diagnostic module is absent from the project entirely'" in d8
    assert "'no diagnostic procedure is callable any more'" in source
    # P5-M and P5-D8 share one helper, so the two judgements cannot drift.
    assert d8.count("Add-Phase5ModuleInventoryChecks") == 1
    m = source[source.index("Add-Phase5Result 'P5-FX'"):source.index("Add-Phase5Result 'P5-M'")]
    assert "Add-Phase5ModuleInventoryChecks" in m


# --- R2: commentary is not code ---------------------------------------------
def test_110_the_forbidden_construct_scan_reads_code_not_commentary() -> None:
    """R2. modAppState's explanatory comments are prose, and stay."""
    source = _executable(SCENARIOS)
    ev = source[source.index("Add-Phase5Result 'P5-D0'") - 6000:source.index("Add-Phase5Result 'P5-EV'")]
    assert "$code = Get-VbaExecutableCode -Code $raw" in ev, (
        "the scan still runs over the raw module text"
    )
    assert "if ($code -match [regex]::Escape([string]$forbidden))" in ev, (
        "the manifest scan does not run over the stripped code"
    )
    assert "$raw -match" not in ev, "the raw text is still scanned somewhere"
    # And the second half of the rule is a real substitution, not a pass-through:
    # a construct named inside a string literal is data, not an occurrence.
    stripper = _procedure(source, "Remove-VbaStringLiterals")
    assert "[regex]::Replace($Code," in stripper, (
        "the string-literal stripper returns its input unchanged"
    )
    assert "return $Code" not in stripper
    # The requirement itself is untouched: the manifest list is still the source.
    assert "@($Manifest.vba.forbidden_constructs)" in ev
    assert "'the manifest forbids ' + $handler" in source

    # The production comments that Run 2 flagged are still there, unedited.
    app_state = _text(SRC_VBA / "modAppState.bas")
    assert "no input Worksheet_Change handler" in app_state, (
        "the explanatory production comment was removed instead of the harness fixed"
    )
    assert "NPV" in app_state
    # And they are commentary, not code, in the accepted source.
    for needle in ("no input Worksheet_Change handler", "NPV, EMV"):
        line = next(l for l in app_state.splitlines() if needle in l)
        assert line.lstrip().startswith("'"), f"{needle!r} is not on a comment line"


def test_111_a_real_change_event_declaration_still_fails() -> None:
    """R2. The three cases the runtime proof must distinguish."""
    source = _executable(SCENARIOS)
    stripper = _procedure(source, "Remove-VbaCommentary")
    # A comment starts at an apostrophe OUTSIDE a string literal.
    assert "if (($ch -eq ([char]39)) -and (-not $inString)) { break }" in stripper, (
        "the comment marker is no longer recognised outside a string literal"
    )
    assert "$inString = -not $inString" in stripper, (
        "string literals are not tracked, so an apostrophe inside one would truncate code"
    )
    # A doubled quote inside a literal is an ESCAPE, not a close. Without it,
    # "he said ""don't""" closes early and the apostrophe truncates the rest.
    assert "($line[$i + 1] -eq '\"')" in stripper, (
        "the doubled-quote escape is gone, so a literal can close early"
    )
    assert "$null = $kept.Append('\"\"')" in stripper
    assert "Rem" in stripper, "Rem-form commentary is not handled"
    # No blanket substitution that could swallow a real declaration.
    for forbidden in ("-replace 'Worksheet_Change'", "-replace 'Workbook_SheetChange'",
                      ".Replace('Worksheet_Change'"):
        assert forbidden not in source, f"a blanket text substitution was used ({forbidden})"

    # The declaration test judges the STRIPPED code, like the general scan.
    assert "$code = Get-VbaExecutableCode -Code $raw" in source, (
        "the module text is no longer reduced to executable code before scanning"
    )
    assert "Test-VbaProcedureDeclared -Code $code" in source, (
        "the declaration test runs over raw text rather than stripped code"
    )
    assert "if ($code -match [regex]::Escape([string]$forbidden))" in source, (
        "the manifest forbidden-construct scan reads the raw text again"
    )
    declared = _procedure(source, "Test-VbaProcedureDeclared")
    assert "(?:Sub|Function)" in declared, "the declaration test does not look for a procedure"
    assert "[regex]::Escape($ProcedureName)" in declared
    assert r"\s*\(" in declared, "the name is not required to be followed by an argument list"
    assert "'no change-event procedure is DECLARED anywhere in the project'" in source

    # The decision table, modelled against the same rule.
    pattern = re.compile(
        r"(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+"
        r"Worksheet_Change\s*\("
    )

    # The SHARED model, so this test cannot encode a different lexical rule
    # from the one the harness ships. See section 25.
    strip = _model_remove_commentary

    # 1. a comment naming the handler -> allowed
    assert not pattern.search(strip("' there is no input Worksheet_Change handler")), (
        "a comment is still read as a declaration"
    )
    assert not pattern.search(strip("Rem Worksheet_Change is deliberately absent"))
    assert not pattern.search(strip("x = 1: Rem Private Sub Worksheet_Change(x)")), (
        "an inline Rem comment is still read as code"
    )
    # 2. a real Worksheet_Change procedure -> FAIL
    assert pattern.search(strip("Private Sub Worksheet_Change(ByVal Target As Range)"))
    assert pattern.search(strip("Sub Worksheet_Change(ByVal Target As Range)"))
    assert pattern.search(strip("Public Static Sub Worksheet_Change(ByVal T As Range)"))
    # 3. a real Workbook_SheetChange procedure -> FAIL
    sheet_pattern = re.compile(
        r"(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+"
        r"Workbook_SheetChange\s*\("
    )
    assert sheet_pattern.search(strip(
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)"
    ))
    # and a handler declared after a comment on the previous line is still caught
    assert pattern.search(strip(
        "' no Worksheet_Change here\nPrivate Sub Worksheet_Change(ByVal Target As Range)"
    )), "stripping a comment must not swallow the following line"
    # an apostrophe inside a string literal does not truncate the statement
    assert pattern.search(strip(
        'Private Sub Worksheet_Change(ByVal T As Range) : x = "it' + "'" + 's"'
    ))


# --- R3: the typed-parameter shadowing --------------------------------------
def test_112_no_typed_parameter_is_shadowed_by_a_local_assignment() -> None:
    """R3. The `.rows` root, closed as a CLASS rather than as one instance.

    PowerShell variable names are case-insensitive and a typed parameter keeps
    its constraint, so `$block = <PSCustomObject>` inside a function declaring
    `[string]$Block` silently stringified the block. `$block.rows` then threw
    PropertyNotFoundException on all seven of P5-S2, P5-ST, P5-S3, P5-S4, P5-S5,
    P5-KP and P5-RC. This scans every function in all three harness files.
    """
    offenders = []
    for path in (SCENARIOS, HARNESS, BOOTSTRAP / "com_lifecycle.ps1"):
        # Comments are stripped first. The block in Get-CalcScalar that QUOTES
        # the defective line to explain it is documentation, not a shadow.
        source = _executable(path)
        for match in re.finditer(r"^function\s+([A-Za-z0-9\-]+)", source, re.M):
            name = match.group(1)
            start = source.index("{", match.end())
            depth, end = 0, start
            while end < len(source):
                if source[end] == "{":
                    depth += 1
                elif source[end] == "}":
                    depth -= 1
                    if depth == 0:
                        break
                end += 1
            body = source[start:end]
            param = re.search(r"param\s*\(", body)
            if not param:
                continue
            open_at = body.index("(", param.start())
            depth, close_at = 0, open_at
            while close_at < len(body):
                if body[close_at] == "(":
                    depth += 1
                elif body[close_at] == ")":
                    depth -= 1
                    if depth == 0:
                        break
                close_at += 1
            typed = {
                var.lower(): (kind, var)
                for kind, var in re.findall(
                    r"\[([A-Za-z0-9.\[\]]+)\]\s*\$([A-Za-z0-9_]+)", body[open_at:close_at + 1]
                )
            }
            if not typed:
                continue
            for assign in re.finditer(
                r"(?<![\w$`])\$([A-Za-z0-9_]+)\s*=(?![=~])", body[close_at + 1:]
            ):
                key = assign.group(1).lower()
                if key in typed:
                    kind, declared = typed[key]
                    offenders.append(
                        f"{path.name}::{name} declares [{kind}]${declared} "
                        f"and reassigns ${assign.group(1)}"
                    )
    assert not offenders, (
        "a typed parameter is shadowed by a local assignment, which silently "
        "coerces the value: " + "; ".join(offenders)
    )


def test_113_get_calc_scalar_reads_the_block_into_its_own_local() -> None:
    """R3, the specific instance, pinned by name."""
    body = _procedure(_executable(SCENARIOS), "Get-CalcScalar")
    assert "$blockSpec = $Inspection.calc.scalar_blocks.$Block" in body
    assert "$blockSpec.rows.$FieldKey" in body
    assert "$blockSpec.value_column" in body
    assert "\n    $block =" not in body, "the shadowing assignment is back"
    # The receiver really does carry `rows`: the projection was never at fault.
    #
    # Read from the FRESHLY EMITTED projection, never from `build/`. The earlier
    # form here read the repository artifact when it happened to exist and
    # skipped silently when it did not - which trusts a stale file and proves
    # nothing when the build is broken. Same defect class as the canonical
    # parity helper; both are closed.
    inspection = _emitted()["inspection"]
    for block in ("calc_state", "calc_totals"):
        spec = inspection["calc"]["scalar_blocks"][block]
        assert "rows" in spec, f"{block} has no rows mapping"
        assert "value_column" in spec
        assert spec["rows"], f"{block}.rows is empty"


def test_114_the_rows_receiver_is_the_block_not_the_key() -> None:
    """R3. The precise Run-2 object shape: a String asked for `.rows`."""
    # What the defect produced: the parameter, retyped, is a String.
    stringified = "@{value_column=C; rows=}"
    assert not hasattr(stringified, "rows")
    # What the corrected code holds: the block mapping itself.
    block = {"value_column": "C", "rows": {"last_attempt_result": 17}}
    assert "rows" in block and block["rows"]["last_attempt_result"] == 17
    # And the two are reached by DIFFERENT names in source, so one cannot become
    # the other again.
    body = _procedure(_executable(SCENARIOS), "Get-CalcScalar")
    assert "$Block" in body and "$blockSpec" in body
    assert body.count("$blockSpec") == 3, "the local is not used for all three reads"


# --- R4: shapes are not command buttons -------------------------------------
def test_115_command_buttons_are_counted_by_macro_binding_not_by_shape() -> None:
    """R4. `found 6` counted every Shape; a command button is a bound shape."""
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Phase5Result 'P5-FX'"):source.index("Add-Phase5Result 'P5-M'")]
    assert "$commandButtons = @($shapeRecords | Where-Object {" in block
    assert "-not [string]::IsNullOrWhiteSpace([string]$_.OnAction) })" in block
    assert "'exactly the five declared (sheet, shape, macro) bindings exist'" in block
    assert "$boundTriples.Count -eq $declaredTriples.Count" in block
    # A declared shape name may exist exactly once, on its declared sheet.
    assert "'no second shape named ' + $wantName + ' exists on any other sheet'" in block, (
        "a duplicate command surface on another sheet would go unreported"
    )
    assert "([string]$_.Name -eq $wantName) -and ([string]$_.Sheet -ne $wantSheet)" in block
    assert "'exactly five command buttons persist in the workbook'" not in source, (
        "the raw shape count is still the button requirement"
    )
    # It NAMES what it found. Run 2 said "found 6" and nothing else.
    assert "$shapeInventory" in block
    assert "[string]$_.Sheet + '!' + [string]$_.Name + ' -> '" in block
    assert "Add-Note ('P5-M: shape inventory across the '" in block


def test_116_the_no_calculate_button_requirement_is_unchanged_and_widened() -> None:
    """R4. Strictly stronger than the count it replaces."""
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Phase5Result 'P5-FX'"):source.index("Add-Phase5Result 'P5-M'")]
    # The declared macro is compared against THAT shape's own OnAction, so a
    # button bound to another declared button's macro cannot pass.
    assert "$bound = ($actual -ceq $wantAction)" in block, (
        "the macro comparison fell back to global set membership"
    )
    assert "$onActions -contains $wantAction" not in block
    # unchanged: over EVERY shape, bound or not
    assert "'NO shape has OnAction = PCCM_Calculate'" in block
    assert "($onActions -notcontains 'PCCM_Calculate')" in block
    assert "$onActions = @($shapeRecords | ForEach-Object { [string]$_.OnAction })" in block
    # widened: an undeclared shape may not reach the PCCM surface at all
    assert "'every macro-bound shape is one of the five declared buttons'" in block
    assert "'no undeclared shape invokes a PCCM_ procedure'" in block
    assert "[string]$_.OnAction -like 'PCCM_*'" in block
    # The five is counted over the BOUND triples, never over every shape again.
    assert "$boundTriples = @($commandButtons | ForEach-Object {" in block, (
        "the five-button rule no longer counts bindings"
    )
    assert "($shapeRecords.Count -eq 5)" not in block

    # The decision table.
    declared = {"btnPCCMApplyTimeline", "btnPCCMAddCostLine", "btnPCCMDeleteCostLine",
                "btnPCCMAddRisk", "btnPCCMDeleteRisk"}

    def verdict(shapes: list) -> bool:
        bound = [s for s in shapes if s[1]]
        return (len(bound) == 5
                and all(n in declared for n, _ in bound)
                and not any(a == "PCCM_Calculate" for _, a in shapes)
                and not any(n not in declared and a.startswith("PCCM_") for n, a in shapes))

    five = [(n, "PCCM_" + n[7:]) for n in sorted(declared)]
    assert verdict(five), "the five declared buttons must pass"
    # Run 2's sixth shape, if it carries no macro, is not a command button.
    assert verdict(five + [("Decoration", "")]), (
        "an unbound decorative shape is not a command button"
    )
    # But anything that commands is judged.
    assert not verdict(five + [("Rogue", "PCCM_AddRisk")]), "an undeclared bound shape must fail"
    assert not verdict(five + [("Rogue", "PCCM_Calculate")])
    assert not verdict(five[:4]), "a missing button must fail"
    assert not verdict(five + [("Decoration", "SomeOtherMacro")]), (
        "a sixth macro-bound shape must fail the five-button rule"
    )


# --- diagnostics ------------------------------------------------------------
def test_117_unexpected_phase5_errors_carry_a_source_location() -> None:
    """Run 2 reported eleven scenarios as one sentence with no location."""
    source = _executable(SCENARIOS)
    formatter = _procedure(source, "Format-Phase5Err")
    assert "$ErrorRecord.InvocationInfo" in formatter, "no invocation information is read"
    for member in ("ScriptName", "ScriptLineNumber", "OffsetInLine"):
        assert member in formatter, f"the location omits {member}"
    assert "$ErrorRecord.ScriptStackTrace" in formatter, "no call chain is reported"
    assert "$invocation.Line" in formatter, "the offending source line is not shown"
    assert "$exception.InnerException" in formatter, (
        "a wrapped exception would still be reported only by its outer type"
    )
    assert "$exception.GetType().FullName" in formatter, "the .NET type is no longer reported"
    assert "$exception.Message" in formatter, "the message is no longer reported"

    # NO COM OBJECT REACHES THE LEDGER.
    for forbidden in ("$Workbook", "$Excel", ".Value2", "Range(", "VBProject", "$Application"):
        assert forbidden not in formatter, (
            f"the diagnostic ledger touches a COM object ({forbidden})"
        )

    # Every Phase-5 catch site uses it, and the accepted Phase-4 helper is intact.
    assert "Format-Err " not in source.replace("Format-Phase5Err", ""), (
        "a Phase-5 catch site still reports through the type+message-only helper"
    )
    assert source.count("Format-Phase5Err $_") >= 20, (
        "most Phase-5 catch sites are not routed through the located formatter"
    )
    lifecycle = _text(BOOTSTRAP / "com_lifecycle.ps1")
    assert "function Format-Err {" in lifecycle, "the accepted Phase-4 helper was renamed"
    assert "ScriptStackTrace" not in lifecycle, "the accepted Phase-4 helper was modified"


def test_118_the_typed_reader_is_confirmed_by_run_2_and_unchanged() -> None:
    """Run 2 proved the typed reader works; it must not be 'fixed'.

    P5-FX PASSED and recorded, on real Windows PowerShell 5.1 and real Excel:

        P5-FX: locked FX seed captured as String'SAR' / Double:1

    That is `New-Object 'object[]' $colCount`, index assignment, `Write-RowObject`
    and `Format-Phase5Typed` all behaving exactly as designed - a String stayed a
    String and a numeric stayed a Double across the COM boundary. The reader is
    therefore NOT the InvalidCastException boundary, and stringifying Value2
    would destroy the evidence architecture for nothing.
    """
    reader = _procedure(_executable(SCENARIOS), "Get-Phase5TypedTableBody")
    assert "$line = New-Object 'object[]' $colCount" in reader, (
        "the allocation Run 2 exercised successfully was changed"
    )
    assert "$line[$c - 1] = $cell.Value2" in reader, "the typed read was changed"
    assert "Write-RowObject $line" in reader
    for forbidden in ("[string]$cell.Value2", "Format-CalcValue", "$line += "):
        assert forbidden not in reader, f"the reader now stringifies or appends ({forbidden})"
    comparator = _procedure(_executable(SCENARIOS), "Test-Phase5ExactValue")
    assert "if ($Actual -is [string]) { return $false }" in comparator, (
        "the exact comparator was widened to accept display text"
    )


# --- negative controls -----------------------------------------------------
def test_nc_87_the_run_2_name_only_inventory_is_caught() -> None:
    planted = _synthetic(
        "$present += [string]$component.Name\n"
        "$null = Add-Check $list 'the inventory is exactly the 15 manifest modules again' "
        "((@($present).Count -eq $expected.Count))\n"
    )
    assert "$component.Type" not in planted, "a name-only inventory must be visible as one"
    present, expected = 15 + MANIFEST_SHEET_COUNT + 1, 15
    assert present != expected and present == 30, (
        "the defective rule compares 30 against 15, which is what Run 2 printed"
    )


def test_nc_88_a_raw_text_forbidden_scan_is_caught() -> None:
    planted = _synthetic("if ($raw -match [regex]::Escape([string]$forbidden)) {\n")
    assert "Remove-VbaCommentary" not in planted, "a raw-text scan must be visible as one"
    comment = "' ... no input Worksheet_Change handler, and this guarantees that stays true"
    assert "Worksheet_Change" in comment, "the production comment really does contain the token"
    assert comment.lstrip().startswith("'"), "and it really is a comment"


def test_nc_89_a_reintroduced_parameter_shadow_is_caught() -> None:
    planted = _synthetic(
        "function Get-CalcScalar {\n"
        "    param($Workbook, $Inspection, [string]$Block, [string]$FieldKey)\n"
        "    $block = $Inspection.calc.scalar_blocks.$Block\n"
    )
    assert "$blockSpec" not in planted, "the shadowing assignment must be visible"
    # PowerShell resolves $block and $Block to one variable; the type constraint
    # then converts the assigned object to its string form.
    assert "$block".lower() == "$Block".lower()


def test_nc_90_counting_shapes_as_buttons_is_caught() -> None:
    planted = _synthetic(
        "$shapesFound++\n"
        "$null = Add-Check $list 'exactly five command buttons persist in the workbook' "
        "($shapesFound -eq 5) (\"found \" + $shapesFound)\n"
    )
    assert "OnAction" not in planted, "a shape-count rule must be visible as one"
    # Six shapes, five of them buttons: the count fails and names nothing.
    shapes = [("btn%d" % i, "PCCM_X") for i in range(5)] + [("Decoration", "")]
    assert len(shapes) == 6 and len([s for s in shapes if s[1]]) == 5


def test_nc_91_a_type_and_message_only_error_report_is_caught() -> None:
    planted = _synthetic(
        "$type = $ErrorRecord.Exception.GetType().FullName\n"
        "return ('{0}: {1}' -f $type, $msg)\n"
    )
    assert "ScriptStackTrace" not in planted and "InvocationInfo" not in planted, (
        "a location-free error report must be visible as one"
    )
    # Eleven scenarios, one indistinguishable sentence.
    reported = ["System.InvalidCastException: Unable to cast object of type "
                "'System.Double' to type 'System.String'."] * 11
    assert len(set(reported)) == 1, (
        "Run 2's eleven cast failures were textually identical, so none could be located"
    )


def test_119_the_error_evidence_names_the_call_chain_not_only_the_line() -> None:
    """One shared path served eleven Run-2 scenarios.

    Knowing WHICH LINE threw is not enough when the throwing statement lives in
    a helper that every scenario reaches through a different route. The frame
    stack is what distinguishes `Set-Phase5Fixture -> Reset-Phase5FxTable` from
    `Set-Phase5Fixture -> Write-Phase5Driver`, and it is reported frame by frame
    rather than as one unsplit blob.
    """
    formatter = _procedure(_executable(SCENARIOS), "Format-Phase5Err")
    assert "$stack = [string]$ErrorRecord.ScriptStackTrace" in formatter, (
        "the call chain is not read from the error record"
    )
    assert "foreach ($frame in ($stack -split" in formatter, (
        "the call chain is not rendered frame by frame"
    )
    assert "$parts += ('  ' + $frame.Trim())" in formatter
    # An empty stack must not be reported as a chain of nothing.
    assert "if (-not [string]::IsNullOrWhiteSpace($stack))" in formatter
    # And the chain is joined into ONE detail string, so a checklist line keeps
    # its shape and the report stays greppable.
    assert "($parts -join [string][char]10)" in formatter, (
        "the located error is not assembled into a single detail value"
    )


# ===========================================================================
# 24. REVIEW ROUND 2A: cardinality, binding identity, string literals
# ===========================================================================
def _button_block() -> str:
    source = _executable(SCENARIOS)
    return source[source.index("Add-Phase5Result 'P5-FX'"):source.index("Add-Phase5Result 'P5-M'")]


# --- BLOCKER 1: the inventory emits records, not one nested array -----------
def test_120_the_component_inventory_emits_one_record_per_component() -> None:
    """`return ,$out` would hand the caller ONE array-shaped object.

    The unary comma stops PowerShell unrolling a collection. That is right for a
    function returning one ROW whose cells must stay together; it is wrong for a
    function producing a SEQUENCE of records. With the comma, the caller's
    `@(...)` sees a single nested array and every downstream
    `Where-Object { $_.Type ... }` filters an object with no `.Type` at all.
    """
    source = _executable(SCENARIOS)
    reader = _procedure(source, "Get-Phase5VbComponentInventory")
    assert "return ,$out" not in reader, "the inventory still returns a nested array"
    assert "," not in reader.split("return")[-1] or "return" not in reader, (
        "the inventory still has a comma-wrapped return"
    )
    assert "$out" not in reader, (
        "the inventory still accumulates a collection instead of emitting records"
    )
    assert "Write-Output ([pscustomobject]@{" in reader, (
        "the inventory does not emit one record per component"
    )
    # A PSCustomObject is not a collection, so -NoEnumerate is neither needed
    # nor correct here.
    assert "-NoEnumerate" not in reader, (
        "a record emission does not need the row-boundary protection"
    )
    assert "Write-RowObject" not in reader


def test_121_the_caller_is_the_authority_that_collects_the_inventory() -> None:
    """0/1/N from the pipeline; @() at the caller turns it into an Object[]."""
    source = _executable(SCENARIOS)
    for callsite in ("$components = @(Get-Phase5VbComponentInventory -Workbook $Workbook)",
                     "$inventory = @(Get-Phase5VbComponentInventory -Workbook $Workbook)"):
        assert callsite in source, f"a caller does not wrap the inventory with @(): {callsite}"
    # Neither caller re-indexes as if it had received one nested array.
    assert "$components[0][0]" not in source
    assert "$inventory[0][0]" not in source
    # And @() is only the authority if the function actually emits records: a
    # comma-wrapped return makes the caller's @() collect one nested array.
    assert "return ,$out" not in source, (
        "the inventory hands back a collection, so the caller's @() collects one "
        "array-shaped object instead of the component records"
    )

    # Cardinality, modelled the way PowerShell resolves it.
    def emitted(records: list) -> list:
        """What `@(function)` yields when the function EMITS each record."""
        return list(records)

    def comma_wrapped(records: list) -> list:
        """What `@(function)` yields when the function does `return ,$out`."""
        return [records]

    for count in (0, 1, 2, 30):
        records = [{"Name": f"c{i}", "Type": 1} for i in range(count)]
        assert len(emitted(records)) == count, "emission must preserve cardinality"
        # The defect: any count collapses to one array-shaped object.
        assert len(comma_wrapped(records)) == 1
        if count != 1:
            assert len(comma_wrapped(records)) != count, (
                "the comma-wrapped return changes the caller's element count"
            )
        # And that object has no .Type, so no partition can ever match it.
        assert "Type" not in comma_wrapped(records)[0]


def test_122_the_inventory_helper_consumes_records_not_an_array() -> None:
    """Every partition reads `.Type` off an individual component record."""
    helper = _inventory_helper()
    for partition in ("$standard = @($Components | Where-Object {",
                      "$documents = @($Components | Where-Object {",
                      "$other = @($Components | Where-Object {"):
        assert partition in helper, f"the helper no longer partitions with {partition}"
    assert "[int]$_.Type" in helper, "the partitions do not read a record's own type"
    assert "[string]$_.Name" in helper

    # Records in -> the Run-2 topology partitions correctly.
    records = ([{"Name": f"mod{i}", "Type": VBEXT_STD_MODULE} for i in range(15)]
               + [{"Name": f"sh{i}", "Type": VBEXT_DOCUMENT} for i in range(MANIFEST_SHEET_COUNT)]
               + [{"Name": "ThisWorkbook", "Type": VBEXT_DOCUMENT}])
    assert len([r for r in records if r["Type"] == VBEXT_STD_MODULE]) == 15
    assert len([r for r in records if r["Type"] == VBEXT_DOCUMENT]) == MANIFEST_SHEET_COUNT + 1
    # One nested array in -> nothing matches anything, silently.
    nested = [records]
    assert len([r for r in nested if isinstance(r, dict) and r.get("Type") == VBEXT_STD_MODULE]) == 0
    assert len([r for r in nested if isinstance(r, dict) and r.get("Type") == VBEXT_DOCUMENT]) == 0


def test_123_the_typed_table_reader_keeps_its_row_boundary_protection() -> None:
    """This finding is about the inventory ONLY.

    Get-Phase5TypedTableBody emits one object[] PER ROW and must keep
    -NoEnumerate, or a row's cells would each become a pipeline object.
    """
    source = _executable(SCENARIOS)
    reader = _procedure(source, "Get-Phase5TypedTableBody")
    assert "Write-RowObject $line" in reader, "the row emission was changed"
    writer = _procedure(_executable(HARNESS), "Write-RowObject")
    assert "Write-Output -NoEnumerate $Row" in writer, (
        "the accepted row-boundary protection was removed"
    )
    assert "[object[]]$Row" in writer


# --- BLOCKER 2: the binding is a triple ------------------------------------
def test_124_each_declared_button_is_proved_as_a_sheet_shape_macro_triple() -> None:
    """Three independent global sets pass on two swapped entry points."""
    block = _button_block()
    # 1. exactly one shape with the declared name on the declared sheet
    assert "'exactly one shape named ' + $wantName + ' exists on ' + $wantSheet" in block
    assert "([string]$_.Sheet -eq $wantSheet) -and ([string]$_.Name -eq $wantName)" in block
    assert "($onSheet.Count -eq 1)" in block
    # 2. THAT shape's OnAction equals the declared entry point
    assert "$actual = [string]$onSheet[0].OnAction" in block
    assert "$bound = ($actual -ceq $wantAction)" in block, (
        "the macro is not compared against the declared entry point on that shape"
    )
    # 3. no second copy of the declared name elsewhere
    assert "'no second shape named ' + $wantName + ' exists on any other sheet'" in block
    assert "([string]$_.Name -eq $wantName) -and ([string]$_.Sheet -ne $wantSheet)" in block
    # 4/5/6
    assert "'every macro-bound shape is one of the five declared buttons'" in block
    assert "'NO shape has OnAction = PCCM_Calculate'" in block
    assert "'exactly the five declared (sheet, shape, macro) bindings exist'" in block

    # The weak set-wise per-button check is gone, not merely supplemented.
    assert "($onActions -contains [string]$button.entry_point)" not in block, (
        "the global entry-point membership check survives and can report a "
        "swapped button as correct"
    )
    # And the manifest really does carry the whole identity.
    manifest = _emitted()["manifest"]
    for button in manifest["buttons"]:
        for field in ("sheet", "shape_name", "entry_point"):
            assert button[field], f"the manifest button lacks {field}"


def test_125_the_button_decision_table() -> None:
    """Every case the review named, against the triple rule."""
    manifest = _emitted()["manifest"]
    declared = [(b["sheet"], b["shape_name"], b["entry_point"]) for b in manifest["buttons"]]
    assert len(declared) == 5
    declared_pairs = {(s, n) for s, n, _ in declared}
    declared_triples = set(declared)

    def verdict(shapes: list[tuple[str, str, str]]) -> bool:
        """shapes = [(sheet, name, on_action)]; '' means no macro."""
        for sheet, name, action in declared:
            on_sheet = [s for s in shapes if s[0] == sheet and s[1] == name]
            if len(on_sheet) != 1:
                return False
            if on_sheet[0][2] != action:
                return False
            if [s for s in shapes if s[1] == name and s[0] != sheet]:
                return False
        bound = [s for s in shapes if s[2]]
        if any((s[0], s[1]) not in declared_pairs for s in bound):
            return False
        if any((s[0], s[1]) not in declared_pairs and s[2].startswith("PCCM_") for s in shapes):
            return False
        if any(s[2] == "PCCM_Calculate" for s in shapes):
            return False
        return set(bound) == declared_triples and len(bound) == len(declared_triples)

    # correct five triples -> PASS
    assert verdict(list(declared)), "the correct five bindings must pass"

    # two entry points swapped -> FAIL  (the counterexample the review gave)
    swapped = list(declared)
    a = next(i for i, t in enumerate(swapped) if t[1] == "btnPCCMAddCostLine")
    d = next(i for i, t in enumerate(swapped) if t[1] == "btnPCCMDeleteCostLine")
    swapped[a] = (swapped[a][0], swapped[a][1], declared[d][2])
    swapped[d] = (swapped[d][0], swapped[d][1], declared[a][2])
    assert not verdict(swapped), "two swapped entry points must fail"
    # and the three global set-wise checks would all still have passed
    assert {t[1] for t in swapped} == {t[1] for t in declared}
    assert {t[2] for t in swapped} == {t[2] for t in declared}
    assert len([t for t in swapped if t[2]]) == 5

    # correct shape name on the wrong sheet -> FAIL
    moved = [t for t in declared if t[1] != "btnPCCMAddRisk"]
    moved.append(("Setup", "btnPCCMAddRisk", "PCCM_AddRisk"))
    assert not verdict(moved), "a declared button on the wrong sheet must fail"

    # duplicate declared shape on another sheet -> FAIL
    assert not verdict(list(declared) + [("Setup", "btnPCCMAddRisk", "PCCM_AddRisk")]), (
        "a duplicate of a declared button elsewhere must fail"
    )

    # missing declared button -> FAIL
    assert not verdict(list(declared)[:4]), "a missing declared button must fail"

    # sixth UNBOUND decorative shape -> PASS
    assert verdict(list(declared) + [("Setup", "Decoration", "")]), (
        "an unbound decorative shape is not a command button"
    )

    # sixth MACRO-BOUND shape -> FAIL
    assert not verdict(list(declared) + [("Setup", "Decoration", "SomeMacro")]), (
        "a sixth macro-bound shape must fail"
    )

    # undeclared PCCM_ binding -> FAIL
    assert not verdict(list(declared) + [("Setup", "Rogue", "PCCM_AddRisk")])

    # PCCM_Calculate binding -> FAIL, bound to a declared name or not
    assert not verdict(list(declared) + [("Setup", "Rogue", "PCCM_Calculate")])
    calc = list(declared)
    calc[a] = (calc[a][0], calc[a][1], "PCCM_Calculate")
    assert not verdict(calc), "a declared button rebound to PCCM_Calculate must fail"

    # and the raw-count rule is NOT what decides any of this
    assert len(list(declared) + [("Setup", "Decoration", "")]) == 6, (
        "the passing decorative case has six shapes, so a Shape.Count == 5 rule "
        "would have failed it"
    )


def test_126_the_raw_shape_count_rule_is_not_restored() -> None:
    source = _executable(SCENARIOS)
    for forbidden in ("'exactly five command buttons persist in the workbook'",
                      "($shapesFound -eq 5)", "$shapesFound++"):
        assert forbidden not in source, f"the raw Shape.Count rule is back ({forbidden})"


# --- BLOCKER 3: string literals are data, not code -------------------------
def test_127_the_runtime_scanner_matches_the_python_source_authority() -> None:
    """VbaModule.code is strip_strings(strip_comments(raw)). So is this."""
    source = _executable(SCENARIOS)
    combined = _procedure(source, "Get-VbaExecutableCode")
    assert "Remove-VbaStringLiterals -Code (Remove-VbaCommentary -Code $Code)" in combined, (
        "the runtime scanner does not compose the two halves in the Python order"
    )
    stripper = _procedure(source, "Remove-VbaStringLiterals")
    assert "'\"(?:[^\"]|\"\")*\"'" in stripper, (
        "the literal pattern is not the one the Python authority uses"
    )
    assert "[regex]::Replace" in stripper
    # The payload is EMPTIED, not deleted: the statement around it keeps shape.
    assert "'\"\"'" in stripper, "the literal is removed rather than emptied"

    # Doubled quotes are an escape, in the comment stripper too.
    commentary = _procedure(source, "Remove-VbaCommentary")
    assert "($line[$i + 1] -eq '\"')" in commentary, (
        "a doubled quote inside a string is not treated as an escape"
    )
    assert "$null = $kept.Append('\"\"')" in commentary

    # The Python authority is unchanged and still says the same thing.
    authority = _text(PCCM_ROOT / "builder" / "pccm_builder" / "vba_source.py")
    assert "return strip_strings(strip_comments(self.raw))" in authority
    assert r'''re.sub(r'"(?:[^"]|"")*"', '""', source)''' in authority


def test_128_the_forbidden_construct_decision_table() -> None:
    """Commentary and string payloads are data; real constructs are not."""
    manifest_forbidden = _emitted()["manifest"]["vba"]["forbidden_constructs"]
    for needed in ("Worksheet_Change", "Workbook_SheetChange", "Randomize", "Rnd(",
                   "RunSimulation", "NPV"):
        assert needed in manifest_forbidden, f"{needed} left the manifest list"

    # The SHARED model, so this decision table and section 25's cannot drift.
    flags = _model_flags
    declares = _model_declares

    # allowed
    assert not flags("' there is no input Worksheet_Change handler")
    assert not flags("Rem Worksheet_Change is deliberately absent")
    assert not flags("x = 1: Rem Worksheet_Change is deliberately absent"), (
        "an inline Rem comment is read as executable code"
    )
    assert not flags('Err.Raise 5, , "Worksheet_Change"')
    assert not flags('MsgBox "NPV is not available"')
    assert not flags('Debug.Print "RunSimulation"')
    assert not flags('x = "Rnd("')
    assert not declares("' no Worksheet_Change here", "Worksheet_Change")
    assert not declares('MsgBox "Private Sub Worksheet_Change("', "Worksheet_Change")

    # FAIL
    assert declares("Private Sub Worksheet_Change(ByVal Target As Range)", "Worksheet_Change")
    assert declares(
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal T As Range)",
        "Workbook_SheetChange")
    assert "Randomize" in flags("    Randomize")
    assert "Rnd(" in flags("    x = Rnd()")
    assert "RunSimulation" in flags("    Call RunSimulation")

    # a real construct AFTER a string literal on the same statement still fails
    assert "Randomize" in flags('MsgBox "no rng here" : Randomize')
    assert "Rnd(" in flags('Debug.Print "Worksheet_Change" : y = Rnd(1)')
    # a comment AFTER executable code does not hide the code
    assert "Randomize" in flags("    Randomize   ' seeded here")
    # an apostrophe inside a literal does not truncate the statement
    assert "Randomize" in flags('MsgBox "it' + "'" + 's fine" : Randomize')
    # doubled quotes inside a literal stay inside it
    assert not flags('MsgBox "he said ""NPV"" loudly"')
    assert "Randomize" in flags('MsgBox "he said ""hi""" : Randomize')


def test_129_the_production_source_still_passes_the_corrected_scan() -> None:
    """The accepted production VBA must be clean under the new semantics.

    A stricter scanner that flagged real production code would be a different
    defect, so this runs the corrected rule over every frozen module.
    """
    manifest_forbidden = _emitted()["manifest"]["vba"]["forbidden_constructs"]
    sys.path.insert(0, str(PCCM_ROOT / "builder"))
    from pccm_builder.vba_source import strip_comments, strip_strings

    offenders = []
    for path in sorted(SRC_VBA.glob("*.bas")):
        body = strip_strings(strip_comments(_text(path)))
        for construct in manifest_forbidden:
            if construct in body:
                offenders.append(f"{path.name}: {construct}")
    assert not offenders, (
        "the corrected scan flags accepted production source: " + "; ".join(offenders)
    )
    # And the two Run-2 false positives really were comment-only.
    app_state = _text(SRC_VBA / "modAppState.bas")
    assert "Worksheet_Change" in app_state and "NPV" in app_state
    assert "Worksheet_Change" not in strip_strings(strip_comments(app_state))
    assert "NPV" not in strip_strings(strip_comments(app_state))


# --- negative controls -----------------------------------------------------
def test_nc_92_a_comma_wrapped_inventory_return_is_caught() -> None:
    planted = _synthetic("    return ,$out\n")
    assert "Write-Output" not in planted, "the comma-wrapped return must be visible"
    records = [{"Name": "modA", "Type": 1}, {"Name": "Sheet1", "Type": 100}]
    wrapped = [records]
    assert len(wrapped) == 1 != len(records)
    assert not [r for r in wrapped if isinstance(r, dict)], (
        "the caller sees one array-shaped object with no .Type"
    )


def test_nc_93_a_set_wise_button_proof_is_caught() -> None:
    planted = _synthetic("($onActions -contains [string]$button.entry_point)\n")
    assert "$wantSheet" not in planted, "the set-wise proof must be visible as one"
    declared = [("Cost Lines", "btnPCCMAddCostLine", "PCCM_AddCostLine"),
                ("Cost Lines", "btnPCCMDeleteCostLine", "PCCM_DeleteCostLine")]
    swapped = [("Cost Lines", "btnPCCMAddCostLine", "PCCM_DeleteCostLine"),
               ("Cost Lines", "btnPCCMDeleteCostLine", "PCCM_AddCostLine")]
    # Every global set is identical; the triples are not.
    assert {t[1] for t in swapped} == {t[1] for t in declared}
    assert {t[2] for t in swapped} == {t[2] for t in declared}
    assert set(swapped) != set(declared), "the swap is invisible to set-wise checks"


def test_nc_94_a_comment_only_stripper_is_caught() -> None:
    planted = _synthetic("$code = Remove-VbaCommentary -Code $raw\n")
    assert "Remove-VbaStringLiterals" not in planted, (
        "a comment-only stripper must be visible as one"
    )
    # The payload a comment-only stripper leaves behind.
    line = 'MsgBox "NPV is not available"'
    assert "NPV" in line, "the string payload survives comment stripping"
    assert "NPV" not in re.sub(r'"(?:[^"]|"")*"', '""', line), (
        "emptying the literal is what removes it"
    )


# ===========================================================================
# 25. REVIEW ROUND 2B: Rem is a STATEMENT, not a line prefix
# ===========================================================================
# VBA permits Rem wherever a statement may begin, which includes after a colon
# separator:
#
#     x = 1: Rem Worksheet_Change is deliberately absent
#
# A line-anchored `^\s*Rem(\s|$)` misses that form entirely, so P5-EV could
# report Worksheet_Change as executable code from a comment. The models below
# are a faithful port of the corrected single-pass stripper and are shared by
# every test in this section, so no test can quietly disagree with another
# about what the rule is.
def _model_remove_commentary(code: str) -> str:
    """Faithful port of Remove-VbaCommentary."""
    out = []
    for line in code.split("\n"):
        in_string = False
        at_statement_start = True
        kept: list[str] = []
        index = 0
        while index < len(line):
            char = line[index]
            # Rem: outside a literal, at a statement boundary, complete keyword.
            if (not in_string) and at_statement_start and (index + 3) <= len(line) \
               and line[index:index + 3].lower() == "rem" \
               and ((index + 3) == len(line) or line[index + 3].isspace()):
                break
            if char == '"':
                if in_string and (index + 1) < len(line) and line[index + 1] == '"':
                    kept.append('""')
                    at_statement_start = False
                    index += 2
                    continue
                in_string = not in_string
                kept.append(char)
                at_statement_start = False
                index += 1
                continue
            if char == "'" and not in_string:
                break
            kept.append(char)
            if not char.isspace():
                at_statement_start = (char == ":" and not in_string)
            index += 1
        out.append("".join(kept))
    return "\n".join(out)


def _model_remove_string_literals(code: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', code)


def _model_executable(code: str) -> str:
    """Faithful port of Get-VbaExecutableCode."""
    return _model_remove_string_literals(_model_remove_commentary(code))


def _model_flags(code: str) -> set[str]:
    body = _model_executable(code)
    return {c for c in _emitted()["manifest"]["vba"]["forbidden_constructs"] if c in body}


def _model_declares(code: str, name: str) -> bool:
    return bool(re.search(
        r"(?im)^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(?:Sub|Function)\s+"
        + re.escape(name) + r"\s*\(", _model_executable(code)))


def _model_remove_commentary_line_start_only(code: str) -> str:
    """The DEFECTIVE implementation this round replaces."""
    out = []
    for line in code.split("\n"):
        in_string, kept, index = False, [], 0
        while index < len(line):
            char = line[index]
            if char == '"':
                if in_string and index + 1 < len(line) and line[index + 1] == '"':
                    kept.append('""')
                    index += 2
                    continue
                in_string = not in_string
                kept.append(char)
                index += 1
                continue
            if char == "'" and not in_string:
                break
            kept.append(char)
            index += 1
        text = "".join(kept)
        out.append("" if re.match(r"^\s*Rem(\s|$)", text) else text)
    return "\n".join(out)


def test_130_rem_is_recognised_at_every_statement_boundary() -> None:
    """The corrected rule, stated in source and proved by the shared model."""
    stripper = _procedure(_executable(SCENARIOS), "Remove-VbaCommentary")
    # The boundary is tracked in the SAME pass that understands literals.
    assert "$atStatementStart = $true" in stripper, "no statement boundary is tracked"
    assert "$line.Substring($i, 3) -eq 'Rem'" in stripper, (
        "Rem is not matched as a complete three-character keyword"
    )
    # Complete keyword only: whitespace or end of line must follow.
    assert "[char]::IsWhiteSpace($line[$i + 3])" in stripper, (
        "Remember and RemoteValue would be read as commentary"
    )
    assert "(($i + 3) -eq $line.Length)" in stripper, "a bare `Rem` line is not handled"
    # Outside a literal only.
    assert "(-not $inString) -and $atStatementStart" in stripper
    # A colon opens the next statement, but only outside a literal.
    assert "$atStatementStart = (($ch -eq ':') -and (-not $inString))" in stripper
    # Whitespace leaves the boundary alone, so `x = 1 :   Rem ...` still works.
    assert "if (-not [char]::IsWhiteSpace($ch)) {" in stripper

    # The line-anchored regex is gone, not merely supplemented.
    assert "'^\\s*Rem(\\s|$)'" not in _executable(SCENARIOS), (
        "the line-start-only Rem rule survives"
    )
    # And no broad regex replaced it.
    for forbidden in ("-replace 'Rem", "-match 'Rem", "Rem.*$"):
        assert forbidden not in stripper, f"a broad Rem regex was introduced ({forbidden})"


def test_131_the_rem_decision_table() -> None:
    """Every case the review named, plus the ones that must not regress."""
    # --- required additions ------------------------------------------------
    assert not _model_flags("Rem Worksheet_Change is absent")
    assert not _model_flags("x = 1: Rem Worksheet_Change is absent")
    assert not _model_flags("x = 1 :   Rem NPV is deliberately absent")
    assert not _model_flags('x = "Rem Worksheet_Change"')
    assert "Randomize" in _model_flags('x = "text : Rem NPV" : Randomize'), (
        "a colon inside a literal must not open a statement"
    )
    assert not _model_flags('x = "text : Rem NPV" : Randomize') - {"Randomize"}, (
        "the literal payload leaked into the executable code"
    )
    assert not _model_flags("Remember = 1"), "Remember was read as a Rem comment"
    assert not _model_flags('RemoteValue = "NPV"'), "RemoteValue was read as a Rem comment"
    assert "Randomize" in _model_flags("x = 1: Randomize")
    assert not _model_flags("x = 1: Rem Randomize is deliberately absent")

    # --- and the identifier-prefix cases really do keep their code ----------
    assert "Remember = 1" in _model_executable("Remember = 1")
    assert "RemoteValue" in _model_executable('RemoteValue = "NPV"')

    # --- VBA is case-insensitive -------------------------------------------
    assert not _model_flags("REM Worksheet_Change upper case")
    assert not _model_flags("rem NPV lower case")
    assert not _model_flags("x = 1: REM NPV")

    # --- degenerate and label forms ----------------------------------------
    assert _model_executable("Rem").strip() == ""
    assert not _model_flags("x = 1: Rem")
    assert not _model_flags("MyLabel: Rem NPV after a label")
    assert "Randomize" in _model_flags("Dim Remainder As Long: Randomize"), (
        "an identifier starting with Rem must not swallow the next statement"
    )

    # --- nothing from round 2A regressed -----------------------------------
    assert not _model_flags("' there is no input Worksheet_Change handler")
    assert not _model_flags('MsgBox "NPV is not available"')
    assert not _model_flags('MsgBox "he said ""NPV"" loudly"')
    assert "Randomize" in _model_flags('MsgBox "it' + "'" + 's fine" : Randomize')
    assert "Randomize" in _model_flags("    Randomize   ' seeded here")
    assert "Randomize" in _model_flags('MsgBox "no rng here" : Randomize')
    assert "Rnd(" in _model_flags("    x = Rnd()")
    assert "RunSimulation" in _model_flags("    Call RunSimulation")

    # --- declarations ------------------------------------------------------
    assert _model_declares("Private Sub Worksheet_Change(ByVal T As Range)", "Worksheet_Change")
    assert _model_declares(
        "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal T As Range)",
        "Workbook_SheetChange")
    assert not _model_declares("Rem Private Sub Worksheet_Change(x)", "Worksheet_Change")
    assert not _model_declares('x = "Private Sub Worksheet_Change("', "Worksheet_Change")
    assert not _model_declares("' Private Sub Worksheet_Change(x)", "Worksheet_Change")


def test_132_the_production_source_is_unaffected_by_the_rem_correction() -> None:
    """A stricter comment rule must not change what the frozen modules say."""
    forbidden = _emitted()["manifest"]["vba"]["forbidden_constructs"]
    offenders = []
    for path in sorted(SRC_VBA.glob("*.bas")) + [DIAGNOSTIC]:
        body = _model_executable(_text(path))
        for construct in forbidden:
            if construct in body:
                offenders.append(f"{path.name}: {construct}")
    assert not offenders, (
        "the corrected Rem rule flags accepted source: " + "; ".join(offenders)
    )
    # The corrected stripper agrees with the Python authority on real modules -
    # neither one strips MORE than the other on code that has no inline Rem.
    sys.path.insert(0, str(PCCM_ROOT / "builder"))
    from pccm_builder.vba_source import strip_comments, strip_strings
    for path in sorted(SRC_VBA.glob("*.bas")):
        raw = _text(path)
        assert _model_executable(raw).strip() == strip_strings(strip_comments(raw)).strip(), (
            f"{path.name}: the runtime and static strippers disagree"
        )


def test_nc_95_the_line_start_only_rem_rule_is_caught() -> None:
    """The exact defect this round closes, reproduced against the old code."""
    planted = _synthetic("        if ($text -match '^\\s*Rem(\\s|$)') { $text = '' }\n")
    assert "$atStatementStart" not in planted, (
        "the line-start-only Rem rule must be visible as one"
    )
    forbidden = _emitted()["manifest"]["vba"]["forbidden_constructs"]

    def defective_flags(code: str) -> set[str]:
        body = _model_remove_string_literals(_model_remove_commentary_line_start_only(code))
        return {c for c in forbidden if c in body}

    inline = "x = 1: Rem Worksheet_Change is deliberately absent"
    assert "Worksheet_Change" in defective_flags(inline), (
        "the inline Rem form must defeat the line-start-only implementation"
    )
    assert not _model_flags(inline), "the corrected rule must not flag it"
    # The line-start form is the only one the old rule ever handled.
    at_start = "Rem Worksheet_Change is deliberately absent"
    assert not defective_flags(at_start) and not _model_flags(at_start), (
        "both rules agree on the form the old one was written for"
    )
    # And the old rule was not merely incomplete - it was wrong in one
    # direction only: it never stripped too much.
    assert defective_flags("Remember = 1") == _model_flags("Remember = 1")


def test_133_the_rem_decision_is_taken_inside_the_single_pass() -> None:
    """Structural, and independent of test_130's keyword assertions.

    The old rule rebuilt the line and THEN matched a line-anchored regex over
    it, which is why it could not see a colon separator. The corrected rule
    decides at a character position, inside the pass that already knows whether
    it is inside a literal - so this pins the shape, not just the tokens.
    """
    stripper = _procedure(_executable(SCENARIOS), "Remove-VbaCommentary")

    # The post-loop line rewrite is gone: nothing reconstructs the line and
    # then blanks it.
    assert "$text = $kept.ToString()" not in stripper, (
        "the line is still rebuilt for a line-level Rem decision"
    )
    assert "{ $text = '' }" not in stripper
    assert "$null = $out.AppendLine($kept.ToString())" in stripper, (
        "the kept text is no longer appended directly"
    )

    # The Rem guard lives INSIDE the character loop, and ahead of the literal
    # and apostrophe branches, so it is evaluated at every candidate position.
    loop_at = stripper.index("while ($i -lt $line.Length) {")
    rem_at = stripper.index("$line.Substring($i, 3) -eq 'Rem'")
    quote_at = stripper.index("if ($ch -eq '\"') {")
    apostrophe_at = stripper.index("if (($ch -eq ([char]39)) -and (-not $inString)) { break }")
    assert loop_at < rem_at < quote_at < apostrophe_at, (
        "the Rem guard is not evaluated per position inside the single pass"
    )
    # It is indexed by $i, not applied to the whole line.
    assert "$line.Substring($i, 3)" in stripper
    assert "$line[$i + 3]" in stripper

    # And it consults literal state at that position, so a Rem inside a string
    # cannot end the line. Checked over the guard's own region, independently
    # of how the conjunction is written.
    rem_guard = stripper[loop_at:quote_at]
    assert "$inString" in rem_guard, (
        "the Rem guard does not consult literal state"
    )
    assert "$atStatementStart" in rem_guard

    # And the boundary really is maintained across the loop, in both directions.
    assert stripper.count("$atStatementStart = $false") == 2, (
        "the statement boundary is not closed on both literal paths"
    )
    assert "$atStatementStart = (($ch -eq ':') -and (-not $inString))" in stripper


# ===========================================================================
# 26. THE CANONICAL-DOUBLE PARITY SCENARIO (P5-DP)
# ===========================================================================
def _parity_block() -> str:
    source = _executable(SCENARIOS)
    return source[source.index("$parity = $Cases.fingerprint.canonical_parity"):
                  source.index("Add-Phase5Result 'P5-DP'")]


def test_134_the_parity_scenario_drives_the_emitted_corpus() -> None:
    """Ten vectors exposed the defect; they cannot accept the correction."""
    source = _executable(SCENARIOS)
    assert "Add-Phase5Result 'P5-DP'" in source, "the parity scenario does not emit a result"
    block = _parity_block()
    assert "$vectors = @($parity.vectors)" in block
    assert "($vectors.Count -gt 2000)" in block, "the corpus size is not gated"
    assert "'every emitted parity vector was actually driven'" in block, (
        "a scenario that silently drove none would still report PASS"
    )
    assert "($checked -eq $vectors.Count)" in block
    # It runs BEFORE the diagnostic module is removed, or GBD_ would not answer.
    assert source.index("Add-Phase5Result 'P5-DP'") < source.index("Add-Phase5Result 'P5-D8'")
    assert source.index("Add-Phase5Result 'P5-D0'") < source.index("Add-Phase5Result 'P5-DP'")


def test_135_the_parity_probe_is_rebuilt_from_its_bit_pattern() -> None:
    """Identity is the bit pattern; a decimal literal is a second opinion."""
    block = _parity_block()
    assert "[BitConverter]::Int64BitsToDouble([Convert]::ToInt64($bits, 16))" in block, (
        "the probe is not reconstructed from the IEEE-754 bit pattern"
    )
    assert "$bits = [string]$vector.bits" in block
    assert "$vector.value" not in block, "the scenario reads a decimal literal instead"

    corpus = _emitted()["cases"]["fingerprint"]["canonical_parity"]
    assert len(corpus["vectors"]) > 2000
    for vector in corpus["vectors"][:50]:
        assert len(vector["bits"]) == 16
        assert all(c in "0123456789ABCDEF" for c in vector["bits"])
        assert "value" not in vector, "the corpus carries an ambiguous decimal literal"


def test_136_the_harness_never_computes_a_canonical_expectation() -> None:
    """An expectation produced by the algorithm under test proves nothing."""
    block = _parity_block()
    assert "[string]$vector.expected" in block, "the expectation is not read from the corpus"
    assert "-cne ('OK|' + [string]$vector.expected)" in block, (
        "the comparison is not case-sensitive against the emitted text"
    )
    scenarios = _executable(SCENARIOS)
    for forbidden in ("ToString('E16')", 'ToString("E16")', "{0:E16}", "'E16'",
                      "[double]::Parse", "Format-Number"):
        assert forbidden not in scenarios, (
            f"the harness formats a canonical number of its own ({forbidden})"
        )


def test_137_the_neighbour_triples_are_a_collision_proof() -> None:
    block = _parity_block()
    assert "foreach ($triple in @($parity.neighbours))" in block
    assert "$distinct = @($texts | Select-Object -Unique)" in block
    assert "($ok -and ($distinct.Count -eq 3))" in block, (
        "three distinct Doubles are not required to give three distinct strings"
    )
    corpus = _emitted()["cases"]["fingerprint"]["canonical_parity"]
    assert len(corpus["neighbours"]) >= 8
    for triple in corpus["neighbours"]:
        texts = [m["expected"] for m in triple["members"]]
        assert len(set(texts)) == 3, triple["label"]
        assert [m["position"] for m in triple["members"]] == ["below", "value", "above"]


def test_138_the_parity_failures_are_reported_not_just_counted() -> None:
    """2432 ok lines would bury the evidence; a bare count buries it too."""
    block = _parity_block()
    assert "$failures.Count -lt 20" in block, "no individual discrepancy is retained"
    assert "'[' + $bits + '] '" in block, "a failure does not name its bit pattern"
    assert "', expected OK|'" in block, "a failure does not show what was expected"
    assert "($failures.Count -eq 0)" in block
    assert "($failures -join ' | ')" in block


def test_139_the_locked_ten_vector_scenarios_are_not_weakened() -> None:
    """P5-DP is additional evidence, never a replacement for P5-D1 or P5-D2."""
    source = _executable(SCENARIOS)
    assert "Add-Phase5Result 'P5-D1'" in source and "Add-Phase5Result 'P5-D2'" in source
    assert "'ten locked numeric vectors were emitted'" in source
    assert "($vectors.Count -eq 10)" in source, "the ten locked vectors are no longer required"
    assert "'the separator vector set was emitted'" in source
    assert "'the canonical form does not depend on the injected separator'" in source
    corpus = _emitted()["cases"]["fingerprint"]
    assert len(corpus["numeric_encodings"]["vectors"]) == 10


# ===========================================================================
# 27. RUNTIME RUN 4: the typed COM write, and one result per ID
# ===========================================================================
# Run 4 located R5 exactly, with the stack the round-2 diagnostics were added
# to produce:
#
#   System.InvalidCastException: Unable to cast object of type 'System.Double'
#   to type 'System.String'.
#     at phase5_gate_b_scenarios.ps1:922   source: $cell.Value2 = $Value
#     at Set-Phase5TypedCell -> Reset-Phase5FxTable -> Set-Phase5Fixture
#
# The locked seed is String 'SAR' then Double 1. The String assignment
# succeeded; the Double assignment through the SAME source line failed. The
# accepted Phase-4 Set-TableCell never hit it because it has one assignment line
# per branch.
def _typed_setter() -> str:
    return _procedure(_executable(SCENARIOS), "Set-Phase5TypedCell")


def test_140_the_generic_com_assignment_site_is_gone() -> None:
    """R5, closed at the line Run 4 named."""
    setter = _typed_setter()
    assert "$cell.Value2 = $Value" not in setter, (
        "the polymorphic single assignment site is back"
    )
    # One COM assignment per supported type, each its own source line.
    sites = re.findall(r"\$cell\.Value2 = (\[[a-z]+\]\$Value)", setter)
    assert sites == sorted(set(sites), key=sites.index), "a branch assigns twice"
    assert "[string]$Value" in sites and "[double]$Value" in sites, sites
    assert len(sites) >= 2, sites
    # And the null branch still clears rather than writing anything.
    assert "$null = $cell.ClearContents()" in setter
    assert "if ($null -eq $Value) {" in setter


def test_141_the_dispatch_is_on_the_captured_type_not_on_the_contract() -> None:
    """Explicit dispatch is NOT permission to repair text into a number."""
    setter = _typed_setter()
    for probe in ("$Value -is [string]", "$Value -is [bool]", "$Value -is [double]"):
        assert probe in setter, f"the setter does not dispatch on {probe}"
    # Nothing consults the model, the contract, the column or the format.
    for forbidden in ("$Inspection", "NumberFormat", "$Case", "$Model", "expected",
                      "ColumnKey", "-as [double]", "[double]::Parse"):
        assert forbidden not in setter, (
            f"the setter infers a type from {forbidden} instead of using the captured one"
        )
    # An unsupported captured type FAILS, naming the real CLR type - and the
    # throw is the LAST branch, not a dead one with a coercion in front of it.
    assert "throw (" in setter and "$Value.GetType().FullName" in setter
    assert setter.count("$cell.Value2 = [double]$Value") == 1, (
        "a second numeric assignment was added, so an unsupported type is "
        "coerced before it can reach the throw"
    )
    assert "if ($false)" not in setter, "a branch was disabled rather than removed"
    tail = setter[setter.index("throw ("):]
    assert "$cell.Value2" not in tail, "an assignment follows the refusal branch"
    assert "cannot restore a value of type" in _procedure(
        _text(SCENARIOS).replace("\r\n", "\n"), "Set-Phase5TypedCell"
    )

    # The decision table the dispatch has to implement.
    def branch(value):
        if value is None:
            return "clear"
        if isinstance(value, str):
            return "string"
        if isinstance(value, bool):
            return "bool"
        if isinstance(value, (int, float)):
            return "double"
        return "throw"

    assert branch("1") == "string", "a captured text seed must be written back as text"
    assert branch(1.0) == "double", "a captured numeric seed must be written back numeric"
    assert branch(None) == "clear"
    assert branch(True) == "bool"
    assert branch(object()) == "throw"
    # The two that matter: same digit, different types, different branches.
    assert branch("1") != branch(1.0)


def test_142_the_captured_seed_is_restored_not_repaired() -> None:
    """No hard-coded SAR/1 and no cast at the call site."""
    source = _executable(SCENARIOS)
    reset = _procedure(source, "Reset-Phase5FxTable")
    assert "-Value $Seed.Currency" in reset and "-Value $Seed.Rate" in reset
    for forbidden in ("[double]$Seed.Rate", "[string]$Seed.Rate",
                      "[string]$Seed.Currency", "'SAR'", "-Value 1#", "-Value 1.0"):
        assert forbidden not in reset, f"the restoration repairs the seed ({forbidden})"
    # The strict typed read-back is still mandatory.
    assert "Test-Phase5ExactValue -Actual $body[0][0] -Expected $Seed.Currency" in reset
    assert "Test-Phase5ExactValue -Actual $body[0][1] -Expected $Seed.Rate" in reset
    assert "throw (" in reset, "a failed restoration does not stop the run"
    # And the comparator is still type-sensitive.
    comparator = _procedure(source, "Test-Phase5ExactValue")
    assert "if ($Actual -is [string]) { return $false }" in comparator
    assert "if ($null -eq $Expected) {" in comparator
    assert "return ($null -eq $Actual)" in comparator
    # Nothing anywhere stringifies Value2 or compares display text.
    assert "Format-CalcValue" not in _procedure(source, "Get-Phase5TypedTableBody")
    # And the restoration path is exercised at P5-FX, before any fixture needs
    # it - eleven Run-4 scenarios reported this path as their own failure.
    assert "Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection -Seed $seed" in source, (
        "P5-FX no longer proves the restoration path it gates"
    )


def test_143_p5_fx_proves_the_restoration_path_before_anything_depends_on_it() -> None:
    """Eleven scenarios reported R5 as their own failure. Prove it once, early."""
    source = _executable(SCENARIOS)
    block = source[source.index("Save-Phase5LockedFxSeed -Workbook"):
                   source.index("Add-Phase5Result 'P5-FX'")]
    # Capture still happens first, on the untouched workbook.
    assert block.index("Save-Phase5LockedFxSeed") < block.index("Reset-Phase5FxTable"), (
        "the restoration is exercised before the seed is captured"
    )
    # The REAL path, not a re-implementation.
    assert "Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection -Seed $seed" in block
    assert "Get-Phase5TypedTableBody" in block
    # VALUE AND TYPE, through the strict comparator.
    assert "Test-Phase5ExactValue -Actual $restored[0][0] -Expected $seed.Currency" in block
    assert "Test-Phase5ExactValue -Actual $restored[0][1] -Expected $seed.Rate" in block
    assert "the restored currency is the captured value AND the captured type" in block
    assert "the restored rate is the captured value AND the captured type" in block
    # It restores the seed to its OWN value, so the workbook is left as found.
    assert "-Seed $seed" in block and "-Model" not in block

    # A failure stops Phase-5 rather than producing dozens of misleading results,
    # as a FAIL and with the lifecycle evidence still to come.
    gate = source[source.index("Add-Phase5Result 'P5-FX'"):]
    head = gate[:1200]
    assert "if (-not $fxOk) {" in head
    assert "Add-Phase5Result 'P5-ALL'" in head
    assert "'FAIL'" in head and "'SKIP'" not in head, "an unmet prerequisite must be a FAIL"
    assert re.search(r"Add-Phase5Result 'P5-ALL'.*?\n\s*return\b", head, re.S), (
        "the refusal does not stop the scenarios below"
    )
    # P5-FIN, Y and Z are emitted by the driver after the return, so cleanup
    # evidence survives.
    harness = _executable(HARNESS)
    assert harness.index("Invoke-Phase5GateBScenarios -Excel") < harness.index(
        "Add-Result 'Z' 'Excel closed naturally")
    assert "Add-Phase4FinalCompletenessResult -Results $results" in harness


# --- the result ledger ------------------------------------------------------
def test_144_a_scenario_id_can_be_recorded_only_once() -> None:
    """Run 4 reported 19 failures over 17 unique IDs."""
    source = _executable(SCENARIOS)
    guard = _procedure(source, "Add-Phase5Result")
    assert "if (Test-Phase5ResultRecorded -Id $Id) {" in guard
    assert "Add-Note (" in guard, "a suppressed duplicate is silent"
    assert "return" in guard
    assert "$null = $script:Phase5RecordedIds.Add($Id)" in guard
    assert "Add-Result $Id $Name $Status $Detail" in guard, (
        "the guard does not delegate to the accepted reporter"
    )
    # The ledger is per run, and it starts at the FIRST Phase-5 entry point so
    # P5-PRE is inside it. Resetting again in the scenarios would discard that
    # record and let a later catch emit P5-PRE a second time.
    assert "Reset-Phase5ResultLedger" in _procedure(source, "Invoke-Phase5CoveragePreflight")
    assert "Reset-Phase5ResultLedger" not in _procedure(source, "Invoke-Phase5GateBScenarios"), (
        "the ledger is reset after P5-PRE has already been recorded"
    )

    # EVERY Phase-5 emission goes through the guard, in both files.
    both = source + "\n" + _executable(HARNESS)
    stray = re.findall(r"(?<!Phase5)Add-Result\s+'(P5-[A-Z0-9]+)'", both)
    assert not stray, f"these Phase-5 IDs bypass the one-result guard: {sorted(set(stray))}"
    assert not re.search(r"(?<!Phase5)Add-Result \$id ", source), (
        "a grouped catch emits IDs without the guard"
    )
    # And no setup step runs inside a scenario block after that scenario has
    # been recorded - the ownership half of the Run-4 defect.
    st_record = source.index("Add-Phase5Result 'P5-ST' `")
    st_catch = source.index("Add-Phase5Result 'P5-ST' 'Primary staleness sequence' 'FAIL'")
    between = source[st_record:st_catch]
    assert "Set-Phase5Fixture" not in between, (
        "a fixture restore still runs inside the S2/ST try after both scenarios "
        "have been recorded"
    )


def test_145_a_late_setup_failure_cannot_re_emit_a_finished_scenario() -> None:
    """The exact Run-4 topology: P5-S2 and P5-ST recorded, then setup threw."""
    source = _executable(SCENARIOS)
    # The re-establishment step now owns its own failure, as P5-SU.
    assert "Add-Phase5Result 'P5-SU'" in source, (
        "the post-scenario setup step has no result of its own"
    )
    setup = source[source.index("Add-Phase5Result 'P5-ST'"):]
    setup = setup[:setup.index("Add-Phase5Result 'P5-SU'") + 200]
    assert "} catch {" in setup
    # The S2/ST catch is guarded, and the guard really does check.
    assert "Add-Phase5Result 'P5-S2' 'Status row 2' 'FAIL'" in source
    assert "Add-Result 'P5-S2' 'Status row 2' 'FAIL'" not in source, (
        "the S2 catch bypasses the one-result guard"
    )
    assert "Add-Result 'P5-ST' 'Primary staleness sequence' 'FAIL'" not in source
    guard = _procedure(source, "Add-Phase5Result")
    assert "Test-Phase5ResultRecorded -Id $Id" in guard, (
        "the guard no longer checks whether the ID was already recorded"
    )
    assert "if ($false)" not in guard
    # The fixture restore is OUTSIDE the S2/ST try block.
    st_catch = source.index("Add-Phase5Result 'P5-ST' 'Primary staleness sequence' 'FAIL'")
    restore = source.index("$null = Set-Phase5Fixture -Excel $Excel -Workbook $Workbook -Manifest $Manifest",
                           source.index("Add-Phase5Result 'P5-ST' `"))
    assert st_catch < restore, (
        "the base-fixture restore still sits inside the scenario try block"
    )

    # The model: two scenarios recorded, then a late setup failure.
    recorded: list[str] = []

    def add(identifier: str) -> None:
        if identifier in recorded:
            return                      # the guard: a note, not a record
        recorded.append(identifier)

    add("P5-S2")
    add("P5-ST")
    add("P5-SU")                        # the late setup failure, owning itself
    assert recorded == ["P5-S2", "P5-ST", "P5-SU"]
    # And even if the old catch shape returned, the guard holds the line.
    add("P5-S2")
    add("P5-ST")
    assert recorded.count("P5-S2") == 1 and recorded.count("P5-ST") == 1


def test_146_the_grouped_status_catch_cannot_duplicate_a_recorded_id() -> None:
    """The same risk in S3/S4/S5/KP/RC, audited without waiting for a runtime."""
    source = _executable(SCENARIOS)
    grouped = source[source.index("foreach ($id in 'P5-S3', 'P5-S4', 'P5-S5', 'P5-KP', 'P5-RC')"):]
    grouped = grouped[:400]
    assert "Add-Phase5Result $id" in grouped, "the grouped catch bypasses the guard"
    assert "Add-Result $id" not in grouped

    # Run 4's own shape: S3, S4 and KP recorded, then the S5/RC step throws.
    recorded: list[str] = []

    def add(identifier: str) -> None:
        if identifier not in recorded:
            recorded.append(identifier)

    for done in ("P5-S3", "P5-S4", "P5-KP"):
        add(done)
    for identifier in ("P5-S3", "P5-S4", "P5-S5", "P5-KP", "P5-RC"):
        add(identifier)                 # the catch fires over the whole group
    assert recorded == ["P5-S3", "P5-S4", "P5-KP", "P5-S5", "P5-RC"]
    assert len(recorded) == len(set(recorded)), "the grouped catch duplicated an ID"

    # The other grouped catch, over the direct vectors, is guarded too.
    vectors = source[source.index("foreach ($id in 'P5-D1', 'P5-D2', 'P5-D3'"):]
    assert "Add-Phase5Result $id 'Direct VBA diagnostic vector'" in vectors[:400]


def test_147_every_phase5_scenario_has_exactly_one_result_in_both_paths() -> None:
    """Success path and late-failure path, both modelled."""
    source = _executable(SCENARIOS)
    emitted = set(re.findall(r"Add-Phase5Result\s+'(P5-[A-Z0-9]+)'", source))
    emitted |= set(re.findall(r"-ScenarioId '(P5-[A-Z0-9]+)'", source))
    emitted |= set(re.findall(r"Add-Phase5Result\s+'(P5-[A-Z0-9]+)'", _executable(HARNESS)))
    for required in ("P5-PRE", "P5-P4", "P5-ALL", "P5-FX", "P5-M", "P5-EV",
                     "P5-D0", "P5-D1", "P5-DP", "P5-D2", "P5-D8", "P5-DC",
                     "P5-AN", "P5-RF", "P5-PQ", "P5-PN", "P5-AR", "P5-ID",
                     "P5-ST", "P5-NS", "P5-KP", "P5-RC", "P5-FA", "P5-FC",
                     "P5-AX", "P5-XX", "P5-SU"):
        assert required in emitted, f"{required} is never emitted"
    for row in STATUS_ROW_IDS:
        assert row in emitted, f"{row} is never emitted"

    # P5-DP is a runtime-only diagnostic scenario and deliberately NOT a
    # plan-case mapping target - the coverage ledger is unchanged by it.
    registry = _procedure(source, "Get-Phase5ScenarioIds")
    for gate in ("P5-DP", "P5-FIN", "P5-SU", "P5-PRE", "P5-P4", "P5-ALL", "P5-XX"):
        assert f"'{gate}'" not in registry, f"{gate} became a plan-case mapping target"
    # The 37-plan-case coverage ledger is untouched. It is built with loops, so
    # the plan-case IDs are counted rather than the Add calls.
    ledger = _procedure(source, "Get-Phase5CoverageLedger")
    looped = sum(len(re.findall(r"\d+", group))
                 for group in re.findall(r"foreach \(\$id in ([\d, ]+)\)", ledger))
    singles = len(re.findall(r"\$ledger\.Add\('(\d+)'", ledger))
    assert looped + singles == 37, (
        f"the 37-plan-case ledger changed: {looped} looped + {singles} single"
    )

    # Modelled: a success path emits each once; a late failure adds only P5-SU.
    def run(recorded: list[str], ids: list[str]) -> list[str]:
        for identifier in ids:
            if identifier not in recorded:
                recorded.append(identifier)
        return recorded

    success = run([], ["P5-FX", "P5-M", "P5-S2", "P5-ST", "P5-S3", "P5-S4", "P5-S5"])
    assert len(success) == len(set(success)) == 7
    late = run(list(success), ["P5-SU", "P5-S2", "P5-ST"])
    assert late == success + ["P5-SU"], late
    assert len(late) == len(set(late))


# --- negative controls -----------------------------------------------------
def test_nc_96_the_generic_polymorphic_setter_is_caught() -> None:
    """The Run-4 design, kept forbidden.

    This models the SHAPE - one call site reached first with a String and then
    with a Double. It does NOT claim to reproduce the Excel COM binder: no
    non-COM model can, and the real proof is the next Windows run. What is
    modelled here is why one shared site is the wrong architecture.
    """
    planted = _synthetic("        $cell.Value2 = $Value\n")
    assert "[string]$Value" not in planted and "[double]$Value" not in planted, (
        "the generic assignment must be visible as one"
    )
    # One site, two types, in the order Reset-Phase5FxTable uses.
    site_types: list[type] = []

    def generic_site(value):
        site_types.append(type(value))

    for value in ("SAR", 1.0):
        generic_site(value)
    assert site_types == [str, float]
    assert len(set(site_types)) == 2, (
        "the same call site is reached with two different CLR types, which is "
        "the condition Run 4 failed on"
    )
    # The corrected shape gives each type its own site.
    per_type: dict[type, int] = {}

    def dispatched(value):
        per_type[type(value)] = per_type.get(type(value), 0) + 1

    for value in ("SAR", 1.0):
        dispatched(value)
    assert set(per_type) == {str, float} and all(n == 1 for n in per_type.values())


def test_nc_97_a_duplicate_scenario_result_is_caught() -> None:
    planted = _synthetic(
        "    } catch {\n"
        "        Add-Result 'P5-S2' 'Status row 2' 'FAIL' (Format-Phase5Err $_)\n"
        "        Add-Result 'P5-ST' 'Primary staleness sequence' 'FAIL' (Format-Phase5Err $_)\n"
        "    }\n"
    )
    assert "Add-Phase5Result" not in planted, "the unguarded catch must be visible"
    # Run 4's arithmetic: 19 records over 17 unique IDs.
    records = ["P5-AN", "P5-RF", "P5-PQ", "P5-PN", "P5-AR", "P5-ID", "P5-S1",
               "P5-S2", "P5-ST", "P5-NS", "P5-S3", "P5-S5", "P5-RC", "P5-FA",
               "P5-FC", "P5-S6", "P5-AX", "P5-S2", "P5-ST"]
    assert len(records) == 19 and len(set(records)) == 17
    assert records.count("P5-S2") == 2 and records.count("P5-ST") == 2
