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
    declares no PCCM_ endpoint, is imported only after P5-CMP has proved the
    whole production project compiles
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
import shutil
import sys
import tempfile
from contextlib import contextmanager
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
HARNESS = BOOTSTRAP / "phase4_functional_test.ps1"
SCENARIOS = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"
LIFECYCLE = BOOTSTRAP / "com_lifecycle.ps1"
DIAGNOSTIC = BOOTSTRAP / "phase5_gate_b_diagnostics.bas"
BUILD_STAGE_B = BOOTSTRAP / "build_stage_b.ps1"
_PHASE5_MANIFEST_MODULES = {
    "modConstants", "modWorkbook", "modAppState", "modTimeline", "modDrivers",
    "modProfiling", "modInflation", "modStructuralCheck", "modCalcContract",
    "modCalcFactors", "modCalcAnalytical", "modCalcFingerprint", "modCalcResolve",
    "modCalcCheck", "modCalcReport",
}
"""The fifteen modules the manifest declared when Phase 5 closed. Frozen by
name so a Phase-6 addition is visible rather than absorbed into a count."""

_PHASE6_MANIFEST_MODULES = {"modSimContract", "modSimRng", "modSimSample",
                            "modSimEngine", "modSimStats", "modSimFingerprint",
                            "modSimNonce", "modSimReport"}
_PHASE7_MANIFEST_MODULES = {"modSimSensitivity", "modSimPostReport", "modSimAnnual",
                            "modSimAnnualRun", "modSimAnnualStore"}

"""Phase-7 hand-written source modules, named on the same terms Phase 6 was:
admitted by name, one at a time, so the earlier half of each inventory
equality below stays exactly as strict as it was."""
"""Every module Phase 6 has added so far, BY NAME. Named rather than counted for
the same reason: a module addition must be a visible edit here, and the
exact-set assertions that consume it keep their full strength."""

_PHASE6_HANDWRITTEN = {"modSimRng", "modSimSample", "modSimEngine", "modSimStats",
                       "modSimFingerprint", "modSimNonce", "modSimReport"}
"""...and which of them are hand-written source rather than generated."""

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

    STEP 12: THE TEMPORARY TREE IS DELETED. This helper is called by more than
    fifty tests and used to leave its `mkdtemp` directory behind on every one of
    them; repeated runs accumulated tens of thousands of `pccm-gateb-*`
    directories and exhausted the writable filesystem. `TemporaryDirectory`
    removes it on the way out, on the exception path as well as the normal one.

    EVERYTHING THE CALLERS NEED IS READ INTO MEMORY FIRST. The returned mapping
    holds parsed JSON and decoded text, so nothing here outlives the directory:
    no caller is handed a path whose validity depends on the tree still
    existing. The former `"dir"` key, which had no consumer at all, is gone -
    returning it would be exactly that kind of dangling handle.
    """
    from pccm_builder import (
        emit_calc_artifacts, emit_inspection, emit_stage_b, load_calc_contract,
        load_contract, load_driver_contract, load_spec, load_structure_contract,
    )

    spec = load_spec(SPEC / "workbook.yaml")
    contract = load_contract(SPEC / "input_contract.yaml")
    drivers = load_driver_contract(SPEC / "driver_contract.yaml")
    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    calc = load_calc_contract(SPEC / "calc_contract.yaml")

    with tempfile.TemporaryDirectory(prefix="pccm-gateb-") as name:
        tmp = Path(name)
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
def test_14_each_api_procedure_has_the_evidence_the_hierarchy_gives_it() -> None:
    """The six API procedures, and what the harness actually establishes for each.

    This test was called `test_14_all_six_api_procedures_are_exercised` and its  # retired-authority
    body searched the scenario source for each name as a string literal. A name
    appearing in source is not an exercise, and the review of ae52bdd is exactly
    about not letting a label outrun its evidence - so the name went with the
    claim. What replaces it is the hierarchy now in force:

        DECLARED    all six, in the persisted project        (P5-M)
        CALLABLE    the five read-only ones, via Excel.Run   (P5-M)
        EXECUTED    PCCM_Calculate, on a valid fixture       (P5-FIX, then P5-AN)
    """
    source = _executable(SCENARIOS)
    emitted = _emitted()["manifest"]["vba"]["api_procedures"]
    assert sorted(emitted) == sorted(PHASE5_API)

    # api_procedures is consumed AS api_procedures, not folded into entry_points.
    assert "$Manifest.vba.api_procedures" in source, (
        "the manifest's api_procedures projection is never read"
    )
    assert "'no API procedure is also an entry point'" in source
    assert "'no API procedure is bound to a button'" in source

    # DECLARED: all six, driven from the manifest list, checked against the
    # persisted project rather than against the manifest that named them.
    p5m = source[source.index("Add-Phase5Result 'P5-FX'"):
                 source.index("Add-Phase5Result 'P5-M'")]
    assert "foreach ($name in $api) {" in p5m
    assert "'the API procedure ' + $name + ' is declared in the persisted project'" in p5m
    assert "($declared -contains $name)" in p5m
    assert "Get-Phase5ProjectProcedureNames -Workbook $Workbook" in p5m

    # CALLABLE: the five, and only through a real Application.Run.
    assert "$probe = $Excel.Run($name)" in p5m
    callable_label = "('the API procedure ' + $name + ' is callable')"
    assert p5m.count(callable_label) == 1
    assert p5m.index("$probe = $Excel.Run($name)") < p5m.index(callable_label)
    # PCCM_Calculate is skipped before that label can be reached. The branch is
    # bounded by its OWN closing brace, not by the label: everything between the
    # two belongs to the five, and slicing that far would read their code as the
    # branch's.
    branch_at = p5m.index("if ($name -eq 'PCCM_Calculate') {")
    branch = p5m[branch_at:p5m.index("\n            }", branch_at)]
    assert "continue" in branch, "the calculation endpoint is not skipped"
    assert "$callable = $true" not in branch, (
        "the calculation endpoint is marked callable inside the branch that skips it"
    )
    assert "Add-Check" not in branch, "the skipping branch emits a check of its own"
    assert branch_at < p5m.index(callable_label), (
        "the skip does not precede the callability label it must avoid"
    )

    # EXECUTED: PCCM_Calculate, first in P5-FIX, then across the corpus in P5-AN.
    driver = source[source.index("function Invoke-Phase5GateBScenarios"):]
    runs = [m.start() for m in re.finditer(r"\$Excel\.Run\('PCCM_Calculate'\)", driver)]
    assert len(runs) >= 2, f"PCCM_Calculate is executed {len(runs)} time(s)"
    assert runs[0] < driver.index("Add-Phase5Result 'P5-FIX'") < runs[1], (
        "the first PCCM_Calculate is not P5-FIX's"
    )
    assert runs[1] < driver.index("Add-Phase5Result 'P5-AN'")
    # And the other five are never executed for their own sake anywhere else -
    # they are read-only probes, which is why callability is safe for them.
    assert "$Excel.Run('PCCM_Calculate')" not in p5m


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
    # PHASE 5 CLOSED AT FIFTEEN AND ALL FIFTEEN ARE STILL HERE. The absolute
    # count stopped being the assertion when Phase-6 Step 6 added
    # modSimContract and modSimRng to the registry; the Phase-5 set is
    # asserted intact and everything beyond it is named.
    assert _PHASE5_MANIFEST_MODULES <= set(declared), (
        sorted(_PHASE5_MANIFEST_MODULES - set(declared))
    )
    assert (set(declared) - _PHASE5_MANIFEST_MODULES
            == _PHASE6_MANIFEST_MODULES | _PHASE7_MANIFEST_MODULES)
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


def test_22_the_diagnostic_import_happens_only_after_the_compile_proof() -> None:
    """No test module may exist before the production project is proved to compile.

    A1 is the first `Application.Run` boundary and nothing more. It used to be
    described here as the first VBA COMPILATION boundary; Runtime Run 7 passed
    A1, passed P5-M as it then stood, and then met a VBE compile error
    in the analytical path, so that description was retired. P5-CMP owns the
    whole-project compile claim, and the diagnostic import must follow IT.
    """
    harness = _executable(HARNESS)
    a1 = harness.index("Add-Result 'A1'")
    invoke = harness.index("Invoke-Phase5GateBScenarios")
    assert a1 < invoke, "the Phase-5 scenarios run before the automation surface is proved"
    # A1 is still the FIRST Application.Run in the harness - that claim survives.
    first_run = harness.index("$excel.Run(")
    assert first_run < a1, "A1 does not contain the first Application.Run"
    assert "PCCM_AutomationBegin" in harness[first_run:first_run + 80], (
        "the first Application.Run is not a production procedure"
    )
    # And the import is inside the scenario file, after BOTH gates.
    scenarios = _executable(SCENARIOS)
    assert "$components.Import($source)" in scenarios
    imported = scenarios.index("$components.Import($source)")
    prerequisite = scenarios.index("Add-Phase5Result 'P5-P4'")
    compiled = scenarios.index("Add-Phase5Result 'P5-CMP'")
    assert prerequisite < imported, (
        "the diagnostic module is imported before the Phase-4 prerequisite is checked"
    )
    assert compiled < imported, (
        "the diagnostic module is imported before the whole-project compile gate, so "
        "a test module could mask the compile proof it is supposed to follow"
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
    assert _PHASE5_MANIFEST_MODULES <= set(modules)
    assert (set(modules) - _PHASE5_MANIFEST_MODULES
            == _PHASE6_MANIFEST_MODULES | _PHASE7_MANIFEST_MODULES)
    on_disk = {path.stem for path in SRC_VBA.glob("*.bas")}
    assert DIAGNOSTIC_MODULE_NAME not in on_disk
    # The thirteen Phase-5 hand-written modules, plus Phase 6's source modules
    # and Phase 7's. Named rather than counted, so an addition is visible - and
    # naming a later phase's module never relaxes the Phase-5 half.
    assert on_disk == ((_PHASE5_MANIFEST_MODULES
                        - {"modConstants", "modCalcContract"})
                       | _PHASE6_HANDWRITTEN | _PHASE7_MANIFEST_MODULES), (
        f"a production module was added or removed: {sorted(on_disk)}"
    )
    # The harness asserts all three of those things at runtime too.
    source = _executable(SCENARIOS)
    assert "'NO shape has OnAction = PCCM_Calculate'" in source
    assert "'exactly the five declared (sheet, shape, macro) bindings exist'" in source
    assert "'every macro-bound shape is one of the five declared buttons'" in source
    assert "'no undeclared shape invokes a PCCM_ procedure'" in source
    assert "'the button ' + $wantSheet + '!' + $wantName + ' calls ' + $wantAction" in source
    assert "'the manifest names a well-formed production module set'" in source
    assert "': the production module ' + $name + ' is a standard module'" in source


# A FIXED PRODUCTION-MODULE COUNT IS A SECOND AUTHORITY.
#
# P5-M used to gate on `$expected.Count -eq 15` beside the exact-set check. The
# helper was already manifest-derived, so the literal proved nothing the set did
# not - until Phase 6 legitimately added modSimContract and modSimRng, at which
# point a CORRECT workbook failed on it.
#
# The defect is the SECOND AUTHORITY, not the number 15. `-eq 17` would be the
# same defect one module later, so the assertions below reject a production-module
# count literal of ANY value, in either direction, and `test_38b` proves that
# rejection is real by reintroducing `-eq 17` and requiring a failure.
_MODULE_COUNT_LITERAL = re.compile(
    r"\$expected\.Count\s+-(?:eq|ge|gt|le|lt|ne)\s+\d+"
    r"|manifest declares \d+ production modules"
    r"|\d+\s+production modules"
)


def _active_p5m_block(source: str) -> str:
    """The ACTIVE P5-M block, comment-free, bounded by its neighbours' results."""
    return source[source.index("Add-Phase5Result 'P5-FX'"):
                  source.index("Add-Phase5Result 'P5-M'")]


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
    # The helper takes its expectation from its parameter and never from a
    # number of its own.
    assert not _MODULE_COUNT_LITERAL.search(helper), _MODULE_COUNT_LITERAL.search(helper)

    block = _active_p5m_block(source)
    # A. P5-M derives the expectation from the manifest.
    assert ("$expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })"
            in block), "P5-M does not read the module set from the manifest"
    # B. and reaches the production namespace through the shared helper, so P5-M
    #    and P5-D8 cannot drift apart.
    assert "Add-Phase5ModuleInventoryChecks -List $list -Components $components" in block
    assert "-ExpectedModules $expected" in block
    # D. and states no production-module count of its own, at any value.
    found = _MODULE_COUNT_LITERAL.search(block)
    assert not found, f"P5-M reintroduced a production-module count literal: {found.group(0)!r}"


def test_38a_p5_d8_returns_the_inventory_to_the_manifest_set_not_a_count() -> None:
    source = _executable(SCENARIOS)
    block = source[source.index("Add-Phase5Result 'P5-M'"):
                   source.index("Add-Phase5Result 'P5-D8'")]
    assert "Add-Phase5ModuleInventoryChecks -List $list -Components $inventory" in block
    found = _MODULE_COUNT_LITERAL.search(block)
    assert not found, f"P5-D8 carries a production-module count literal: {found.group(0)!r}"
    # The CURRENT result titles say what is proved without saying how many.
    raw = _text(SCENARIOS)
    assert ("'Persisted project: manifest module set by name, 5 buttons, "
            "6 API procedures'") in raw
    assert ("'Transient diagnostic module removed; inventory back to the "
            "manifest module set'") in raw


def test_38b_reintroducing_a_hard_coded_module_count_is_refused() -> None:
    """The mutation control: `-eq 17` is the same defect as `-eq 15`."""
    source = _executable(SCENARIOS)
    anchor = "$expected = @($Manifest.vba.modules | ForEach-Object { [string]$_.name })"
    assert source.count(anchor) >= 1
    for reintroduced in (
        "$null = Add-Check $list 'the manifest declares 17 production modules' `\n"
        "            ($expected.Count -eq 17) (\"declared \" + $expected.Count)",
        "$null = Add-Check $list 'inventory' ($expected.Count -ge 17) ''",
        "$null = Add-Check $list 'inventory' ($expected.Count -eq 15) ''",
    ):
        damaged = source.replace(anchor, anchor + "\n        " + reintroduced, 1)
        assert damaged != source
        found = _MODULE_COUNT_LITERAL.search(_active_p5m_block(damaged))
        assert found, f"the detector missed a reintroduced count: {reintroduced!r}"
    # And the accepted source is clean, so the control is not measuring noise.
    assert not _MODULE_COUNT_LITERAL.search(_active_p5m_block(source))


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
    assert len(names) == len(emitted["manifest"]["vba"]["modules"]) + 1, (
        "the inventory growth must be visible"
    )


def test_nc_17_the_diagnostic_imported_before_a1_is_caught() -> None:
    planted = _synthetic(
        "$components.Import($source)\n"
        "$excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null\n"
        "Add-Result 'A1' 'VBA automation surface callable' 'PASS' ''\n"
    )
    assert planted.index("$components.Import($source)") < planted.index("Add-Result 'A1'"), (
        "an import that precedes the first production Application.Run must be visible"
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
def test_46_the_harness_source_states_what_has_and_has_not_been_run() -> None:
    """The distinction this control exists for is source evidence versus runtime
    evidence. What changed is which side of it Phase 5 sits on.

    It used to require the driver banner to say `NO PHASE-5 GATE-B RUN HAS BEEN
    MADE`, which was true when Phase 5 was submitted and false once its
    scenarios had executed on real Excel and Phase 5 had been accepted on that
    evidence. A control that demands a statement the project has outgrown does
    not protect the distinction — it forces the source to misdescribe itself.

    So the demand moves rather than disappearing: the banner must record that
    Phase 5 Gate B is closed, and must still separate what has executed from
    what has not, because Step 13 is the part that is now unproven.
    """
    for path in (SCENARIOS, DIAGNOSTIC):
        assert "NOT" in _text(path) or "not been" in _text(path)

    banner = _text(HARNESS).split("#>")[0]
    assert "NO PHASE-5 GATE-B RUN HAS BEEN MADE" not in banner, (
        "the driver banner denies a Phase-5 Gate-B run that has happened"
    )
    assert "PHASE 5 GATE B IS CLOSED" in banner, (
        "the driver banner does not record that Phase 5 Gate B has executed"
    )
    # AND THE LINE IS STILL DRAWN. A banner that only reported closure would be
    # the same over-claim in the other direction - but WHERE the line falls
    # moved when Step 13 closed. While Step 13 was open it was the Phase-6 block
    # still under runtime validation; now that it is closed, what remains
    # unproven is the bounded set of clauses no run could induce. Naming only
    # the first is what made this control demand a state the harness had left.
    flat = " ".join(banner.split())
    assert re.search(r"still under runtime validation|not yet been exercised"
                     r"|static-only|not induced", flat), (
        "the banner records what has run without saying what has not"
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


def _fixture_steps(source: str) -> str:
    """The fixture CHOREOGRAPHY, which Run 10 moved behind an output boundary.

    `Set-Phase5Fixture` is now a thin boundary that calls the steps, counts what
    they emitted and refuses to return anything it cannot account for. Steps A-H
    live in `Invoke-Phase5FixtureSteps`, so every choreography assertion reads
    that; the boundary has assertions of its own.
    """
    return _procedure(source, "Invoke-Phase5FixtureSteps")


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
    # The CLAIM is unchanged by the Run-6 correction; only the expression the
    # inflation writer branches on moved. It used to be the dynamic lookup
    # `$rates.$year`; it is now the Value of the same PSPropertyInfo the year
    # name came from, which is strictly tighter - the blank branch can no longer
    # be reached through a lookup that missed.
    source = _executable(SCENARIOS)
    for procedure, subject in (("Write-Phase5InflationRates", "$rateValue"),
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
    # And the harness really honours the flag. Run 10 routed the read through
    # the required-property accessor - the direct read is a StrictMode hazard
    # for optional properties and the schema table is now the single authority
    # on which is which - so the assertion follows the read, not its spelling.
    assert "-Name 'apply_timeline'" in body, (
        "the mutation applier never consults apply_timeline at all"
    )
    guard = "if (Get-Phase5RequiredProperty -Object $Mutation -Name 'apply_timeline'"
    assert body.count(guard) == 2, (
        "the applier does not BRANCH on apply_timeline in both the "
        f"entered_structure and config_profile_add kinds: {body.count(guard)}"
    )
    for always in ("if ($true)", "if (1)"):
        assert always not in body, f"the apply is unconditional ({always})"


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
    # RUN 11 replaced the wildcard comparison with a literal substring search:
    # -like read `fraction in [0, 1]` as a character class and rejected a detail
    # that visibly contained it. The rule asserted here is the CONTRACT - the
    # token occurs as literal text - not the spelling of the expression.
    assert "IndexOf(" in helper and "OrdinalIgnoreCase" in helper, (
        "the discriminator is not a literal substring search"
    )
    assert "-like" not in helper, "the wildcard matcher is back"  # refusal-list
    # AND AN EMPTY TOKEN IS NOT A DISCRIMINATOR. Any substring search finds ""
    # at index 0, so without a guard an emitted blank would pass every detail
    # and this test's whole subject - specificity - would be unenforced.
    assert "IsNullOrWhiteSpace($literal)" in helper, (
        "an empty discriminator token would pass vacuously"
    )
    assert helper.index("IndexOf(") > helper.index("IsNullOrWhiteSpace($literal)"), (
        "the token reaches IndexOf before it is known to be non-empty"
    )
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
    fixture = _fixture_steps(source)
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
    """BLOCKER 2, carried forward and tightened by Runtime Run 5.

    The claim is unchanged: a broken fixture must fail at fixture establishment,
    loudly, and never surface as a scenario predicate. What Run 5 changed is
    WHERE the two gates sit. They used to sit after the driver Adds, which is
    what let a failed Add write an orphan row and then be reported by the Apply
    that refused to mutate over it. They now sit before the first Add, so the
    baseline every Add depends on is proved before any Add is attempted.
    """
    source = _executable(SCENARIOS)
    fixture = _fixture_steps(source)

    # THE APPLY IS A CHECKED PRODUCTION MUTATION, not a piped-away call.
    apply_at = fixture.index("-Operation 'PCCM_ApplyTimeline'")
    assert "$Excel.Run('PCCM_ApplyTimeline') | Out-Null" not in fixture, (
        "the fixture still discards the Apply result"
    )
    checked = _procedure(source, "Invoke-Phase5ProductionOperation")
    assert "$result -notlike 'OK|*'" in checked, (
        "the checked-operation helper does not require a success result"
    )
    assert "throw (" in checked, "a failed production mutation does not stop the fixture"

    # AND THE STRUCTURE IS PRODUCTION'S OWN JUDGEMENT, which also throws.
    coherent = _procedure(source, "Assert-Phase5StructurallyCoherent")
    assert "PCCM_StructuralReport" in coherent, (
        "the coherence gate never asks production whether the structure is coherent"
    )
    assert "throw (" in coherent, "an incoherent structure does not stop the fixture"

    # BOTH GATES PRECEDE THE FIRST DRIVER ADD - the Run-5 correction itself.
    report_at = fixture.index("Assert-Phase5StructurallyCoherent")
    add_at = fixture.index("Invoke-Phase5AddDriverAndRequireSuccess")
    assert apply_at < report_at < add_at, (
        "a fixture driver is added before the structural baseline is proved"
    )

    # AND THEY STILL PRECEDE the value writers, exactly as before.
    rates_at = fixture.index("Write-Phase5InflationRates")
    weights_at = fixture.index("Write-Phase5Weights")
    assert report_at < rates_at, "inflation rates are written before the structural gate"
    assert report_at < weights_at, "profiling weights are written before the structural gate"

    # It THROWS. It does not return a diagnostic the caller may ignore.
    assert "throw (" in _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess")
    assert "throw (" in _procedure(source, "Clear-Phase5Registers")


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
    # `$name` is already [string], taken off the PSPropertyInfo since Run 6.
    assert "$rowIndex = Find-GridRow -Workbook $Workbook -Grid $grid -Key $name" in body, (
        "the profile row is not located by name"
    )
    assert "$name = [string]$profileProperty.Name" in body, (
        "the profile name no longer comes from an individual property object"
    )
    assert "$rowIndex++" not in body, "the profile row is still an incremented counter"
    # The column axis stays keyed by calendar-year header.
    assert "[array]::IndexOf($headers, $year)" in body
    assert "$year = [string]$rateProperty.Name" in body, (
        "the calendar year no longer comes from an individual property object"
    )
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
    fixture = _fixture_steps(source)
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
    # Review round 4A: one exact CLR type-identity gate replaces the three
    # separate probes, and also rejects Int32 1 against Double 1.
    assert "if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }" in strict, (
        "a String that looks numeric, or a boxed Int32, is accepted against a number"
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
    # Bounded at both ends: after the procedure's own param block, and before
    # `$vbe = $Excel.VBE`, the first statement that touches anything outside the
    # result ledger. That statement opens P5-CMP, the Run-7 compile gate, which
    # now sits between P5-P4 and the locked FX capture.
    return body[body.index("$required = Get-Phase4RequiredScenarioIds"):
                body.index("$vbe = $Excel.VBE")]


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
    assert "if ($code -match [regex]::Escape([string]$rule.construct))" in ev, (
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
    # The requirement itself is untouched, but its SOURCE moved: once a scoped
    # rule exists the flattened list cannot express `allowed_in`, so enforcement
    # reads the structured rules and the flattened field is display only.
    assert "Get-ForbiddenConstructRules -Manifest $Manifest" in ev
    assert "Test-ConstructForbiddenIn -Rule $rule" in ev, (
        "the scan is not module-aware, so a scoped construct would be read as global"
    )
    assert "@($Manifest.vba.forbidden_constructs)" not in ev, (
        "the flattened list is still the enforcement authority"
    )
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
    assert "if ($code -match [regex]::Escape([string]$rule.construct))" in source, (
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
    # Review round 4A: the three separate type probes became ONE exact CLR
    # type-identity gate, which subsumes them and also catches Int32 1 vs
    # Double 1 - a pair the old `[double] -eq [double]` tail compared EQUAL.
    assert "if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }" in comparator, (
        "the comparator no longer establishes exact CLR type identity first"
    )
    assert "[double]$Actual -eq [double]$Expected" in comparator
    assert "[string]$Actual -ceq [string]$Expected" in comparator


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
    planted = _synthetic("if ($raw -match [regex]::Escape([string]$rule.construct)) {\n")
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


def _scan_production_against_rules(strip=None) -> list:
    """The corrected scan, MODULE-AWARE, over every frozen production module.

    D6-11 is per module: a scoped rule names the one module allowed to contain
    the construct in executable code, and the sweep must respect that or it
    would flag the module that owns it. The structured rules are the authority;
    the flattened list cannot express a scope.
    """
    sys.path.insert(0, str(PCCM_ROOT / "builder"))
    from pccm_builder.vba_source import strip_comments, strip_strings

    if strip is None:
        strip = lambda text: strip_strings(strip_comments(text))  # noqa: E731

    rules = _emitted()["manifest"]["vba"]["forbidden_construct_rules"]
    offenders = []
    for path in sorted(SRC_VBA.glob("*.bas")):
        body = strip(_text(path))
        for rule in rules:
            if path.stem in rule["allowed_in"]:
                continue
            if rule["construct"] in body:
                offenders.append(f"{path.name}: {rule['construct']}")
    return offenders


def test_129_the_production_source_still_passes_the_corrected_scan() -> None:
    """The accepted production VBA must be clean under the new semantics.

    A stricter scanner that flagged real production code would be a different
    defect, so this runs the corrected rule over every frozen module.
    """
    offenders = _scan_production_against_rules()
    assert not offenders, (
        "the corrected scan flags accepted production source: " + "; ".join(offenders)
    )
    sys.path.insert(0, str(PCCM_ROOT / "builder"))
    from pccm_builder.vba_source import strip_comments, strip_strings

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
    rules = _emitted()["manifest"]["vba"]["forbidden_construct_rules"]
    offenders = []
    for path in sorted(SRC_VBA.glob("*.bas")) + [DIAGNOSTIC]:
        body = _model_executable(_text(path))
        for rule in rules:
            if path.stem in rule["allowed_in"]:
                continue
            if rule["construct"] in body:
                offenders.append(f"{path.name}: {rule['construct']}")
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
    for probe in ("$Value -is [string]", "$Value -is [bool]",
                  "$Value.GetType().FullName -ceq 'System.Double'"):
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
    # Review round 4A: the three separate type probes became ONE exact CLR
    # type-identity gate, which subsumes them and also catches Int32 1 vs
    # Double 1 - a pair the old `[double] -eq [double]` tail compared EQUAL.
    assert "if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }" in comparator, (
        "the comparator no longer establishes exact CLR type identity first"
    )
    assert "[double]$Actual -eq [double]$Expected" in comparator
    assert "[string]$Actual -ceq [string]$Expected" in comparator
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
    stray = [identifier for identifier
             in re.findall(r"(?<!Phase5)Add-Result\s+'(P5-[A-Z0-9]+)'", both)
             # P5-LDG is the ledger's own report. It is emitted through
             # Add-Result deliberately, so the ledger can never suppress the
             # result that reports on the ledger, and it carries its own
             # emitted-once flag instead.
             if identifier != "P5-LDG"]
    assert not stray, f"these Phase-5 IDs bypass the one-result guard: {sorted(set(stray))}"
    integrity = _procedure(source, "Add-Phase5LedgerIntegrityResult")
    assert "if ($script:Phase5LedgerReported) { return }" in integrity, (
        "P5-LDG has no emitted-once flag of its own"
    )
    # A duplicate attempt is a VIOLATION, not a note, and P5-LDG is actually
    # called - a guard that records nothing, or a report nobody emits, is the
    # fail-open shape this round replaced.
    assert "$script:Phase5LedgerViolations.Add(" in guard, (
        "a duplicate attempt is not recorded as a violation"
    )
    assert "Add-Phase5LedgerIntegrityResult" in _executable(HARNESS), (
        "the ledger integrity result is never emitted"
    )
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


# ===========================================================================
# 28. REVIEW ROUND 4A: exact captured types, and a ledger that fails closed
# ===========================================================================
# The Run-4 correction accepted Single, Int16, Int32, Int64, Byte and Decimal
# and wrote all of them as Double - normalisation, not restoration - and the
# comparator ended in `[double]$Actual -eq [double]$Expected`, so a captured
# Int32 1 compared EQUAL to a restored Double 1. The setter now fails closed and
# the comparator establishes exact CLR type identity before any value is read.
SUPPORTED_CAPTURED_TYPES = ("System.String", "System.Double", "System.Boolean")
REFUSED_CAPTURED_TYPES = ("System.Single", "System.Int16", "System.Int32",
                          "System.Int64", "System.Byte", "System.Decimal",
                          "System.DateTime")


def test_148_the_setter_supports_only_the_approved_captured_types() -> None:
    setter = _typed_setter()
    # The whole supported set, each on its own COM assignment line.
    assert "if ($null -eq $Value) {" in setter
    assert "$null = $cell.ClearContents()" in setter
    assert "$cell.Value2 = [string]$Value" in setter
    assert "$cell.Value2 = [bool]$Value" in setter
    assert "$cell.Value2 = [double]$Value" in setter
    assert setter.count("$cell.Value2") == 3, (
        "there is not exactly one COM assignment per supported type"
    )
    # Double is matched by EXACT type name.
    #
    # REVIEW ROUND 4A CORRECTION. An earlier version of this comment claimed
    # `-is [double]` "would be true for a boxed Int32 under PowerShell's numeric
    # conversions". That is FALSE and the rationale is withdrawn. `-is` is a
    # .NET instance test: it asks whether the object IS of that type and does no
    # numeric conversion, so `1 -is [double]` is $false.
    #
    # The real defect in the old setter was different: it carried explicit
    # `-is [int]`, `-is [single]`, `-is [long]` (and more) branches and wrote
    # every one of them through `[double]$Value`. The NORMALISATION was in those
    # branches and that cast, not in the `-is [double]` test. Exact-type
    # matching is still the right implementation - it just fixes a widened
    # branch set, not a lying operator.
    assert "$Value.GetType().FullName -ceq 'System.Double'" in setter, (
        "the numeric branch matches by convertibility rather than by exact type"
    )
    # Every widened numeric alias, and the DateTime branch, are gone.
    for retired in ("$Value -is [single]", "$Value -is [int]", "$Value -is [long]",
                    "$Value -is [decimal]", "$Value -is [int16]", "$Value -is [byte]",
                    "$Value -is [datetime]", "$cell.Value2 = [datetime]$Value"):
        assert retired not in setter, f"{retired} is back; it normalises the capture"
    # Anything else throws BEFORE any assignment.
    assert "throw (" in setter and "$Value.GetType().FullName" in setter
    tail = setter[setter.index("throw ("):]
    assert "$cell.Value2" not in tail and "ClearContents" not in tail, (
        "an assignment follows the refusal branch"
    )


def test_149_unsupported_numeric_types_are_refused_not_normalised() -> None:
    """The decision table the setter has to implement."""
    def branch(type_name: str | None) -> str:
        if type_name is None:
            return "clear"
        if type_name == "System.String":
            return "string"
        if type_name == "System.Boolean":
            return "bool"
        if type_name == "System.Double":
            return "double"
        return "throw"

    assert branch(None) == "clear"
    for supported, expected in (("System.String", "string"),
                                ("System.Boolean", "bool"),
                                ("System.Double", "double")):
        assert branch(supported) == expected
    for refused in REFUSED_CAPTURED_TYPES:
        assert branch(refused) == "throw", (
            f"{refused} is converted instead of refused, which normalises the capture"
        )
    # The refusal happens in the SETTER, before the write - not discovered later
    # by the read-back.
    setter = _typed_setter()
    assert setter.index("throw (") > setter.index("$cell = $body.Cells(")
    assert "Test-Phase5ExactValue" not in setter, (
        "the setter defers its own type decision to the read-back comparator"
    )


def test_150_the_comparator_requires_exact_clr_type_identity() -> None:
    comparator = _procedure(_executable(SCENARIOS), "Test-Phase5ExactValue")
    assert "if ($null -eq $Expected) { return ($null -eq $Actual) }" in comparator
    assert "if ($null -eq $Actual) { return $false }" in comparator
    gate = "if ($Actual.GetType().FullName -cne $Expected.GetType().FullName) { return $false }"
    assert gate in comparator, "there is no exact type-identity gate"
    # The gate comes BEFORE any value comparison.
    assert comparator.index(gate) < comparator.index("[string]$Actual -ceq")
    assert comparator.index(gate) < comparator.index("[double]$Actual -eq")
    assert "-ceq [string]$Expected" in comparator, "text is compared case-insensitively"
    assert "Tolerance" not in comparator, "the strict comparator carries a tolerance"
    for forbidden in ("Format-CalcValue", "NumberFormat", ".Text", "-as [double]"):
        assert forbidden not in comparator, f"{forbidden} reached the exact comparator"


def test_151_the_exact_comparator_decision_table() -> None:
    """Every pair the review named, plus the ones that must still pass."""
    class Boxed:
        """A value with an explicit CLR type name, as the comparator sees it."""
        def __init__(self, type_name: str, value):
            self.type_name = type_name
            self.value = value

    def exact(actual, expected) -> bool:
        if expected is None:
            return actual is None
        if actual is None:
            return False
        if actual.type_name != expected.type_name:
            return False
        if expected.type_name == "System.String":
            return actual.value == expected.value          # case-sensitive
        return actual.value == expected.value

    string_one = Boxed("System.String", "1")
    double_one = Boxed("System.Double", 1.0)
    empty = Boxed("System.String", "")

    # MUST FAIL - every one of these passed the old comparator or the old setter.
    for captured, restored, why in (
        (Boxed("System.Int32", 1), double_one, "Int32 1 -> Double 1"),
        (Boxed("System.Int64", 1), double_one, "Int64 1 -> Double 1"),
        (Boxed("System.Single", 1.0), double_one, "Single 1 -> Double 1"),
        (Boxed("System.Decimal", 1), double_one, "Decimal 1 -> Double 1"),
        (string_one, double_one, 'String "1" -> Double 1'),
        (double_one, string_one, 'Double 1 -> String "1"'),
        (None, empty, "null -> empty String"),
    ):
        assert not exact(restored, captured), f"{why} was accepted as an exact restoration"

    # MUST PASS.
    assert exact(string_one, Boxed("System.String", "1"))
    assert exact(double_one, Boxed("System.Double", 1.0))
    assert exact(Boxed("System.Boolean", True), Boxed("System.Boolean", True))
    assert exact(None, None)
    # And case still matters for text.
    assert not exact(Boxed("System.String", "sar"), Boxed("System.String", "SAR"))


def test_152_p5_fx_proves_exact_type_through_the_strict_comparator() -> None:
    """The early gate is only worth having if it can see a normalisation."""
    source = _executable(SCENARIOS)
    block = source[source.index("Save-Phase5LockedFxSeed -Workbook"):
                   source.index("Add-Phase5Result 'P5-FX'")]
    assert "Test-Phase5ExactValue -Actual $restored[0][0] -Expected $seed.Currency" in block
    assert "Test-Phase5ExactValue -Actual $restored[0][1] -Expected $seed.Rate" in block
    assert "the captured value AND the captured type" in block
    # It goes through the same comparator that now gates on type identity.
    comparator = _procedure(source, "Test-Phase5ExactValue")
    assert "GetType().FullName -cne" in comparator
    # And through the real restoration path, so a refused type throws there.
    assert "Reset-Phase5FxTable -Workbook $Workbook -Inspection $Inspection -Seed $seed" in block


# --- the ledger fails closed ------------------------------------------------
FINAL_COMPLETENESS_CALL = "Add-Phase4FinalCompletenessResult -Results $results"
LEDGER_VERDICT_CALL = "Add-Phase5LedgerIntegrityResult"
FAIL_COUNT_SUMMARY = "$failed = @($results | Where-Object { $_.Status -eq 'FAIL' })"


def _assert_final_ledger_order(harness: str) -> None:
    """REVIEW ROUND 4A. The ledger verdict is the LAST Phase-5 result.

    Add-Phase4FinalCompletenessResult is not a bystander here: it ends by
    emitting P5-FIN through Add-Phase5Result, so it is a GUARDED Phase-5 result.
    Emitting the verdict before it was fail-open - a duplicate P5-FIN attempt
    would be recorded as a violation after P5-LDG had already reported PASS, the
    emitted-once flag would suppress any later verdict, and the FAIL count could
    still be zero.

    `ledger < summary` alone does not say this. The completeness call has to sit
    BETWEEN the scenarios and the verdict, which is what is asserted here.
    """
    for call in (FINAL_COMPLETENESS_CALL, LEDGER_VERDICT_CALL, FAIL_COUNT_SUMMARY):
        assert harness.count(call) == 1, (call, harness.count(call))
    completeness = harness.index(FINAL_COMPLETENESS_CALL)
    verdict = harness.index(LEDGER_VERDICT_CALL)
    summary = harness.index(FAIL_COUNT_SUMMARY)
    assert completeness < verdict, (
        "P5-FIN is emitted AFTER the ledger verdict, so a duplicate P5-FIN "
        "attempt cannot reach P5-LDG and the run can finish green on it"
    )
    assert verdict < summary, (
        "the ledger verdict is emitted after the FAIL count is taken"
    )


def test_153_a_duplicate_attempt_is_recorded_as_a_violation() -> None:
    source = _executable(SCENARIOS)
    guard = _procedure(source, "Add-Phase5Result")
    assert "$script:Phase5LedgerViolations.Add(" in guard, (
        "a duplicate attempt is not recorded anywhere"
    )
    assert "Add-Note (" in guard, "the duplicate attempt is invisible"
    assert "return" in guard, "a second result is appended"
    # It must NOT throw: that would take Y, Z and P5-FIN down with it.
    assert "throw" not in guard, (
        "the guard throws, which would destroy the cleanup and lifecycle evidence"
    )
    assert "$null = $script:Phase5RecordedIds.Add($Id)" in guard
    reset = _procedure(source, "Reset-Phase5ResultLedger")
    for field in ("Phase5RecordedIds", "Phase5LedgerViolations", "Phase5LedgerReported"):
        assert field in reset, f"{field} is not reset at the start of a run"
    # And the report is conditional on the violations, not on nothing.
    integrity = _procedure(source, "Add-Phase5LedgerIntegrityResult")
    assert "$violations = @($script:Phase5LedgerViolations)" in integrity
    assert "if ($violations.Count -eq 0) {" in integrity, (
        "P5-LDG reports the same verdict whether or not a duplicate was attempted"
    )
    assert "if ($true)" not in integrity


def test_154_any_duplicate_attempt_forces_the_run_to_fail() -> None:
    """Notes do not contribute to the FAIL count. P5-LDG does."""
    source = _executable(SCENARIOS)
    integrity = _procedure(source, "Add-Phase5LedgerIntegrityResult")
    assert "if ($script:Phase5LedgerReported) { return }" in integrity, (
        "many duplicate attempts could produce many P5-LDG results"
    )
    assert "$script:Phase5LedgerReported = $true" in integrity
    assert "Add-Result 'P5-LDG'" in integrity
    assert "'FAIL'" in integrity and "'PASS'" in integrity
    assert "($violations.Count -eq 0)" in integrity
    assert "'SKIP'" not in integrity, "an integrity violation must never be a skip"

    # Wired after Y, Z AND P5-FIN, so cleanup evidence survives and every
    # guarded Phase-5 result has been attempted, and before the summary.
    harness = _executable(HARNESS)
    assert "Add-Phase5LedgerIntegrityResult" in harness, "P5-LDG is never emitted"
    assert harness.index("Add-Result 'Y' 'Transient COM releases' 'PASS'") < \
        harness.index("Add-Phase5LedgerIntegrityResult")
    assert harness.index("Add-Result 'Z' 'Excel closed naturally after the functional run' 'PASS'") < \
        harness.index("Add-Phase5LedgerIntegrityResult")
    _assert_final_ledger_order(harness)
    # The run's verdict is the FAIL count, so a P5-LDG FAIL makes it red.
    assert "if ($failed.Count -eq 0) {" in harness and "exit 1" in harness
    # And nothing de-duplicates at print time.
    for forbidden in ("Select-Object -Unique -Property Id", "Group-Object Id",
                      "Sort-Object Id -Unique"):
        assert forbidden not in harness, f"the report de-duplicates ({forbidden})"


def test_155_the_ledger_models_both_paths_and_the_grouped_catch() -> None:
    """Success, duplicate, and the grouped-catch shape, end to end."""
    class Ledger:
        def __init__(self):
            self.results: list[tuple[str, str]] = []
            self.violations: list[str] = []
            self.reported = False

        def add(self, identifier: str, status: str) -> None:
            if any(i == identifier for i, _ in self.results):
                self.violations.append(f"{identifier} (attempted as {status})")
                return
            self.results.append((identifier, status))

        def integrity(self) -> None:
            if self.reported:
                return
            self.reported = True
            self.results.append(
                ("P5-LDG", "PASS" if not self.violations else "FAIL"))

        @property
        def failed(self) -> list[str]:
            return [i for i, s in self.results if s == "FAIL"]

    # SUCCESS PATH: no duplicate, no violation, no integrity FAIL.
    clean = Ledger()
    for identifier in ("P5-FX", "P5-S2", "P5-ST", "P5-S3", "P5-S4"):
        clean.add(identifier, "PASS")
    clean.integrity()
    assert not clean.violations
    assert ("P5-LDG", "PASS") in clean.results
    assert clean.failed == [], clean.failed

    # DUPLICATE PATH: P5-S2 PASS, then attempted FAIL.
    duplicate = Ledger()
    duplicate.add("P5-S2", "PASS")
    duplicate.add("P5-S2", "FAIL")
    duplicate.integrity()
    assert [i for i, _ in duplicate.results].count("P5-S2") == 1, "a second result appeared"
    assert duplicate.results[0] == ("P5-S2", "PASS"), "the first result did not stand"
    assert duplicate.violations, "the duplicate attempt vanished"
    assert duplicate.failed == ["P5-LDG"], duplicate.failed
    assert duplicate.failed, "the run could still finish green on a duplicate attempt"

    # GROUPED CATCH: S3 and S4 recorded, then the catch attempts the whole group.
    grouped = Ledger()
    grouped.add("P5-S3", "PASS")
    grouped.add("P5-S4", "PASS")
    for identifier in ("P5-S3", "P5-S4", "P5-S5", "P5-KP", "P5-RC"):
        grouped.add(identifier, "FAIL")
    grouped.integrity()
    ids = [i for i, _ in grouped.results]
    assert ids.count("P5-S3") == 1 and ids.count("P5-S4") == 1
    assert ("P5-S3", "PASS") in grouped.results and ("P5-S4", "PASS") in grouped.results
    assert ("P5-S5", "FAIL") in grouped.results
    assert len(grouped.violations) == 2, grouped.violations
    assert "P5-LDG" in grouped.failed

    # MANY duplicates still produce exactly ONE integrity result.
    many = Ledger()
    many.add("P5-S2", "PASS")
    for _ in range(5):
        many.add("P5-S2", "FAIL")
    many.integrity()
    many.integrity()
    assert [i for i, _ in many.results].count("P5-LDG") == 1
    assert len(many.violations) == 5

    # ------------------------------------------------------------------
    # REVIEW ROUND 4A: THE FINAL-RESULT BOUNDARY.
    #
    # P5-FIN is the LAST guarded Phase-5 result, and it is emitted by
    # Add-Phase4FinalCompletenessResult. The two cases below are the whole point
    # of moving the verdict after it: the second one is unreachable while the
    # verdict runs first, because the violation would be recorded after P5-LDG
    # had already said PASS.
    # ------------------------------------------------------------------
    # A. CLEAN FINALISATION.
    final_clean = Ledger()
    for identifier in ("P5-FX", "P5-S2", "P5-ST"):
        final_clean.add(identifier, "PASS")
    final_clean.add("P5-FIN", "PASS")
    final_clean.integrity()
    assert [i for i, _ in final_clean.results].count("P5-FIN") == 1
    assert ("P5-FIN", "PASS") in final_clean.results
    assert [i for i, _ in final_clean.results].count("P5-LDG") == 1
    assert ("P5-LDG", "PASS") in final_clean.results
    assert final_clean.violations == []
    assert final_clean.failed == [], final_clean.failed
    # The verdict is genuinely last.
    assert [i for i, _ in final_clean.results][-1] == "P5-LDG"

    # B. A DUPLICATE P5-FIN ATTEMPT, BEFORE THE VERDICT.
    final_dup = Ledger()
    for identifier in ("P5-FX", "P5-S2", "P5-ST"):
        final_dup.add(identifier, "PASS")
    final_dup.add("P5-FIN", "PASS")
    final_dup.add("P5-FIN", "FAIL")     # a future ownership defect
    final_dup.integrity()
    ids = [i for i, _ in final_dup.results]
    assert ids.count("P5-FIN") == 1, "a second P5-FIN result appeared"
    assert ("P5-FIN", "PASS") in final_dup.results, "the first P5-FIN did not stand"
    assert ("P5-FIN", "FAIL") not in final_dup.results
    assert len(final_dup.violations) == 1, final_dup.violations
    assert ids.count("P5-LDG") == 1
    assert ("P5-LDG", "FAIL") in final_dup.results
    assert "P5-LDG" in final_dup.failed, final_dup.failed
    assert final_dup.failed, "the run finished green on a duplicate P5-FIN attempt"

    # AND THE ORDERING IS WHAT MAKES B WORK. Run the same events with the
    # verdict taken BEFORE P5-FIN - commit a291853's shape - and the run is
    # green despite the violation.
    early = Ledger()
    for identifier in ("P5-FX", "P5-S2", "P5-ST"):
        early.add(identifier, "PASS")
    early.integrity()                   # the verdict, too soon
    early.add("P5-FIN", "PASS")
    early.add("P5-FIN", "FAIL")
    early.integrity()                   # the emitted-once flag swallows this
    assert len(early.violations) == 1, "the violation was still recorded"
    assert [i for i, _ in early.results].count("P5-LDG") == 1
    assert ("P5-LDG", "PASS") in early.results, (
        "the early verdict should have reported PASS - that is the defect"
    )
    assert early.failed == [], (
        "the early-verdict ordering is only fail-open if the run finishes green"
    )


def test_155a_the_old_early_ledger_ordering_is_refused_by_the_detector() -> None:
    """MUTATION CONTROL. Recreate commit a291853's driver order and require FAIL.

    The mutation moves Add-Phase5LedgerIntegrityResult to immediately before
    Add-Phase4FinalCompletenessResult, which is exactly the shape review round
    4A rejected. It is applied to the REAL harness text, and the REAL detector
    the conformance test uses is run against it.
    """
    harness = _executable(HARNESS)
    # The accepted order passes the detector.
    _assert_final_ledger_order(harness)

    damaged = harness.replace("\n" + LEDGER_VERDICT_CALL, "", 1)
    assert damaged != harness, "the mutation changed nothing"
    assert damaged.count(LEDGER_VERDICT_CALL) == 0, (
        "the verdict call was not removed cleanly"
    )
    damaged = damaged.replace(
        FINAL_COMPLETENESS_CALL,
        LEDGER_VERDICT_CALL + "\n" + FINAL_COMPLETENESS_CALL, 1)
    assert damaged.count(LEDGER_VERDICT_CALL) == 1
    assert damaged.index(LEDGER_VERDICT_CALL) < damaged.index(FINAL_COMPLETENESS_CALL), (
        "the mutation did not actually recreate the early-ledger ordering"
    )

    # THE DETECTOR MUST REFUSE IT.
    try:
        _assert_final_ledger_order(damaged)
    except AssertionError:
        pass
    else:
        raise AssertionError(
            "the a291853 ordering survived the final-ledger-order detector"
        )

    # AND THE CONTROL IS AIMED AT THE RIGHT INVARIANT. The weaker assertion this
    # round replaced - ledger before the summary - still holds on the damaged
    # text, so it could never have caught this.
    assert damaged.index(LEDGER_VERDICT_CALL) < damaged.index(FAIL_COUNT_SUMMARY), (
        "the mutation broke ledger-before-summary too, so the control would not "
        "prove that the P5-FIN boundary is what is being protected"
    )
    assert damaged.index("Add-Result 'Y' 'Transient COM releases' 'PASS'") < \
        damaged.index(LEDGER_VERDICT_CALL), (
        "the mutation broke the Y ordering too, for the same reason"
    )


# --- negative controls -----------------------------------------------------
def test_nc_98_a_normalising_numeric_branch_is_caught() -> None:
    """The old widened design, restored exactly.

    WHY THIS MUTATION NORMALISES. Not because `Int32 -is [double]` is true - it
    is not; `-is` is a .NET instance test that performs no numeric conversion.
    It normalises because the branch explicitly ALSO accepts `-is [int]`, and
    then writes whatever it accepted through `[double]$Value`. The widening is
    in the accepted set and the cast, which is precisely what the exact-type
    gate removes.
    """
    planted = _synthetic(
        "        } elseif (($Value -is [double]) -or ($Value -is [int])) {\n"
        "            $cell.Value2 = [double]$Value\n"
    )
    assert "-is [int]" in planted, (
        "the widened branch must be visible: it is the `-is [int]` arm, not the "
        "`-is [double]` arm, that lets an Int32 through"
    )
    assert "$cell.Value2 = [double]$Value" in planted, (
        "the cast that performs the normalisation must be visible too"
    )
    assert "GetType().FullName -ceq 'System.Double'" not in planted
    # And the old comparator could not see the result.
    assert float(1) == float(1.0), (
        "Int32 1 and Double 1 compare equal once both are cast to Double, which "
        "is why the type gate has to come first"
    )


def test_nc_99_a_note_only_duplicate_guard_is_caught() -> None:
    planted = _synthetic(
        "    if (Test-Phase5ResultRecorded -Id $Id) {\n"
        "        Add-Note ('suppressed')\n"
        "        return\n"
        "    }\n"
    )
    assert "Phase5LedgerViolations" not in planted, "the fail-open guard must be visible"
    # The arithmetic of failing open: a PASS, a suppressed FAIL, and a green run.
    results = [("P5-X", "PASS")]
    notes = ["P5-X attempted as FAIL"]
    assert [i for i, s in results if s == "FAIL"] == []
    assert notes, "the only trace of the failure is a note"
    assert len([i for i, _ in results]) == len({i for i, _ in results})


# ===========================================================================
# 23. RUNTIME RUN 5: the fixture-establishment ordering defect
# ===========================================================================
# Run 5 is VALID EVIDENCE. It ran on real Windows against a291853, it closed R5
# (P5-FX PASS), and it found a new root: Set-Phase5Fixture performed its
# production mutations in an order where the driver Adds could not succeed, then
# ignored their results and wrote fixture data into the rows they had failed to
# key. The refusal that followed was correct production behaviour reporting a
# workbook the harness had corrupted.
#
# The fifteen regressions below are the enumerated contract for that correction.
def test_156_r1_the_timeline_is_applied_before_the_first_driver_add() -> None:
    """R1. THE RUN-5 ROOT, stated as an ordering invariant.

    Every driver Add runs modStructuralCheck.ValidateStructure on its way out.
    An Add attempted while the Config profile master and Inflation!tblInflation
    disagree therefore CANNOT succeed - so the Apply that reconciles them must
    come first, with no driver added yet.
    """
    source = _executable(SCENARIOS)
    fixture = _fixture_steps(source)
    apply_at = fixture.index("-Operation 'PCCM_ApplyTimeline'")
    add_at = fixture.index("Invoke-Phase5AddDriverAndRequireSuccess")
    assert apply_at < add_at, (
        "PCCM_ApplyTimeline is still invoked after the fixture drivers are added"
    )
    # And there is exactly ONE Apply in the fixture: a second one after the Adds
    # would restore the old shape while leaving this assertion true.
    assert fixture.count("PCCM_ApplyTimeline'") == 1, (
        "the fixture applies the timeline more than once"
    )
    # The step order is declared where the procedure is, not only implied.
    for step in ("--- A.", "--- B.", "--- C.", "--- D.", "--- E.", "--- F.",
                 "--- G.", "--- H."):
        assert step in _text(SCENARIOS), f"fixture step {step} is not marked in the source"


def test_157_r2_no_production_mutation_runs_inside_the_incoherent_window() -> None:
    """R2. Step D deliberately creates the disagreement; step E closes it.

    Between Set-Phase5InflationProfileMaster and the Apply, the workbook is
    KNOWINGLY incoherent - the master has the fixture's profiles and the grid
    still has the previous fixture's. No production endpoint may be called in
    that window, because production is required to refuse there.
    """
    source = _executable(SCENARIOS)
    fixture = _fixture_steps(source)
    master_at = fixture.index("Set-Phase5InflationProfileMaster")
    apply_at = fixture.index("$applied = Invoke-Phase5ProductionOperation")
    assert master_at < apply_at, (
        "the Config profile master is written after the Apply that reads it"
    )
    assert fixture.index("-Operation 'PCCM_ApplyTimeline'") > apply_at, (
        "the Apply statement was not located; the window below would be wrong"
    )
    window = fixture[master_at:apply_at]
    for endpoint in ("PCCM_AddCostLine", "PCCM_AddRisk", "PCCM_DeleteCostLineById",
                     "PCCM_DeleteRiskById", "PCCM_Calculate",
                     "Invoke-Phase5AddDriverAndRequireSuccess",
                     "Invoke-Phase5ProductionOperation"):
        assert endpoint not in window, (
            f"{endpoint} is invoked while the master and the inflation grid disagree"
        )
    # And production really does rebuild the grid from the master during Apply,
    # which is why moving the Apply is the fix rather than a workaround.
    inflation = _text(SRC_VBA / "modInflation.bas")
    sync = inflation[inflation.index("Public Sub SyncProfileRows"):]
    sync = sync[:sync.index("\nEnd Sub")]
    assert "TBL_INFLATION_PROFILES" in sync and "SH_CONFIG" in sync
    assert "modInflation.SyncProfileRows" in _text(SRC_VBA / "modTimeline.bas")
    # And the structural check really is what refuses in the window.
    structural = _text(SRC_VBA / "modStructuralCheck.bas")
    assert "CheckInflationProfiles" in structural
    check = structural[structural.index("Private Function CheckInflationProfiles"):]
    check = check[:check.index("\nEnd Function")]
    assert "modInflation.ProfileNameSet()" in check, (
        "the structural check no longer compares the grid against the Config master"
    )
    drivers = _vba_executable(SRC_VBA / "modDrivers.bas")
    assert "modStructuralCheck.ValidateStructure()" in drivers, (
        "a driver Add no longer revalidates the structure, so the reasoning needs restating"
    )


def test_158_r3_no_production_endpoint_is_invoked_with_its_result_discarded() -> None:
    """R3. `$Excel.Run('PCCM_Add...') | Out-Null` is the defect itself.

    Every mutating endpoint the Phase-5 harness drives goes through the checked
    helper. A bare invocation whose result is piped away is the exact shape that
    let a failed Add look like a successful one.
    """
    source = _executable(SCENARIOS)
    for endpoint in ("PCCM_AddCostLine", "PCCM_AddRisk", "PCCM_DeleteCostLineById",
                     "PCCM_DeleteRiskById", "PCCM_ApplyTimeline"):
        for shape in (f"$Excel.Run('{endpoint}') | Out-Null",
                      f"$Excel.Run('{endpoint}', $id) | Out-Null",
                      f"$Excel.Run('{endpoint}', [string]$id) | Out-Null"):
            assert shape not in source, (
                f"a production mutation is still invoked as {shape}"
            )
    # The endpoints are named ONLY where they are dispatched through the helper.
    for endpoint in ("PCCM_AddCostLine", "PCCM_AddRisk"):
        add_helper = _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess")
        assert f"'{endpoint}'" in add_helper, f"{endpoint} is not dispatched by the checked Add"
    clear = _procedure(source, "Clear-Phase5Registers")
    for endpoint in ("PCCM_DeleteCostLineById", "PCCM_DeleteRiskById"):
        assert f"'{endpoint}'" in clear, f"{endpoint} is not dispatched by the checked clear"


def test_159_r4_every_checked_operation_requires_an_ok_result() -> None:
    """R4. PCCM_AutomationResult is the authority, not the fact that Run returned."""
    source = _executable(SCENARIOS)
    checked = _procedure(source, "Invoke-Phase5ProductionOperation")
    assert "$Excel.Run('PCCM_AutomationResult')" in checked, (
        "the checked helper never reads the operation result"
    )
    assert "$result -notlike 'OK|*'" in checked, (
        "the checked helper accepts a result that is not a success"
    )
    assert "throw (" in checked, "a failed operation does not stop the caller"
    # The failure text names the operation AND what production said, so the
    # report identifies the statement that failed rather than a later symptom.
    assert "$Operation" in checked and "$result" in checked and "$Stage" in checked
    # NOTHING downgrades the failure to a note or a warning.
    for soft in ("Add-Note", "Write-Warning", "Write-Host", "return $false"):
        assert soft not in checked, f"the checked helper softens a failure with {soft}"


def test_160_r5_the_result_is_cleared_before_the_operation_runs() -> None:
    """R5. gAutomationLastResult is a global that survives until overwritten.

    An endpoint that fails BEFORE reaching RecordResult leaves the previous
    operation's `OK|...` in place, so a reader that did not clear first would
    accept a stale success as this operation's own. PCCM_AutomationBegin calls
    ClearAutomation, and the accepted Phase-4 Set-AppliedTimeline has used that
    idiom since the matrix was written.
    """
    source = _executable(SCENARIOS)
    checked = _procedure(source, "Invoke-Phase5ProductionOperation")
    begin_at = checked.index("$Excel.Run('PCCM_AutomationBegin', $true, '')")
    run_at = checked.index("$Excel.Run($Operation")
    read_at = checked.index("$Excel.Run('PCCM_AutomationResult')")
    assert begin_at < run_at < read_at, (
        "the result is not cleared before the operation it is supposed to describe"
    )
    # Production really does clear it there.
    app_state = _vba_executable(SRC_VBA / "modAppState.bas")
    begin = app_state[app_state.index("Public Sub PCCM_AutomationBegin"):]
    begin = begin[:begin.index("End Sub")]
    assert "ClearAutomation" in begin, (
        "PCCM_AutomationBegin no longer clears the last result"
    )
    clear_sub = app_state[app_state.index("Sub ClearAutomation"):]
    clear_sub = clear_sub[:clear_sub.index("End Sub")]
    assert "gAutomationLastResult = vbNullString" in clear_sub, (
        "ClearAutomation no longer resets the recorded result"
    )
    # And the accepted Phase-4 helper is the precedent, not an invention here.
    applied = _procedure(_executable(HARNESS), "Set-AppliedTimeline")
    assert applied.index("PCCM_AutomationBegin") < applied.index("PCCM_ApplyTimeline")

    # THE ARITHMETIC OF THE HOLE, modelled.
    def read(last: str, cleared: bool, recorded: str | None) -> str:
        state = "" if cleared else last
        return recorded if recorded is not None else state

    # An endpoint that raised before RecordResult, after a previous success.
    assert read("OK|Cost Line CL-001 added.", cleared=False, recorded=None).startswith("OK|"), (
        "without the clear, a stale success is read as this operation's result"
    )
    assert read("OK|Cost Line CL-001 added.", cleared=True, recorded=None) == "", (
        "with the clear, an operation that recorded nothing reads as nothing"
    )
    assert not "".startswith("OK|"), "an empty result must not satisfy the OK gate"


def test_161_r6_the_key_is_proved_before_any_fixture_data_is_written() -> None:
    """R6. The postcondition that would have stopped Run 5 dead."""
    source = _executable(SCENARIOS)
    add = _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess")
    invoke_at = add.index("Invoke-Phase5ProductionOperation")
    ids_at = add.index("Get-IdColumnValues")
    body_at = add.index("Get-TableBody")
    write_at = add.index("Write-Phase5Driver")
    assert invoke_at < ids_at < body_at < write_at, (
        "the Add postconditions do not all precede the driver write"
    )
    assert "IsNullOrWhiteSpace($issued)" in add, (
        "the row is not proved to carry a permanent identifier at all"
    )
    assert add.index("IsNullOrWhiteSpace($issued)") < write_at, (
        "fixture data is written before the row is proved to be keyed"
    )
    # The register grew by exactly one keyed row per Add.
    assert "$ids.Count -ne $RowIndex" in add, (
        "the register is not proved to hold exactly one keyed row per Add"
    )


def test_162_r7_the_issued_identifier_is_the_emitted_one_binary() -> None:
    """R7. Not merely `a key`, but THE key the corpus names, case-sensitively."""
    source = _executable(SCENARIOS)
    add = _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess")
    assert "$Driver.permanent_id" in add, "the emitted identifier is never consulted"
    assert "$issued -cne $expected" in add, (
        "the identifier comparison is not binary; 'cl-001' would pass as 'CL-001'"
    )
    assert "-ne $expected" not in add.replace("-cne $expected", ""), (
        "a case-insensitive identifier comparison survives"
    )
    # A driver with no emitted identifier is a fixture that cannot be checked,
    # and that is a failure rather than a skipped check.
    assert "IsNullOrEmpty($expected)" in add, (
        "a driver with no permanent_id silently skips the identity proof"
    )

    # AND THE CORPUS REALLY DOES NAME THEM IN ISSUE ORDER, which is what makes
    # the Nth-Add-issues-the-Nth-id postcondition a proof rather than a guess.
    cases = _emitted()["cases"]
    models = 0
    for case in cases["plan_cases"]:
        model = case.get("model")
        if not model:
            continue
        models += 1
        costs = [d["permanent_id"] for d in model.get("cost_lines", [])]
        risks = [d["permanent_id"] for d in model.get("risks", [])]
        assert costs == [f"CL-{i:03d}" for i in range(1, len(costs) + 1)], (
            f"plan case {case['id']} names cost identifiers out of issue order: {costs}"
        )
        assert risks == [f"R-{i:03d}" for i in range(1, len(risks) + 1)], (
            f"plan case {case['id']} names risk identifiers out of issue order: {risks}"
        )
    assert models >= 20, f"only {models} plan cases carry a model"
    # And the counters really are reset to an initial the first Add turns into 1.
    counters = _emitted()["manifest"]["counters"]
    assert counters, "the manifest declares no identity counters"
    for counter in counters:
        assert int(counter["initial"]) == 0, (
            f"{counter['defined_name']} does not reset to 0, so CL-001 is not the first issue"
        )


def test_163_r8_write_phase5_driver_has_exactly_one_call_site() -> None:
    """R8. The check cannot be bypassed by reaching past it."""
    source = _executable(SCENARIOS)
    calls = [line for line in source.splitlines()
             if "Write-Phase5Driver" in line and "function Write-Phase5Driver" not in line]
    assert len(calls) == 1, (
        f"Write-Phase5Driver is called from {len(calls)} places: {calls}"
    )
    # And that one site is INSIDE the checked Add, by position in the file.
    add_at = source.index("function Invoke-Phase5AddDriverAndRequireSuccess")
    add_end = add_at + len(_procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess"))
    site_at = source.index(calls[0])
    assert add_at < site_at < add_end, (
        "the single Write-Phase5Driver call site is outside the checked Add"
    )


def test_164_r9_every_delete_is_checked_and_proved_to_have_taken() -> None:
    """R9. A register that refuses to empty must not reach the next fixture."""
    source = _executable(SCENARIOS)
    clear = _procedure(source, "Clear-Phase5Registers")
    assert "Invoke-Phase5ProductionOperation" in clear, (
        "the delete endpoints are still invoked without a result check"
    )
    assert "$remaining -contains $id" in clear, (
        "a delete that reported success is not proved to have removed the row"
    )
    # WHY that second check exists: DeleteDriver answers a declined confirmation
    # with a SUCCESS that removes nothing, so OK| alone does not mean gone.
    drivers = _text(SRC_VBA / "modDrivers.bas")
    delete = drivers[drivers.index("Public Function DeleteDriver"):]
    delete = delete[:delete.index("\nEnd Function")]
    assert "If Not modAppState.AskConfirm" in delete and "Succeeded(vbNullString)" in delete, (
        "a declined delete no longer reports success, so the reasoning needs restating"
    )


def test_165_r10_the_registers_are_proved_empty_and_free_of_unkeyed_data() -> None:
    """R10. Contamination is EXPOSED, at both ends, and never repaired."""
    source = _executable(SCENARIOS)
    clear = _procedure(source, "Clear-Phase5Registers")
    assert clear.count("Assert-Phase5NoUnkeyedRegisterData") == 2, (
        "the unkeyed-data scan does not run both before and after the emptying"
    )
    assert "$left.Count -ne 0" in clear, (
        "the register is not proved to hold no identifier afterwards"
    )
    scan = _procedure(source, "Assert-Phase5NoUnkeyedRegisterData")
    assert "throw (" in scan, "an orphan row does not stop the fixture"
    # THE HARNESS MUST EXPOSE CONTAMINATION, NOT LAUNDER IT. No repair path.
    for repair in ("Set-TableCell", "Remove-TableRow", "Set-Phase5TypedCell",
                   "Add-BlankTableRow", "$Excel.Run"):
        assert repair not in scan, (
            f"the orphan scan {repair}s the row it found instead of reporting it"
        )
    # The predicate is production's own, term for term.
    workbook = _text(SRC_VBA / "modWorkbook.bas")
    orphan = workbook[workbook.index("Public Function OrphanRows"):]
    orphan = orphan[:orphan.index("\nEnd Function")]
    assert "Len(TextOf(CellIn(Target, r, KeyColumn))) = 0" in orphan and \
           "Not IsEmptyCell(CellIn(Target, r, c))" in orphan, (
        "production's orphan predicate changed; the harness mirror needs restating"
    )
    assert "IsNullOrWhiteSpace" in scan, (
        "the mirror does not use a trimmed-empty test, so it is not production's predicate"
    )


def test_166_r11_the_identity_counters_are_proved_typed_and_exact() -> None:
    """R11. modDrivers.TryReadCounter refuses a counter that is not a number."""
    source = _executable(SCENARIOS)
    clear = _procedure(source, "Clear-Phase5Registers")
    assert "Get-Phase5TypedNamedValue" in clear, (
        "the counter is read back through the stringifying reader"
    )
    assert "Test-Phase5ExactValue" in clear, (
        "the counter read-back does not use the strict comparator"
    )
    assert "Get-NamedValue -Workbook $Workbook -DefinedName $counter.defined_name)" not in clear
    typed = _procedure(source, "Get-Phase5TypedNamedValue")
    assert "return $rng.Value2" in typed, "the typed named reader does not return Value2"
    body = typed[typed.index("try {"):]
    for lossy in ("[string]", "Format", "IsNullOrEmpty", "[double]"):
        assert lossy not in body, (
            f"the typed named reader {lossy}s the value on the way out"
        )
    # Production's own refusal is what makes a text counter fatal.
    drivers = _vba_executable(SRC_VBA / "modDrivers.bas")
    assert "IsWholeInRange(raw, 0, ID_COUNTER_MAX, d)" in drivers, (
        "TryReadCounter no longer requires a whole number; the reasoning needs restating"
    )


def test_167_r12_coherence_is_proved_at_both_ends_of_the_fixture() -> None:
    """R12. Every production mutation begins and ends structurally coherent."""
    source = _executable(SCENARIOS)
    fixture = _fixture_steps(source)
    assert fixture.count("Assert-Phase5StructurallyCoherent") == 2, (
        "the fixture does not prove coherence both after the Apply and at the end"
    )
    first = fixture.index("Assert-Phase5StructurallyCoherent")
    last = fixture.rindex("Assert-Phase5StructurallyCoherent")
    add_at = fixture.index("Invoke-Phase5AddDriverAndRequireSuccess")
    weights_at = fixture.index("Write-Phase5Weights")
    assert first < add_at, "the first coherence gate is after the driver Adds"
    assert weights_at < last, "the closing coherence gate is before the value writers"
    coherent = _procedure(source, "Assert-Phase5StructurallyCoherent")
    assert "$Excel.Run('PCCM_StructuralReport')" in coherent, (
        "coherence is judged by something other than production"
    )
    # And PCCM_StructuralReport really is ValidateStructure.
    structural = _text(SRC_VBA / "modStructuralCheck.bas")
    assert "PCCM_StructuralReport = ValidateStructure()" in structural


def test_168_r13_a_failed_fixture_postcondition_throws() -> None:
    """R13. Fixture establishment fails LOUDLY or not at all.

    A postcondition that returned a diagnostic, wrote a note or continued to the
    next driver would put the run back where Run 5 was: a broken baseline
    carried into scenarios that report it as their own predicate failing.
    """
    source = _executable(SCENARIOS)
    for name in ("Invoke-Phase5ProductionOperation", "Assert-Phase5StructurallyCoherent",
                 "Assert-Phase5NoUnkeyedRegisterData",
                 "Invoke-Phase5AddDriverAndRequireSuccess", "Clear-Phase5Registers",
                 "Set-Phase5Fixture"):
        body = _procedure(source, name)
        for soft in ("Add-Note", "Write-Warning", "Write-Host", "-ErrorAction SilentlyContinue",
                     "try {"):
            assert soft not in body, (
                f"{name} softens or swallows a fixture-establishment failure ({soft})"
            )
    # Each guard actually raises.
    for name, count in (("Invoke-Phase5ProductionOperation", 1),
                        ("Assert-Phase5StructurallyCoherent", 1),
                        ("Assert-Phase5NoUnkeyedRegisterData", 1),
                        ("Invoke-Phase5AddDriverAndRequireSuccess", 5),
                        ("Clear-Phase5Registers", 3)):
        body = _procedure(source, name)
        assert body.count("throw (") >= count, (
            f"{name} has {body.count('throw (')} throws, expected at least {count}"
        )


def test_169_r14_the_fixture_self_proof_gates_the_scenarios_that_depend_on_it() -> None:
    """R14. A fixture defect gets a result of its own, before thirteen symptoms.

    Run 5 reported an ordering defect as thirteen separate scenario failures,
    none of which had reached its own predicate, and the defect itself had no
    result anywhere in the ledger.
    """
    scenarios = _executable(SCENARIOS)
    assert _result_call("P5-FIX") in scenarios, "the fixture self-proof emits no result"
    # It runs BEFORE the first scenario that establishes a fixture for itself.
    driver = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    fix_at = driver.index("Add-Phase5Result 'P5-FIX'")
    an_at = driver.index("Add-Phase5Result 'P5-AN'")
    assert fix_at < an_at, "the self-proof is recorded after the analytical scenarios"
    # THE VERY FIRST fixture the run establishes is the self-proof's own: the
    # next one after it comes only after P5-FIX has been recorded.
    first_fixture = driver.index("Set-Phase5Fixture -Excel $Excel")
    second_fixture = driver.index("Set-Phase5Fixture -Excel $Excel", first_fixture + 1)
    assert first_fixture < fix_at < second_fixture, (
        "a scenario other than the self-proof establishes the first fixture"
    )
    # It uses the REAL fixture procedure, not a copy of it.
    block = driver[driver.index("Add-Phase5Result 'P5-D8'"):an_at]
    assert "Set-Phase5Fixture -Excel $Excel" in block, (
        "the self-proof does not drive the real fixture procedure"
    )
    # A FAIL, never a SKIP, and it gates through P5-ALL exactly as P5-FX does.
    assert "'SKIP'" not in block, "a broken fixture is recorded as a SKIP"
    assert "Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL'" in block
    assert block.count("return") >= 2, (
        "a broken fixture does not stop the dependent scenarios on both paths"
    )
    # AND THE LIFECYCLE EVIDENCE SURVIVES THE GATE. Returning from the scenario
    # driver leaves the caller's shutdown, Z, Y, the final completeness gate and
    # the ledger verdict untouched - in that order, since P5-FIN is itself a
    # guarded Phase-5 result and has to be attempted before the verdict.
    harness = _executable(HARNESS)
    phase5_at = harness.index("Invoke-Phase5GateBScenarios -Excel")
    z_at = harness.index("Add-Result 'Z' 'Excel closed naturally")
    y_at = harness.index("$transient = @(Get-TransientFailures)")
    fin_at = harness.index("Add-Phase4FinalCompletenessResult -Results $results")
    ledger_at = harness.index("Add-Phase5LedgerIntegrityResult")
    assert phase5_at < z_at < y_at < fin_at < ledger_at, (
        "the gate would take the post-session lifecycle evidence down with it"
    )
    assert "exit" not in block, "the gate exits the process instead of returning"
    # And it is registered, so a coverage mapping could name it.
    registry = _procedure(scenarios, "Get-Phase5ScenarioIds")
    assert "'P5-FIX'" in registry, "the self-proof is not a scenario the harness declares"


def test_170_r15_the_run_5_defect_replayed_as_a_state_machine() -> None:
    """R15. THE MUTATION CONTROL: the exact Run-5 bug, both orders.

    Production is modelled from its own source contract - an Add revalidates the
    structure and rolls its identifier allocation back when that fails, Apply
    rebuilds the inflation grid from the Config master, and a mutation over an
    unkeyed row is refused. Neither order is asserted to be correct; the model
    is run and the outcomes are compared.
    """
    class Workbook:
        def __init__(self, applied_profiles):
            self.master = list(applied_profiles)       # Config!tblInflationProfiles
            self.grid = list(applied_profiles)         # Inflation!tblInflation
            self.rows: list[dict] = []                 # tblCostLines
            self.counter = 0
            self.result = ""

        # --- modStructuralCheck ------------------------------------------
        def orphan_rows(self) -> list[int]:
            return [i + 1 for i, row in enumerate(self.rows)
                    if not row["id"] and any(v not in (None, "") for k, v in row.items()
                                             if k != "id")]

        def validate_structure(self) -> str:
            problems = []
            if sorted(self.master) != sorted(self.grid):
                problems.append("[inflation_profile_rows] tblInflation and the Config "
                                "profile master disagree.")
            if self.orphan_rows():
                problems.append("[no_orphan_structural_data] tblCostLines row(s) "
                                + ", ".join(str(r) for r in self.orphan_rows())
                                + " hold data but carry no key.")
            return "".join(problems)

        def pre_mutation_check(self) -> str:
            return ("[no_orphan_structural_data] tblCostLines row(s) "
                    + ", ".join(str(r) for r in self.orphan_rows())
                    + " hold data but carry no key.") if self.orphan_rows() else ""

        # --- modDrivers.RunDriverOperation -------------------------------
        def add_cost_line(self) -> None:
            self.result = ""
            problems = self.pre_mutation_check()
            if problems:
                self.result = "FAIL|Add Cost Line was refused.|" + problems
                return
            before_rows, before_counter = [dict(r) for r in self.rows], self.counter
            self.counter += 1
            self.rows.append({"id": f"CL-{self.counter:03d}", "description": None})
            problems = self.validate_structure()
            if problems:
                # Failure: the register and the counter are restored exactly.
                self.rows, self.counter = before_rows, before_counter
                self.rows.append({"id": "", "description": None})   # the reserved blank row
                self.result = "FAIL|Structural revalidation failed:|" + problems
                return
            self.result = f"OK|Cost Line CL-{self.counter:03d} added."

        # --- modTimeline.PCCM_ApplyTimeline ------------------------------
        def apply_timeline(self) -> None:
            self.result = ""
            problems = self.pre_mutation_check()
            if problems:
                self.result = ("FAIL|Apply / Update Timeline was refused. Nothing has "
                               "been changed.|" + problems)
                return
            self.grid = list(self.master)               # SyncProfileRows
            self.result = "OK|Timeline applied."

    def set_master(book: Workbook, profiles) -> None:
        book.master = list(profiles)

    def write_driver(book: Workbook, row_index: int) -> None:
        book.rows[row_index - 1]["description"] = "GateB CL-001"

    fixture_profiles = ["GateBEscalation"]
    previous_profiles = ["Standard"]

    # --- THE RUN-5 ORDER: master, Add, write, Apply ----------------------
    old = Workbook(previous_profiles)
    set_master(old, fixture_profiles)
    old.add_cost_line()
    assert old.result.startswith("FAIL|"), "the Add would have succeeded; the model is wrong"
    assert "Config profile" in old.result or "disagree" in old.result
    # The harness discarded that result and wrote the data anyway.
    write_driver(old, 1)
    assert old.orphan_rows() == [1], "the orphan the harness manufactured is missing"
    old.apply_timeline()
    assert old.result.startswith("FAIL|"), "Apply did not refuse over the orphan"
    assert "no_orphan_structural_data" in old.result
    assert "tblCostLines row(s) 1 hold data but carry no key" in old.result, old.result
    assert old.counter == 0, "a failed Add left its identifier allocation behind"

    # --- THE CORRECTED ORDER: master, Apply, checked Add, write ----------
    new = Workbook(previous_profiles)
    set_master(new, fixture_profiles)
    new.apply_timeline()
    assert new.result.startswith("OK|"), new.result
    assert new.validate_structure() == "", "the baseline is not coherent before the Adds"
    new.add_cost_line()
    assert new.result.startswith("OK|"), new.result
    assert new.rows[0]["id"] == "CL-001", new.rows
    write_driver(new, 1)
    assert new.orphan_rows() == [], "the corrected order still manufactures an orphan"
    assert new.validate_structure() == "", "the fixture does not end coherent"

    # --- AND THE CHECKED ADD REFUSES TO WRITE UNDER THE OLD ORDER --------
    # This is the harness half of the correction: even with the ordering
    # regressed, no fixture data reaches an unkeyed row.
    guarded = Workbook(previous_profiles)
    set_master(guarded, fixture_profiles)
    guarded.add_cost_line()
    stopped = not guarded.result.startswith("OK|")
    assert stopped, "the checked Add would have proceeded past a failed operation"
    assert guarded.orphan_rows() == [], "data was written before the result was checked"
    # The row exists and is unkeyed; the guard is what keeps data out of it.
    assert guarded.rows and guarded.rows[0]["id"] == "", guarded.rows


# --- negative controls -----------------------------------------------------
def test_nc_100_the_run_5_ordering_is_caught() -> None:
    """The old fixture order, planted, fails R1."""
    planted = _synthetic(
        "    Set-Phase5InflationProfileMaster -Workbook $Workbook -Inspection $Inspection\n"
        "    foreach ($line in @($Model.cost_lines)) {\n"
        "        Invoke-Phase5AddDriverAndRequireSuccess -Excel $Excel\n"
        "    }\n"
        "    $applied = Invoke-Phase5ProductionOperation -Operation 'PCCM_ApplyTimeline'\n"
    )
    apply_at = planted.index("-Operation 'PCCM_ApplyTimeline'")
    add_at = planted.index("Invoke-Phase5AddDriverAndRequireSuccess")
    assert not (apply_at < add_at), "the planted regression must violate R1"
    window = planted[planted.index("Set-Phase5InflationProfileMaster"):apply_at]
    assert "Invoke-Phase5AddDriverAndRequireSuccess" in window, (
        "the planted regression must also violate R2"
    )


def test_nc_101_a_discarded_add_result_is_caught() -> None:
    """The literal Run-5 statement, planted, fails R3."""
    planted = _synthetic(
        "        $Excel.Run('PCCM_AddCostLine') | Out-Null\n"
        "        Write-Phase5Driver -Workbook $Workbook -Register $costReg\n"
    )
    assert "$Excel.Run('PCCM_AddCostLine') | Out-Null" in planted
    assert "PCCM_AutomationResult" not in planted, "the planted regression checks nothing"
    assert planted.index("Out-Null") < planted.index("Write-Phase5Driver"), (
        "the planted regression must write data behind an unchecked Add"
    )


def test_nc_102_reading_a_stale_result_is_caught() -> None:
    """A result read without clearing first, planted, fails R5."""
    planted = _synthetic(
        "    $Excel.Run($Operation) | Out-Null\n"
        "    $result = [string]$Excel.Run('PCCM_AutomationResult')\n"
    )
    assert "PCCM_AutomationBegin" not in planted, "the planted regression must not clear"
    # And it is not a theoretical hole: the previous success is what comes back.
    last = "OK|Cost Line CL-001 added."
    recorded = None                      # the endpoint raised before RecordResult
    assert (recorded if recorded is not None else last).startswith("OK|"), (
        "the stale success must be what an unguarded read returns"
    )


def test_nc_103_writing_before_the_key_proof_is_caught() -> None:
    """The driver write hoisted above the postconditions, planted, fails R6."""
    planted = _synthetic(
        "    $null = Invoke-Phase5ProductionOperation -Excel $Excel -Operation $endpoint\n"
        "    Write-Phase5Driver -Workbook $Workbook -Register $Register -RowIndex $RowIndex\n"
        "    $ids = @(Get-IdColumnValues -Workbook $Workbook -Info $Register)\n"
        "    if ($ids.Count -ne $RowIndex) { throw ('...') }\n"
    )
    assert planted.index("Write-Phase5Driver") < planted.index("Get-IdColumnValues"), (
        "the planted regression must write before it proves"
    )
    assert "IsNullOrWhiteSpace($issued)" not in planted, (
        "the planted regression must skip the key proof entirely"
    )


# --- the same contract, stated a second way --------------------------------
# Every rule above is pinned by the test that names it. These four state the
# same contract in a DIFFERENT formulation - statement order rather than
# membership, a regex sweep rather than a literal absence, a forbidden-reader
# sweep rather than a required-reader presence - so no single planted regression
# is caught by only one assertion.
def test_171_the_statement_order_of_every_checked_step_is_fixed() -> None:
    """Each fixture-establishment procedure walked as an ORDERED sequence.

    The per-rule tests above compare a handful of indices. This walks each
    procedure once and requires the whole sequence, so a step that is deleted or
    moved is caught by its position in the walk rather than by a rule that
    happens to mention it.
    """
    source = _executable(SCENARIOS)

    def walk(name: str, tokens: tuple[str, ...]) -> None:
        body = _procedure(source, name)
        cursor = 0
        for token in tokens:
            at = body.find(token, cursor)
            assert at >= 0, (
                f"{name} no longer performs '{token}' in order; the sequence broke "
                f"after position {cursor}"
            )
            cursor = at + len(token)

    # THE CHECKED OPERATION: clear, invoke, read, gate.
    walk("Invoke-Phase5ProductionOperation", (
        "$Excel.Run('PCCM_AutomationBegin', $true, '')",
        "$Excel.Run($Operation",
        "$Excel.Run('PCCM_AutomationResult')",
        "$result -notlike 'OK|*'",
        "throw (",
    ))
    # EMPTYING A REGISTER: scan, enumerate, delete, prove gone, prove empty,
    # scan again, then the counters.
    walk("Clear-Phase5Registers", (
        "Assert-Phase5NoUnkeyedRegisterData",
        "Get-IdColumnValues",
        "Invoke-Phase5ProductionOperation",
        "$remaining -contains $id",
        "$left.Count -ne 0",
        "Assert-Phase5NoUnkeyedRegisterData",
        "$Manifest.counters",
        "Get-Phase5TypedNamedValue",
        "Test-Phase5ExactValue",
    ))
    # ONE ADD: invoke, count, read the row, prove keyed, prove WHICH key, write.
    walk("Invoke-Phase5AddDriverAndRequireSuccess", (
        "Invoke-Phase5ProductionOperation",
        "$ids.Count -ne $RowIndex",
        "Get-TableBody",
        "IsNullOrWhiteSpace($issued)",
        "$issued -cne $expected",
        "Write-Phase5Driver",
    ))
    # THE FIXTURE: A to H, in order.
    walk("Invoke-Phase5FixtureSteps", (
        "Clear-Phase5Registers",
        "Reset-Phase5FxTable",
        "Set-Phase5InflationProfileMaster",
        "Invoke-Phase5ProductionOperation",
        "-Operation 'PCCM_ApplyTimeline'",
        "Assert-Phase5StructurallyCoherent",
        "Invoke-Phase5AddDriverAndRequireSuccess",
        "Write-Phase5InflationRates",
        "Write-Phase5Weights",
        "Assert-Phase5StructurallyCoherent",
        "return $applied",
    ))
    # AND THE CLOSING GATE IS THE LAST THING THAT HAPPENS. A coherence proof
    # with a mutation after it proves the state before the mutation.
    fixture = _fixture_steps(source)
    tail = fixture[fixture.rindex("Assert-Phase5StructurallyCoherent"):]
    for later in ("$Excel.Run", "Set-TableCell", "Set-NamedValue", "Write-Phase5",
                  "Invoke-Phase5"):
        assert later not in tail, (
            f"the fixture {later}s after it claims to have proved itself coherent"
        )
    # AND THE DRIVER WRITE IS THE LAST THING THE CHECKED ADD DOES.
    add = _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess")
    after_write = add[add.index("Write-Phase5Driver"):]
    assert "throw (" not in after_write, (
        "a postcondition runs after the data it was supposed to gate has been written"
    )


def test_172_no_identity_comparison_in_the_correction_is_case_insensitive() -> None:
    """An identifier is an identity: 'cl-001' is not 'CL-001'.

    A sweep over the comparison operators actually present, rather than a check
    that one particular lax spelling is absent.
    """
    source = _executable(SCENARIOS)
    # The PERMANENT IDENTIFIERS and the PROFILE NAMES - the values production
    # treats as identities. A plan-case ID and a manifest key are neither, and
    # the accepted harness compares those with -eq throughout.
    identity_holders = ("$issued", "$expected", "$declared",
                        "$gridRows", "$emittedProfiles")
    regions = {
        "Invoke-Phase5AddDriverAndRequireSuccess":
            _procedure(source, "Invoke-Phase5AddDriverAndRequireSuccess"),
        "the P5-FIX self-proof":
            source[source.index("Add-Phase5Result 'P5-D8'"):
                   source.index("Add-Phase5Result 'P5-AN'")],
    }
    lax = re.compile(r"(\$[A-Za-z0-9_.\[\]$-]+)\s+-(eq|ne|ceq|cne)\s+(\S+)")
    for label, body in regions.items():
        for left, operator, right in lax.findall(body):
            touches_identity = any(
                holder in left or holder in right for holder in identity_holders)
            # A COUNT comparison is not an identity comparison.
            if ".Count" in left or ".Count" in right:
                continue
            if not touches_identity:
                continue
            assert operator in ("ceq", "cne"), (
                f"{label} compares identities case-insensitively: "
                f"{left} -{operator} {right}"
            )
    # The self-proof's own identifier comparison is binary too.
    assert "-ceq ($expected -join" in regions["the P5-FIX self-proof"], (
        "the self-proof compares the issued identifier list case-insensitively"
    )


def test_173_the_fixture_reads_identity_state_through_typed_readers() -> None:
    """A stringifying reader cannot tell a numeric 0 from the text "0".

    Stated as a sweep for the LOSSY readers, which is the opposite direction
    from asserting that the typed one is present.
    """
    source = _executable(SCENARIOS)
    clear = _procedure(source, "Clear-Phase5Registers")
    counters = clear[clear.index("$Manifest.counters"):]
    for lossy in ("Get-NamedValue", "[string]$readBack", "[double]$readBack",
                  "Format-CalcValue", "Test-CalcValue"):
        assert lossy not in counters, (
            f"the counter read-back goes through {lossy}, which cannot see the type"
        )
    assert "Get-Phase5TypedNamedValue" in counters
    assert "Test-Phase5ExactValue" in counters
    # The strict comparator really is type-first, so using it is the whole claim.
    comparator = _procedure(source, "Test-Phase5ExactValue")
    assert "GetType().FullName -cne" in comparator, (
        "the strict comparator no longer gates on CLR type identity"
    )


def test_174_no_phase5_gate_downgrades_a_failure_to_a_skip() -> None:
    """A prerequisite that did not hold is a FAIL, and it stops the run below it.

    Every `P5-ALL` emission in the harness is a gate. A SKIP would leave the
    dependent scenarios running against a baseline nobody proved, and a gate
    that did not return would run them anyway.
    """
    source = _executable(SCENARIOS)
    emissions = [match for match in re.finditer(
        r"Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' '(\w+)'", source)]
    assert len(emissions) >= 4, f"only {len(emissions)} P5-ALL gates exist"
    for match in emissions:
        assert match.group(1) == "FAIL", (
            f"a P5-ALL gate reports {match.group(1)} instead of FAIL"
        )
        # ... and the next statement of substance is the return.
        tail = source[match.end():match.end() + 700]
        assert "return" in tail, "a P5-ALL gate does not stop the scenarios below it"
        before_return = tail[:tail.index("return")]
        for runs_on in ("$Excel.Run", "Set-Phase5Fixture", "Add-Check"):
            assert runs_on not in before_return, (
                f"a P5-ALL gate {runs_on}s before it returns"
            )
    # And the self-proof is one of them.
    fix_block = source[source.index("Add-Phase5Result 'P5-D8'"):
                       source.index("Add-Phase5Result 'P5-AN'")]
    assert fix_block.count("Add-Phase5Result 'P5-ALL'") == 2, (
        "the fixture self-proof does not gate on both its success and its catch path"
    )


# ===========================================================================
# 24. RUNTIME RUN 6: the empty rate object
# ===========================================================================
# Run 6 carried the Run-5 fixture choreography onto real Excel and reached
# step G, which is runtime evidence that the checked Adds and the structural
# Apply before it did what they claim. P5-FIX then failed inside
# Write-Phase5InflationRates on
#
#     foreach ($year in $rates.PSObject.Properties.Name)
#
# with "The property 'Name' cannot be found on this object", because `$rates` is
# legitimately `{}` for the golden plan case. `{}` is the CORRECT encoding of
# "this timeline needs no inflation calendar year", not malformed data.
def _rate_maps() -> list[tuple[str, str, dict]]:
    """(plan case id, profile name, rate mapping) for every emitted model."""
    out = []
    for case in _emitted()["cases"]["plan_cases"]:
        model = case.get("model")
        if not model:
            continue
        for name, rates in (model.get("inflation") or {}).items():
            out.append((str(case["id"]), str(name), rates))
    return out


def test_175_r1_the_golden_plan_case_carries_an_empty_rate_object() -> None:
    """R1. `inflation["Standard"] == {}` for plan case 1, and it is intentional.

    Intentional is proved from the TIMELINE, not asserted: the inflation grid
    spans Base Year + 1 .. Start Year + Duration - 1, and for case 1 that span
    is empty, so there is no calendar year a rate could belong to.
    """
    cases = _emitted()["cases"]
    golden = next(case for case in cases["plan_cases"] if str(case["id"]) == "1")
    timeline = golden["model"]["timeline"]
    assert timeline == {"base_year": 2026, "start_year": 2026, "duration": 1}, timeline
    first_year = int(timeline["base_year"]) + 1
    last_year = int(timeline["start_year"]) + int(timeline["duration"]) - 1
    assert last_year < first_year, (
        f"case 1 does span calendar years {first_year}..{last_year}, so an empty "
        "rate map would NOT be the correct encoding"
    )
    assert golden["model"]["inflation"] == {"Standard": {}}, golden["model"]["inflation"]

    # AND THE EMPTINESS IS EXACTLY THE ZERO-SPAN CASES, corpus-wide. A rate map
    # that were empty for any other reason would be a corpus defect, not a
    # harness one.
    for case in cases["plan_cases"]:
        model = case.get("model")
        if not model:
            continue
        timeline = model["timeline"]
        first_year = int(timeline["base_year"]) + 1
        last_year = int(timeline["start_year"]) + int(timeline["duration"]) - 1
        span = max(last_year - first_year + 1, 0)
        for name, rates in (model.get("inflation") or {}).items():
            assert (len(rates) == 0) == (span == 0), (
                f"plan case {case['id']} profile {name}: {len(rates)} rate(s) for a "
                f"span of {span} calendar year(s)"
            )


def test_176_r2_the_writer_treats_an_empty_rate_object_as_zero_assignments() -> None:
    """R2. Enumerating the property COLLECTION, never projecting its `.Name`."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    assert "foreach ($rateProperty in $rates.PSObject.Properties)" in body, (
        "the rate mapping is not enumerated as property objects"
    )
    # There is no guard, no count test and no early return: zero properties
    # simply means zero iterations, which is what the loop already does.
    for crutch in ("$rates.PSObject.Properties.Count", "if ($null -eq $rates)",
                   "Get-Member", "-eq 0) { continue }"):
        assert crutch not in body, (
            f"the writer special-cases the empty rate object with {crutch} instead "
            "of enumerating a collection that is simply empty"
        )


def test_177_r3_no_fixture_invents_an_inflation_year_to_avoid_the_empty_case() -> None:
    """R3. The corpus is not repaired to make the loop non-empty."""
    cases = _emitted()["cases"]
    empty = [(cid, name) for cid, name, rates in _rate_maps() if not rates]
    assert len(empty) == 11, f"the empty-rate cases changed: {empty}"
    assert ("1", "Standard") in empty, "the golden case no longer carries the empty map"
    # And nothing in the harness writes a rate the corpus did not emit.
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    assert not re.search(r"-Value \(\[double\]0", body), "the writer plants a zero rate"
    assert "$rateValue" in body and "$Model.inflation" in body, (
        "the writer no longer takes its rates from the emitted model"
    )
    # The corpus itself is untouched by this round.
    golden = next(case for case in cases["plan_cases"] if str(case["id"]) == "1")
    assert golden["model"]["inflation"]["Standard"] == {}


def test_178_r4_the_empty_rate_path_still_locates_the_profile_row() -> None:
    """R4. An empty rate map does not skip the profile.

    The OUTER loop is over profiles and the inner one over that profile's rates.
    An empty rate map therefore still runs the outer body - the row is still
    located, and the profile still exists in the Config master and the grid.
    """
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    outer = body.index("foreach ($profileProperty in $Model.inflation.PSObject.Properties)")
    find = body.index("Find-GridRow")
    inner = body.index("foreach ($rateProperty in $rates.PSObject.Properties)")
    assert outer < find < inner, (
        "the profile row is located inside the rate loop, so an empty rate map "
        "would skip the profile entirely"
    )
    # And the profile identity still reaches the Config master, which is what
    # creates the grid row in the first place.
    fixture = _fixture_steps(source)
    assert "Set-Phase5InflationProfileMaster" in fixture
    assert "$Model.inflation.PSObject.Properties.Name" in fixture, (
        "the profile master is no longer driven from the emitted profile set"
    )


def test_179_r5_r6_r7_r8_the_enumeration_semantics_modelled() -> None:
    """R5-R8. One rate, several rates, a null rate, and the empty object.

    A pure-Python model of the INTENDED behaviour. It proves what the corrected
    loop is supposed to do with each shape and that `{}` and `{"2028": null}`
    are different paths. It does NOT claim to reproduce the Windows PowerShell
    property adapter: whether the real adapter yields zero properties for an
    empty PSCustomObject is a runtime fact, and the next Windows run is where
    that is proved.
    """
    HEADERS = ["2027", "2028", "2029"]

    def write_rates(rates: dict) -> list[tuple[str, object]]:
        """The corrected loop: enumerate properties, read Name and Value."""
        written: list[tuple[str, object]] = []
        for year, value in rates.items():          # zero properties -> zero passes
            assert year in HEADERS, f"no generated inflation column for {year}"
            written.append((year, None if value is None else float(value)))
        return written

    # R5: one property, one rate.
    assert write_rates({"2028": 0.05}) == [("2028", 0.05)]
    # R6: several properties, each written once.
    assert write_rates({"2027": 0.05, "2028": 0.06, "2029": 0.07}) == [
        ("2027", 0.05), ("2028", 0.06), ("2029", 0.07)]
    # R7: a property whose value is null still ENTERS the loop and writes BLANK.
    assert write_rates({"2028": None}) == [("2028", None)]
    # R8: `{}` and `{"2028": null}` are different paths.
    assert write_rates({}) == []
    assert write_rates({"2028": None}) != write_rates({})
    assert len(write_rates({})) == 0, "the empty object wrote a cell"
    assert len(write_rates({"2028": None})) == 1, "the null rate wrote no cell"
    # And the blank is a genuine blank, never a zero. [double]$null is 0.0 in
    # PowerShell, which is the Run-4 defect this must not reintroduce.
    assert write_rates({"2028": None})[0][1] is None
    assert write_rates({"2028": 0})[0][1] == 0.0
    assert write_rates({"2028": None})[0][1] != write_rates({"2028": 0})[0][1]

    # THE TWO SHAPES BOTH EXIST IN THE CORPUS, so neither path is hypothetical.
    maps = dict(((cid, name), rates) for cid, name, rates in _rate_maps())
    assert maps[("1", "Standard")] == {}
    assert maps[("14", "Standard")] == {"2027": 0.05, "2028": None, "2029": 0.05}
    assert None in maps[("14", "Standard")].values(), (
        "plan case 14 no longer carries the blank required rate"
    )


def test_180_r9_r10_the_failing_expression_is_gone_from_the_empty_capable_site() -> None:
    """R9 and R10. No `.Properties.Name` projection over a rate mapping."""
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    assert "$rates.PSObject.Properties.Name" not in body, (
        "the exact Run-6 expression is still present"
    )
    # Names come from INDIVIDUAL property objects, in both loops.
    assert "$name = [string]$profileProperty.Name" in body
    assert "$year = [string]$rateProperty.Name" in body
    # And the value comes off the SAME property object, so no dynamic lookup
    # can disagree with the name that selected it.
    assert "$rates = $profileProperty.Value" in body
    assert "$rateValue = $rateProperty.Value" in body
    for lookup in ("$rates.$year", "$Model.inflation.$name"):
        assert lookup not in body, f"the writer still resolves {lookup} dynamically"


def test_181_r11_the_run_6_shape_does_not_raise_in_the_modelled_enumeration() -> None:
    """R11. `Standard -> {}` driven end to end through the modelled loop.

    Again: this proves the INTENDED semantics of the corrected shape. The real
    PowerShell property adapter is proved by the next Windows run, not here.
    """
    written: list[tuple[str, str, object]] = []
    model_inflation = {"Standard": {}}          # the exact Run-6 shape
    for profile, rates in model_inflation.items():
        located = f"row for {profile}"
        assert located, "the profile row must still be located"
        for year, value in rates.items():
            written.append((profile, year, value))
    assert written == [], "the empty rate map wrote something"

    # And the same walk over the golden model as the builder actually emits it.
    golden = next(case for case in _emitted()["cases"]["plan_cases"]
                  if str(case["id"]) == "1")
    seen = 0
    for profile, rates in golden["model"]["inflation"].items():
        assert isinstance(rates, dict), f"{profile} maps to {type(rates).__name__}"
        for _year, _value in rates.items():
            seen += 1
    assert seen == 0, f"the golden model yielded {seen} rate assignments"


def test_182_r12_the_run_6_gate_is_unchanged() -> None:
    """R12. P5-FIX failing still means P5-ALL not attempted, Y/Z/LDG/FIN alive.

    Run 6 proved the gate itself works. Nothing in this round may weaken it.
    """
    scenarios = _executable(SCENARIOS)
    driver = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    block = driver[driver.index("Add-Phase5Result 'P5-D8'"):
                   driver.index("Add-Phase5Result 'P5-AN'")]
    assert "Add-Phase5Result 'P5-FIX'" in block
    assert block.count("Add-Phase5Result 'P5-ALL'") == 2, (
        "the self-proof no longer gates on both its success and its catch path"
    )
    assert "'SKIP'" not in block, "a broken fixture became a SKIP"
    assert block.count("return") >= 2
    # The catch path carries the message Run 6 actually recorded.
    assert ("'not attempted: Gate-B fixture establishment failed on the golden plan case'"
            in block), "the Run-6 gate message changed"
    # AND THE LIFECYCLE EVIDENCE IS STILL DOWNSTREAM OF THE GATE.
    harness = _executable(HARNESS)
    order = [harness.index(token) for token in (
        "Invoke-Phase5GateBScenarios -Excel",
        "Add-Result 'Z' 'Excel closed naturally",
        "$transient = @(Get-TransientFailures)",
        "Add-Phase4FinalCompletenessResult -Results $results",
        "Add-Phase5LedgerIntegrityResult")]
    assert order == sorted(order), (
        "the gate would take Z, Y, P5-FIN or P5-LDG down with it"
    )


def test_183_every_property_name_projection_is_classified() -> None:
    """THE AUDIT, as an executable ledger rather than a note in a document.

    Every executable `.PSObject.Properties.Name` site in the harness is listed
    with the container it enumerates and whether that container can legally be
    empty. A site that is only PROVEN non-empty is pinned against the corpus
    here, so a future case that makes it empty-capable fails this test instead
    of failing Run 7.
    """
    source = _executable(SCENARIOS)
    projections = re.findall(r"([\w$.\[\]()]*)\.PSObject\.Properties\.Name", source)
    # `@(` and `(` are wrapping syntax, not part of the container expression.
    containers = sorted({token.lstrip("@(") for token in projections})
    # The classified set. Anything new must be classified before it can ship.
    classified = {
        "$rows",                            # inspection scalar_blocks.<block>.rows
        "$Model.inflation",                 # the profile map
        "$golden.model.inflation",          # the same map, in P5-FIX
        "$wanted",                          # one emitted expected-row object
        "$expected.resolved_fx",
        "$expected.totals",
        "$Inspection.calc.tables",
    }
    assert set(containers) == classified, (
        f"an unclassified .Properties.Name projection appeared: "
        f"{sorted(set(containers) - classified)}"
    )
    # The rate mapping is NOT among them any more.
    assert "$rates" not in containers, (
        "the empty-capable rate mapping is still projected"
    )

    # --- and the non-empty claims, proved against the emitted artifacts ------
    emitted = _emitted()
    cases, inspection = emitted["cases"], emitted["inspection"]

    for block, spec in inspection["calc"]["scalar_blocks"].items():
        assert spec["rows"], f"scalar block {block} has no rows"
    assert inspection["calc"]["tables"], "the inspection projects no calc tables"

    models = 0
    for case in cases["plan_cases"]:
        model = case.get("model")
        if model is not None:
            models += 1
            # $Model.inflation is non-empty BECAUSE every model has at least one
            # driver and every driver names a profile. That derivation is what
            # makes it class B rather than "empty so far".
            drivers = (model.get("cost_lines") or []) + (model.get("risks") or [])
            assert drivers, f"plan case {case['id']} has no driver"
            referenced = {str(d["inflation_profile"]) for d in drivers}
            assert set(model["inflation"]) == referenced, (
                f"plan case {case['id']}: profiles {sorted(model['inflation'])} "
                f"but drivers reference {sorted(referenced)}"
            )
            assert model["inflation"], f"plan case {case['id']} has no profile"
        expected = case.get("expected")
        if not expected:
            continue
        assert expected["totals"], f"plan case {case['id']} emits no totals"
        assert expected["resolved_fx"], f"plan case {case['id']} emits no resolved_fx"
        for key in ("calc_years", "resolved_fx_rows", "drivers", "annual"):
            for index, wanted in enumerate(expected.get(key) or []):
                assert wanted, f"plan case {case['id']} {key}[{index}] is empty"
    assert models == 28, f"the model count changed: {models}"


# --- negative controls -----------------------------------------------------
def test_nc_104_the_run_6_expression_is_caught() -> None:
    """The literal failing line, planted, fails R9."""
    planted = _synthetic(
        "        foreach ($year in $rates.PSObject.Properties.Name) {\n"
        "            $ordinal = [array]::IndexOf($headers, [string]$year) + 1\n"
    )
    assert "$rates.PSObject.Properties.Name" in planted, (
        "the planted regression must contain the failing projection"
    )
    assert "$rateProperty" not in planted, "the planted regression must not enumerate objects"
    # And it is a MEMBER ENUMERATION over a collection, which is the shape that
    # cannot answer when the collection is empty. Modelled:
    def project_name(collection: list) -> list:
        if not collection:
            raise AttributeError("The property 'Name' cannot be found on this object.")
        return [item["Name"] for item in collection]

    assert project_name([{"Name": "2028"}]) == ["2028"]
    try:
        project_name([])
    except AttributeError as error:
        assert "cannot be found" in str(error)
    else:
        raise AssertionError("the empty projection must be the failing path")
    # Enumerating the collection itself has no such edge.
    assert [item["Name"] for item in []] == []


def test_nc_105_collapsing_the_null_rate_into_the_empty_object_is_caught() -> None:
    """A writer that skipped null-valued rates, planted, fails R7 and R8."""
    def broken(rates: dict) -> list:
        return [(y, float(v)) for y, v in rates.items() if v is not None]

    assert broken({}) == []
    assert broken({"2028": None}) == [], (
        "the planted regression must make the two shapes indistinguishable"
    )
    assert broken({}) == broken({"2028": None}), "the regression must be visible"
    # Which is exactly the case-14 damage: a required rate that is meant to be
    # BLANK would simply never be written, and the refusal could not fire.
    maps = dict(((cid, name), rates) for cid, name, rates in _rate_maps())
    assert None in maps[("14", "Standard")].values()
    assert len(broken(maps[("14", "Standard")])) == 2, (
        "case 14 would lose its blank year under the planted regression"
    )
    assert len(maps[("14", "Standard")]) == 3


def test_nc_106_a_count_guard_instead_of_an_enumeration_is_caught() -> None:
    """Special-casing the empty object, planted, fails R2.

    A guard would work and is still wrong: it leaves the projection in place for
    every non-empty case and puts the harness one corpus change away from the
    same failure somewhere else.
    """
    planted = _synthetic(
        "        if ($rates.PSObject.Properties.Count -gt 0) {\n"
        "            foreach ($year in $rates.PSObject.Properties.Name) {\n"
    )
    assert "$rates.PSObject.Properties.Count" in planted
    assert "$rates.PSObject.Properties.Name" in planted, (
        "the planted guard leaves the failing projection in place"
    )
    assert "foreach ($rateProperty in $rates.PSObject.Properties)" not in planted


def test_184_the_inflation_writer_walks_as_a_fixed_sequence() -> None:
    """The corrected writer stated as an ORDERED walk, not as membership.

    The rules above compare a handful of indices and check for the absence of
    the failing projection. This walks the procedure once and requires the whole
    sequence, so a step that is deleted, hoisted or replaced is caught by its
    position rather than by a rule that happens to mention it.
    """
    source = _executable(SCENARIOS)
    body = _procedure(source, "Write-Phase5InflationRates")
    cursor = 0
    for token in (
        # the stale-rate blanking pass, which must still run first
        "-ColumnIndex $column -Value $null",
        # the profile loop, over property OBJECTS
        "foreach ($profileProperty in $Model.inflation.PSObject.Properties)",
        "$name = [string]$profileProperty.Name",
        "$rates = $profileProperty.Value",
        # the row is located ONCE per profile, outside the rate loop, so an
        # empty rate map still resolves its profile
        "$rowIndex = Find-GridRow -Workbook $Workbook -Grid $grid -Key $name",
        # the rate loop, over property OBJECTS
        "foreach ($rateProperty in $rates.PSObject.Properties)",
        "$year = [string]$rateProperty.Name",
        "$rateValue = $rateProperty.Value",
        "[array]::IndexOf($headers, $year)",
        "throw (",
        # the blank branch WRITES A BLANK; it does not skip the year
        "if ($null -eq $rateValue) {",
        "-ColumnIndex $ordinal -Value $null",
        "} else {",
        "-ColumnIndex $ordinal -Value ([double]$rateValue)",
    ):
        at = body.find(token, cursor)
        assert at >= 0, (
            f"Write-Phase5InflationRates no longer performs '{token}' in order; "
            f"the sequence broke after position {cursor}"
        )
        cursor = at + len(token)

    # A NULL RATE IS WRITTEN, NEVER SKIPPED. `continue` in the blank branch
    # would make `{"2028": null}` indistinguishable from `{}` at the worksheet.
    blank = body[body.index("if ($null -eq $rateValue) {"):body.index("} else {")]
    assert "Set-TableCell" in blank, "the blank branch writes nothing"
    for skip in ("continue", "break", "return"):
        assert skip not in blank, (
            f"the blank branch {skip}s instead of writing the blank cell, which "
            "collapses a null rate into an absent one"
        )
    # And the two loops are genuinely nested, not sequential.
    inner = body.index("foreach ($rateProperty in $rates.PSObject.Properties)")
    outer = body.index("foreach ($profileProperty in $Model.inflation.PSObject.Properties)")
    assert outer < inner < body.rindex("}"), "the rate loop is not inside the profile loop"


def test_185_every_gate_states_why_it_did_not_attempt() -> None:
    """A gate that says only 'not attempted' has thrown its evidence away.

    Run 6's whole value was that the gate named the reason: P5-ALL carried
    'not attempted: Gate-B fixture establishment failed on the golden plan
    case', which is what made one result enough to locate the root.
    """
    source = _executable(SCENARIOS)
    reasons = re.findall(r"'not attempted:? ?([^']*)'", source)
    assert len(reasons) >= 5, f"only {len(reasons)} gate messages found"
    for reason in reasons:
        assert len(reason.strip()) >= 20, (
            f"a gate reports 'not attempted' with no usable reason: {reason!r}"
        )
    # Every P5-ALL emission is one of them.
    gates = re.findall(
        r"Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL' `\s*\n\s*\(?'"
        r"not attempted:", source)
    # Five before Run 7, seven after: P5-CMP added a success gate and a catch
    # gate of its own, exactly as P5-FX and P5-FIX each carry.
    assert len(gates) == 7, f"the P5-ALL gate set changed: {len(gates)}"


# ===========================================================================
# 25. RUNTIME RUN 7: a real whole-project compile gate
# ===========================================================================
# Run 7 passed A1 ("PCCM_AutomationBegin is callable") and P5-M (six API
# procedures callable) and then met a VBE compile error inside the analytical
# path. Callability is not compilation, so the stronger claim gets a scenario of
# its own, before anything that depends on it.
def _compile_gate() -> str:
    """The P5-CMP block: from the P5-P4 gate's return to the FX capture."""
    driver = _executable(SCENARIOS)
    driver = driver[driver.index("function Invoke-Phase5GateBScenarios"):]
    return driver[driver.index("$vbe = $Excel.VBE"):
                  driver.index("Save-Phase5LockedFxSeed")]


def test_186_r19_the_compile_gate_runs_before_anything_that_depends_on_it() -> None:
    """R19. P5-CMP precedes P5-FX, P5-FIX and every fixture."""
    scenarios = _executable(SCENARIOS)
    driver = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    order = {
        name: driver.index(f"Add-Phase5Result '{name}'")
        for name in ("P5-P4", "P5-CMP", "P5-FX", "P5-FIX", "P5-AN")
    }
    assert order["P5-P4"] < order["P5-CMP"] < order["P5-FX"] < order["P5-FIX"] < order["P5-AN"], (
        f"the gate order is wrong: {sorted(order, key=order.get)}"
    )
    # Nothing touches the workbook before the compile gate does.
    before = driver[:driver.index("$vbe = $Excel.VBE")]
    for touch in ("Save-Phase5LockedFxSeed", "Set-Phase5Fixture", "Set-TableCell",
                  "$Excel.Run(", "Get-TableBody"):
        assert touch not in before, (
            f"{touch} runs before the project is proved to compile"
        )
    # It is a declared scenario.
    assert "'P5-CMP'" in _procedure(scenarios, "Get-Phase5ScenarioIds")


def test_187_the_compile_gate_addresses_the_command_by_id_and_proves_it_exists() -> None:
    """The mechanism, and the three ways it could have been fake.

    A caption lookup would find nothing on a non-English Excel and would report
    success for a project that never compiled. A missing control would do the
    same. And executing without reading Enabled afterwards would prove only that
    a menu command was invoked.
    """
    block = _compile_gate()
    # THE LOOKUP, not the prose. "Compile VBAProject" is fine in a check LABEL;
    # what may never appear is a caption reaching FindControl, or any read of
    # .Caption, because both are localised.
    #
    # RUN 8 retired the old one-argument-plus-$null form: the whole argument
    # list is now pinned, so neither a reintroduced $null nor a changed Type
    # can slip through as "still by ID".
    # THREE lookups, all identical: the discovery call, the Run-8 diagnostic
    # probe, and the Run-9 settlement reacquisition. Pinning the whole list
    # keeps a fourth from appearing with looser criteria.
    lookups = re.findall(r"FindControls?\(([^)]*)\)", block)
    assert lookups == ["$msoControlButton, 578, $missing, $missing"] * 3, lookups
    assert ".Caption" not in block, "the gate reads a localised caption"
    assert "FindControl('" not in block and 'FindControl("' not in block
    # `$missing` has to BE Missing. Named arguments would omit Tag and Visible;
    # PowerShell cannot, so the sentinel is the only thing standing between
    # "omitted" and "supplied as a criterion nothing matches".
    assert "$missing = [System.Reflection.Missing]::Value" in block, (
        "the omitted arguments are not omitted with a Missing sentinel"
    )
    # The control must EXIST, and its absence is a failure of this gate.
    assert "'the Compile VBAProject command (ID 578) exists'" in block
    assert "($null -ne $control)" in block
    # It is EXECUTED when there is something to compile, and only once the
    # control has said what it is.
    assert "$control.Execute()" in block
    assert "$before = [bool]$control.Enabled" in block
    assert "if ($before) {" in block
    # AND THE BRANCH IS GUARDED BY THE PROOF, not merely preceded by it. The M6
    # run showed the difference: swapping the condition back to
    # `$null -ne $control` leaves the assignment sitting harmlessly above an
    # Execute it no longer controls, so an ordering assertion sees nothing.
    guard = re.search(
        r"if \(([^)]*)\) \{\n\s+\$before = \[bool\]\$control\.Enabled", block)
    assert guard, "the Enabled read is not inside a guarded branch at all"
    assert guard.group(1).strip() == "$controlProved", (
        f"Execute is guarded by {guard.group(1).strip()}, not by the control-identity proof"
    )
    assert "$controlProved = ($idOk -and $typeOk)" in block
    # ...and the positive evidence is that the command has gone quiet. RUN 9
    # retired the immediate read on the SAME handle: Enabled is cached UI state,
    # and one statement after Execute it measures the harness's timing rather
    # than the compiler's outcome. The evidence is now a REACQUIRED control that
    # reports the command disabled.
    assert "$after = [bool]$control.Enabled" not in block, (
        "the compiled state is still read from the cached post-Execute handle"
    )
    assert "$lastEnabled = [bool]$poll.Enabled" in block, (
        "the gate never reads Enabled from a reacquired control"
    )
    assert "if (-not $lastEnabled) { $settled = $true }" in block
    assert "$settled `" in block, (
        "the compiled-state check is not decided by the settlement observation"
    )
    # VBProject access is COM: every transient is released, including the two
    # project handles the target-project identity gate opens. Each release is
    # matched as a whole statement - variable AND label - so a handle released
    # under the wrong label, or a label with no release, is caught.
    flat = re.sub(r"[ \t]+", " ", block)   # the releases are column-aligned
    for variable, label in (("$control", "CommandBarControl"),
                            ("$bars", "CommandBars"),
                            ("$activeProject", "VBProject(active)"),
                            ("$targetProject", "VBProject(target)"),
                            ("$vbe", "VBE")):
        assert f"Release-Transient {variable} '{label}'" in flat, (
            f"{variable} is not released as '{label}'"
        )
    # Eight in total: the five above, the Run-8 diagnostic collection, the
    # stale control dropped straight after Execute, and the handle each Run-9
    # settlement observation reacquires. The last three are released where they
    # are opened, inside the branches that open them.
    assert block.count("Release-Transient") == 8, block.count("Release-Transient")
    assert "Release-Transient $control 'CommandBarControl'" in flat
    assert "Release-Transient $poll 'CommandBarControl(settle)'" in flat, (
        "the reacquired settlement control is never released"
    )
    assert "Release-Transient $probeControls 'CommandBarControls(probe)'" in flat, (
        "the diagnostic FindControls collection is never released"
    )
    # And the five accepted handles are all in the same, last finally, so no
    # path through the gate can skip one.
    finally_at = block.rindex("} finally {")
    assert block.count("Release-Transient", finally_at) == 5, (
        "a release happens outside the finally that guarantees it"
    )


def test_188_the_compile_gate_fails_closed_and_destroys_no_evidence() -> None:
    """It may not pass by accident, and it may not tidy away a dialog."""
    scenarios = _executable(SCENARIOS)
    block = _compile_gate()
    # NO SUPPRESSION. An unexpected compile-error dialog is diagnostic evidence.
    for suppressor in ("-ErrorAction SilentlyContinue", "-ErrorAction Ignore",
                       "DisplayAlerts = $true", "DisplayAlerts=$true",
                       "SendKeys", "$ErrorActionPreference"):
        assert suppressor not in block, f"the compile gate {suppressor}s"
    # NO SUCCESS FROM IMPORTING OR FROM ONE MACRO. The gate's own checks are the
    # control's existence and the control's state, nothing else.
    for false_proof in ("VBComponents.Import", "PCCM_AutomationBegin", "PCCM_Calculate"):
        assert false_proof not in block, (
            f"the compile gate claims success from {false_proof}"
        )
    # EXACTLY ONE RESULT, on both paths.
    whole = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    section = whole[:whole.index("Save-Phase5LockedFxSeed")]
    assert section.count("Add-Phase5Result 'P5-CMP'") == 2, (
        "P5-CMP must emit on exactly its success path and its catch path"
    )
    assert "Add-Phase5Result 'P5-CMP' 'Whole VBA project compile gate' 'FAIL' (Format-Phase5Err $_)" \
        in section, "a throw inside the gate is not reported as a FAIL"
    # A FAIL, never a SKIP - checked over the gate's OWN block, because the
    # P5-P4 gate above it legitimately counts SKIPs in the Phase-4 matrix.
    assert "'SKIP'" not in block, "the compile gate can be skipped"
    gate_tail = section[section.index("$vbe = $Excel.VBE"):]
    assert "'SKIP'" not in gate_tail, "the compile gate reports a SKIP"


def test_189_r20_a_compile_gate_failure_leaves_the_lifecycle_evidence_reachable() -> None:
    """R20. Y, Z, P5-LDG and P5-FIN still run after the gate fails."""
    scenarios = _executable(SCENARIOS)
    whole = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    section = whole[:whole.index("Save-Phase5LockedFxSeed")]
    # It RETURNS from the scenario driver; it does not exit the process.
    assert section.count("return") >= 2, "the compile gate does not stop the run below it"
    for fatal in ("exit 1", "exit(", "throw", "[Environment]::Exit"):
        assert fatal not in section.split("$vbe = $Excel.VBE")[1], (
            f"the compile gate {fatal}s, which would take the shutdown with it"
        )
    assert "Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL'" in section
    # And the caller's lifecycle path is downstream of the whole scenario call.
    harness = _executable(HARNESS)
    order = [harness.index(token) for token in (
        "Invoke-Phase5GateBScenarios -Excel",
        "Add-Result 'Z' 'Excel closed naturally",
        "$transient = @(Get-TransientFailures)",
        "Add-Phase4FinalCompletenessResult -Results $results",
        "Add-Phase5LedgerIntegrityResult")]
    assert order == sorted(order)


def test_190_a1_and_p5_m_claim_only_what_they_observe() -> None:
    """R17 and R18, from the harness side.

    Run 7 is the counterexample that retires the inference, and it is recorded
    where the claim used to be so a future round cannot quietly restore it.
    """
    harness = _text(HARNESS)
    assert "'PCCM_AutomationBegin is callable' $true" in harness
    retired = "'PCCM_AutomationBegin is callable (the VBA project compiles)'"  # retired-authority
    assert retired not in harness
    assert "RUN-7 CORRECTION" in harness, (
        "the retirement of the compile claim is not recorded where it was made"
    )
    scenarios = _text(SCENARIOS)
    assert "whole-project claim belongs to P5-CMP alone." in scenarios, (
        "P5-M does not record that callability is not compilation"
    )
    assert "runtime execution is deferred to P5-FIX" in scenarios, (
        "P5-M does not record that PCCM_Calculate's execution is deferred"
    )
    # AND THE RUN-7 HISTORY LINE DISTINGUISHES REPORT FROM PROOF. A per-line
    # claim, deliberately not the file-wide phrase sweep test_201 runs: the line
    # that remembers Run 7 may not assert callability as fact, because one of
    # the six claims it is remembering was borrowed evidence.
    history_line = next(line for line in scenarios.splitlines()
                        if "#   P5-M   PASS" in line)
    assert "reported" in history_line, (
        "the P5-M history line states callability as fact rather than as what "
        f"the then-current harness reported: {history_line.strip()!r}"
    )
    assert "procedures callable" not in history_line, history_line.strip()
    assert scenarios.count("#   P5-M   PASS") == 1, (
        "the Run-7 P5-M history is told more than once"
    )
    # No check label anywhere in either file claims compilation except P5-CMP's.
    for text in (_executable(HARNESS), _executable(SCENARIOS)):
        for label in _check_labels(text):
            if "compil" in label.lower():
                assert label in P5_CMP_COMPILE_LABELS, label


# ===========================================================================
# 26. REVIEW OF ae52bdd: one compile authority, and no borrowed evidence
# ===========================================================================
# Two blockers, both about evidence AUTHORITY rather than behaviour.
#
#   1. The Run-7 round retired "A1 proves the project compiles" from A1's own
#      check label and stopped there. The same claim was still standing in the
#      harness overview, the diagnostic module header, the scenario commentary,
#      the P5-D0 result title, two docs and this suite's own docstrings.
#   2. P5-M emitted "the API procedure PCCM_Calculate is callable" as a PASS
#      without ever invoking it.
RETIRED_AUTHORITY_PHRASES = (
    "A1 has proved the production project compiles",  # retired-authority
    "A1 has proved the production VBA project compiles",  # retired-authority
    "AFTER the A1 production compile",  # retired-authority
    "A1 production compile",  # retired-authority
    "A1 IS the first real VBA compilation boundary",  # retired-authority
    "A1 is the first real VBA compilation boundary",  # retired-authority
    "A1 remains the first real VBA compilation boundary",  # retired-authority
    "A1 first Application.Run of the run  ->  the PRODUCTION project compiles",  # retired-authority
    "the VBA project compiles)",  # retired-authority
)

# A line may quote a retired phrase for exactly one reason: it is the test or
# the note that FORBIDS it. Those lines carry this marker, in the same spirit as
# the `# refusal-list` marker the COM-lifecycle sweep already uses.
AUTHORITY_EXEMPTION_MARKER = "retired-authority"

# The ONLY check labels in either PowerShell file that may mention compiling.
# They all belong to P5-CMP, which is the single whole-project compile
# authority; anything else claiming a compile is the defect ae52bdd removed.
# Kept in one place so the census tests cannot drift from the evidence chain.
P5_CMP_COMPILE_LABELS = (
    "Compile VBAProject was executed at most once",
    "the Compile VBAProject command (ID 578) exists",
    "the compiled state was read by reacquiring the exact Id-578 control",
    "the target PCCM VBProject reached the VBE compiled state",
)


def _authority_scan_files() -> list[Path]:
    """Every .ps1, .bas, .py and .md the correction has to cover."""
    roots = (PCCM_ROOT / "bootstrap", PCCM_ROOT / "docs", PCCM_ROOT / "tests",
             PCCM_ROOT / "src", PCCM_ROOT / "builder", PCCM_ROOT / "spec")
    found: list[Path] = []
    for root in roots:
        if not root.is_dir():
            continue
        for pattern in ("**/*.ps1", "**/*.bas", "**/*.py", "**/*.md"):
            found.extend(path for path in root.glob(pattern)
                         if "__pycache__" not in path.parts)
    found.append(PCCM_ROOT / "README.md")
    return [path for path in found if path.is_file()]



def _check_labels(text: str) -> list[str]:
    """Every Add-Check label, including those on a backtick continuation line.

    `Add-Check $list `\n    'the label'` is the harness's normal shape for a
    long label, and a regex anchored to `$list` sees none of them. That is how a
    label scan can report zero offenders while the offending label is right
    there - so the continuations are joined first.
    """
    joined = re.sub(r"`\s*\n\s*", " ", text)
    # `\s+`, not a single space: joining a continuation leaves the space that was
    # before the backtick as well as the one it became.
    return re.findall(r"Add-Check \$list\s+\(?'([^']*)'", joined)


def test_191_r1_r2_r3_no_file_claims_a1_proved_the_project_compiles() -> None:
    """R1, R2, R3. The whole subtree, not just executable check labels.

    Run 7 is the counterexample: A1 PASS, P5-M PASS, then a VBE compile error in
    the analytical path. Any surviving statement of the retired inference is a
    false authority wherever it sits - a comment, a result title, a docstring,
    an assertion message or documentation prose.
    """
    files = _authority_scan_files()
    assert len(files) >= 25, f"the sweep only found {len(files)} files"
    assert any(path.suffix == ".bas" for path in files)
    assert any(path.suffix == ".ps1" for path in files)
    assert any(path.suffix == ".md" for path in files)

    offenders: list[str] = []
    exempted = 0
    for path in files:
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            for phrase in RETIRED_AUTHORITY_PHRASES:
                if phrase not in line:
                    continue
                if AUTHORITY_EXEMPTION_MARKER in line:
                    exempted += 1
                    continue
                offenders.append(
                    f"{path.relative_to(PCCM_ROOT)}:{number}: {phrase!r}")
    assert not offenders, (
        "the retired A1 whole-project compile authority is still asserted:\n  "
        + "\n  ".join(offenders)
    )
    # The exemption is real and narrow: only the sites that FORBID the phrases
    # may contain them, and there must be some, or this sweep proves nothing.
    assert exempted >= 3, (
        f"only {exempted} marked exemption(s); the sweep may be matching nothing"
    )


def test_192_r4_r5_only_p5_cmp_owns_the_whole_project_compile_claim() -> None:
    """R4 and R5. P5-D0 follows P5-CMP, and one scenario owns the claim."""
    scenarios = _text(SCENARIOS)
    harness = _text(HARNESS)
    diagnostics = DIAGNOSTIC.read_text(encoding="utf-8")

    # THE DOCUMENTATION STATES THE CORRECTED HIERARCHY POSITIVELY, not merely
    # by the absence of the retired one. A doc that deleted the sentence would
    # pass a sweep for forbidden phrases and still leave the reader guessing.
    doc = (PCCM_ROOT / "docs" / "phase5_gate_b_harness.md").read_text(encoding="utf-8")
    assert "A1 is the first `Application.Run` boundary, not a compilation boundary." in doc, (
        "the harness doc does not state what A1 actually proves"
    )
    assert "P5-CMP  the WHOLE production project compiles" in doc, (
        "the lifecycle diagram does not name the compile gate"
    )

    # R4. The module header, the harness overview and the result title all name
    # P5-CMP, and none of them names A1.
    assert "only AFTER scenario P5-CMP has proved the production VBA project compiles" \
        in diagnostics
    assert "imported only AFTER P5-CMP has\n          proved the production project compiles" \
        in harness
    assert ("Add-Phase5Result 'P5-D0' 'Transient diagnostic module imported AFTER "
            "the P5-CMP whole-project compile'") in scenarios

    # R5. Exactly one scenario claims whole-project compilation, and its checks
    # are the only ones that mention compiling.
    executable_scenarios = _executable(SCENARIOS)
    labels = _check_labels(executable_scenarios) + _check_labels(_executable(HARNESS))
    # Not every label is a literal - many are built by concatenation from a case
    # id - so this is a floor on the LITERAL labels, not a total.
    assert len(labels) > 140, f"the label scan found only {len(labels)} labels"
    compile_labels = sorted({label for label in labels if "compil" in label.lower()})
    assert compile_labels == sorted(P5_CMP_COMPILE_LABELS), compile_labels
    # And the result TITLES. Any title that mentions compilation must either be
    # P5-CMP's own, or must name P5-CMP as the authority it defers to. A title
    # that mentions a compile and names A1 is the defect this round removed.
    joined = re.sub(r"`\s*\n\s*", " ", executable_scenarios)
    pairs = re.findall(r"Add-Phase5Result '([\w-]+)' \(?'([^']*)'", joined)
    mentions = [(sid, title) for sid, title in pairs if "compil" in title.lower()]
    assert mentions, "no result title mentions compilation at all"
    for sid, title in mentions:
        if sid == "P5-CMP":
            continue
        assert "P5-CMP" in title, (
            f"{sid} mentions a compile without naming P5-CMP as the authority: {title!r}"
        )
        assert "A1" not in title, f"{sid} still attributes a compile to A1: {title!r}"
    assert {sid for sid, _ in mentions} == {"P5-CMP", "P5-D0"}, sorted(
        {sid for sid, _ in mentions})


def test_193_r12_r13_the_compile_gate_still_precedes_everything_it_gates() -> None:
    """R12 and R13, restated now that P5-D0 depends on it too."""
    scenarios = _executable(SCENARIOS)
    driver = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    order = {name: driver.index(f"Add-Phase5Result '{name}'")
             for name in ("P5-P4", "P5-CMP", "P5-FX", "P5-M", "P5-D0", "P5-FIX", "P5-AN")}
    assert order["P5-P4"] < order["P5-CMP"], "the compile gate runs before the Phase-4 gate"
    for dependent in ("P5-FX", "P5-M", "P5-D0", "P5-FIX", "P5-AN"):
        assert order["P5-CMP"] < order[dependent], (
            f"{dependent} runs before the whole-project compile gate"
        )
    # The diagnostic IMPORT itself, not just its result, follows the gate.
    assert driver.index("$components.Import($source)") > order["P5-CMP"], (
        "a test module is imported before the production project is proved to compile"
    )
    # R13. The lifecycle evidence is still downstream of the whole scenario call.
    harness = _executable(HARNESS)
    positions = [harness.index(token) for token in (
        "Invoke-Phase5GateBScenarios -Excel",
        "Add-Result 'Z' 'Excel closed naturally",
        "$transient = @(Get-TransientFailures)",
        "Add-Phase4FinalCompletenessResult -Results $results",
        "Add-Phase5LedgerIntegrityResult")]
    assert positions == sorted(positions)


def _p5m_block() -> str:
    source = _executable(SCENARIOS)
    return source[source.index("Add-Phase5Result 'P5-FX'"):
                  source.index("Add-Phase5Result 'P5-M'")]


def test_194_r6_r7_pccm_calculate_is_never_called_callable_without_being_called() -> None:
    """R6 and R7. The overclaim, removed at its source.

    P5-M used to set `$callable = $true` inside a branch whose whole purpose was
    to NOT invoke PCCM_Calculate, then emit the callability label anyway.
    """
    block = _p5m_block()
    # THE BRANCH THAT LIED IS GONE.
    assert "$detail = 'exercised by the analytical scenarios below'" not in block, (
        "P5-M still counts a future exercise as present evidence"
    )
    # No assignment of the callability flag survives inside a PCCM_Calculate branch.
    calculate_branch = re.search(
        r"if \(\$name -eq 'PCCM_Calculate'\) \{(.*?)\n            \}", block, re.S)
    assert calculate_branch, "the PCCM_Calculate branch is gone entirely"
    body = calculate_branch.group(1)
    assert "$callable = $true" not in body, (
        "PCCM_Calculate is still marked callable without being called"
    )
    assert "Add-Check" not in body, (
        "the PCCM_Calculate branch still emits a check of its own"
    )
    assert "continue" in body, "the branch does not skip the callability check"
    assert "deferred to P5-FIX" in body, (
        "the branch does not say where the runtime evidence actually comes from"
    )

    # R7. The callability label exists only where an Excel.Run precedes it.
    label = "('the API procedure ' + $name + ' is callable')"
    assert block.count(label) == 1, block.count(label)
    before = block[:block.index(label)]
    run_at = before.rindex("$probe = $Excel.Run($name)")
    assert run_at > before.rindex("continue"), (
        "the callability check is not downstream of a real Application.Run"
    )
    # And PCCM_Calculate is not invoked anywhere in P5-M.
    assert "$Excel.Run('PCCM_Calculate')" not in block, (
        "P5-M drives the stateful calculation just to satisfy a label"
    )

    # NO CHECK IN P5-M IS VACUOUSLY TRUE. `$callable = $true` in a branch that
    # never called anything was one shape of that; a condition written as a bare
    # ($true) is the other, and it is how a real check gets hollowed out while
    # its label survives. Every condition here must read something.
    for vacuous in ("($true)", "($true) `", "-Value $true"):
        assert vacuous not in block, (
            f"P5-M emits a check whose condition is the constant {vacuous}"
        )
    # Every one of the six is still asked a real question about the project.
    assert block.count("($declared -contains $name)") == 1
    assert "Get-Phase5ProjectProcedureNames" in block


def test_195_r8_r9_declaration_for_six_callability_for_five() -> None:
    """R8 and R9. Each of the six gets the evidence it actually has."""
    block = _p5m_block()
    # R8. All six are proved DECLARED, out of the persisted project's own code.
    assert "Get-Phase5ProjectProcedureNames -Workbook $Workbook" in block
    assert "($declared -contains $name)" in block
    assert "is declared in the persisted project'" in block
    # The declaration reader is not a manifest re-read: it goes to CodeModule
    # text and strips comments and literals first, like P5-EV does.
    reader = _procedure(_executable(SCENARIOS), "Get-Phase5ProjectProcedureNames")
    assert "$component.CodeModule" in reader
    assert "Get-VbaExecutableCode -Code $raw" in reader, (
        "a procedure named in a comment or a string would count as declared"
    )
    assert "$Manifest" not in reader, "the reader reads the manifest, not the project"
    for handle in ("'CodeModule'", "'VBComponent'", "'VBComponents'", "'VBProject'"):
        assert handle in reader, f"{handle} is not released"

    # R9. The five that ARE claimed callable really cross Application.Run.
    assert "$probe = $Excel.Run($name)" in block
    assert "$callable = $true" in block
    flag = block.index("$callable = $true")
    assert block.index("$probe = $Excel.Run($name)") < flag, (
        "the callability flag is set before the call that justifies it"
    )
    # The manifest still projects six, and the split is five plus one.
    assert "'the manifest projects exactly six API procedures'" in block
    assert "($api.Count -eq 6)" in block
    api = {name for name in _emitted()["manifest"]["vba"]["api_procedures"]}
    assert len(api) == 6 and "PCCM_Calculate" in api, sorted(api)


def test_196_r10_r11_p5_fix_is_the_first_valid_fixture_calculate() -> None:
    """R10 and R11. The commentary matches the evidence, and P5-FIX owns first run."""
    scenarios_raw = _text(SCENARIOS)
    # R10. The retired claim about all six crossing COM is gone.
    assert "the call crosses the COM boundary" not in scenarios_raw, (
        "P5-M still claims all six API procedures cross the COM boundary"
    )
    assert "DECLARATION   the name exists in the persisted VBA project" in scenarios_raw
    assert "CALLABILITY   Application.Run reached it and it answered" in scenarios_raw
    assert "EXECUTION     it ran against a valid fixture and did its work" in scenarios_raw

    # R11. The first PCCM_Calculate of the run is P5-FIX's, on a real fixture.
    source = _executable(SCENARIOS)
    driver = source[source.index("function Invoke-Phase5GateBScenarios"):]
    calls = [m.start() for m in re.finditer(r"\$Excel\.Run\('PCCM_Calculate'\)", driver)]
    assert calls, "PCCM_Calculate is never executed at all"
    fix_at = driver.index("Add-Phase5Result 'P5-FIX'")
    an_at = driver.index("Add-Phase5Result 'P5-AN'")
    assert calls[0] < fix_at, "the first PCCM_Calculate is not inside P5-FIX"
    assert calls[1] < an_at, "P5-AN does not drive the calculation after P5-FIX"
    # And it is preceded by a real fixture in the same block.
    block = driver[driver.index("Add-Phase5Result 'P5-D8'"):fix_at]
    assert block.index("Set-Phase5Fixture -Excel $Excel") < block.index(
        "$Excel.Run('PCCM_Calculate')"), (
        "P5-FIX calculates before it establishes a fixture"
    )
    # Nothing earlier in the run touches it.
    assert driver.index("Add-Phase5Result 'P5-M'") < calls[0]
    assert driver.index("Add-Phase5Result 'P5-CMP'") < calls[0]


# ===========================================================================
# 27. REVIEW OF d21e1d7: which VBProject did command 578 compile?
# ===========================================================================
# Command 578 is a VBE command and it acts on the VBE's ACTIVE project. Reading
# Enabled and calling Execute without binding that to the workbook under test
# proves "some active project compiled" - and a fresh owned Excel instance can
# still carry an add-in, a startup workbook or PERSONAL.XLSB, each with its own
# VBProject. Gate B may not assume the right one is active.
def test_197_r1_r2_r3_the_gate_binds_the_command_to_the_workbook_project() -> None:
    """R1, R2, R3. Both projects are read, and compared, BEFORE the command."""
    block = _compile_gate()
    # R1 and R2: both handles are obtained.
    assert "$targetProject = $Workbook.VBProject" in block, (
        "the gate never reads the Stage-B workbook's own VBProject"
    )
    assert "$activeProject = $vbe.ActiveVBProject" in block, (
        "the gate never reads the project the VBE command will actually act on"
    )
    # Each read is checked, so an unreadable project fails rather than passing
    # through as $null.
    assert "'the Stage-B workbook exposes its VBProject'" in block
    assert "'the VBE reports an active VBProject'" in block

    # R3: the identity comparison precedes the command entirely - not just the
    # Execute, but FindControl and the Enabled read too.
    identity_at = block.index("$targetIsActive =")
    for later in ("FindControl", "$control.Enabled", "$control.Execute()"):
        assert identity_at < block.index(later), (
            f"{later} is reached before the target-project identity is established"
        )
    assert "'the active VBE project IS the Stage-B workbook project (by file path)'" in block


def test_198_r4_identity_is_the_file_path_not_the_project_name() -> None:
    """R4. `VBAProject` is the default name every project gets."""
    block = _compile_gate()
    assert "$targetFile = [string]$targetProject.FileName" in block
    assert "$activeFile = [string]$activeProject.FileName" in block
    # Normalised for Windows filesystem equivalence, and ONLY for that.
    assert "[System.IO.Path]::GetFullPath($targetFile)" in block
    assert "[System.IO.Path]::GetFullPath($activeFile)" in block
    assert "[System.StringComparison]::OrdinalIgnoreCase" in block, (
        "the path comparison is case-sensitive; NTFS paths are not"
    )
    # THE OPERANDS ARE THE PATHS. Checking that GetFullPath appears somewhere
    # nearby is not enough: the comparison itself could still be handed the
    # names while the path lines sit unused above it.
    comparison = re.search(
        r"\$sameFile = \[string\]::Equals\(([^,]+), ([^,]+),", block)
    assert comparison, "the identity comparison is not a [string]::Equals"
    assert comparison.group(1).strip() == "$targetFull", comparison.group(1)
    assert comparison.group(2).strip() == "$activeFull", comparison.group(2)
    # And $targetFull / $activeFull are the NORMALISED FILE PATHS, nothing else.
    assert "$targetFull = [System.IO.Path]::GetFullPath($targetFile)" in block
    assert "$activeFull = [System.IO.Path]::GetFullPath($activeFile)" in block
    # THE NAMES ARE CONTEXT, NEVER THE TEST.
    assert "$targetIsActive = $haveFiles -and $sameFile" in block, (
        "the identity decision is not the path comparison"
    )
    assert "$targetIsActive = $true" not in block, (
        "the identity decision is hard-coded true"
    )
    # An unsaved project has no FileName, and two empty strings are not an
    # identity: both sides must actually name a file.
    assert "$haveFiles = (-not [string]::IsNullOrWhiteSpace($targetFile)) -and `" in block
    assert "'both VBProjects name a file, so identity is comparable at all'" in block
    # No caption, no substring, no display text anywhere in the decision.
    for forbidden in (".Caption", "-like", "-match", "Contains("):
        assert forbidden not in block, f"the identity gate uses {forbidden}"


def test_199_r5_r6_r7_a_mismatched_active_project_is_a_fail_and_no_compile() -> None:
    """R5, R6, R7. Wrong project: no Execute, no PASS, and a reason."""
    block = _compile_gate()
    # R6: Execute is inside the branch the identity gate guards.
    guard = block.index("if (-not $targetIsActive) {")
    else_at = block.index("} else {", guard)
    mismatch = block[guard:else_at]
    matched = block[else_at:]
    assert "$control.Execute()" in matched, "the compile never happens on the good path"
    assert "$control.Execute()" not in mismatch, (
        "the gate compiles a project it has not identified"
    )
    assert "FindControl" not in mismatch, (
        "the gate even looks the command up before knowing whose project it is"
    )
    # THE ACTIVE PROJECT IS READ FROM THE VBE, NOT ALIASED TO THE TARGET.
    # Aliasing would make the comparison compare a thing with itself, which
    # passes always and proves nothing.
    assert "$activeProject = $vbe.ActiveVBProject" in block
    for alias in ("$activeProject = $targetProject", "$targetProject = $activeProject"):
        assert alias not in block, f"the two projects are aliased ({alias})"
    # And the decision cannot be short-circuited to a constant.
    for constant in ("$targetIsActive = $true", "$targetIsActive = $True",
                     "$sameFile = $true", "$haveFiles = $true"):
        assert constant not in block, f"the identity gate is hard-coded ({constant})"
    # NOR DECIDED ON THE NAMES. `VBAProject` is the default name every project
    # gets, so a name comparison would report identity between two unrelated
    # projects. The names may be REPORTED; they may not be compared.
    for named in ("[string]::Equals($targetName", "[string]::Equals($activeName",
                  "$targetName -eq $activeName", "$activeName -eq $targetName",
                  "$sameFile = [string]::Equals($targetName"):
        assert named not in block, (
            f"the identity gate decides on the project NAME ({named})"
        )
    decision_line = next(line for line in block.splitlines()
                         if "$sameFile = [string]::Equals" in line)
    assert "Name" not in decision_line, decision_line

    # R5: the mismatch cannot be a PASS. `$targetIsActive` is itself an
    # Add-Check, so a false one fails the checklist that decides the result.
    assert "$null = Add-Check $list `\n                    'the active VBE project IS the Stage-B workbook project (by file path)' `\n                    $targetIsActive" in _text(SCENARIOS), (
        "the identity is not itself a recorded check, so a mismatch could pass"
    )
    assert "$compileOk = Test-ChecklistOk $list" in _executable(SCENARIOS)
    # And the mismatch says WHY, rather than failing silently.
    assert "is NOT the PCCM workbook" in mismatch
    assert "was NOT executed" in mismatch

    # R7: a DISABLED command is not evidence either, unless the identity held -
    # the Enabled read is on the same guarded branch as the Execute.
    assert "$before = [bool]$control.Enabled" in matched
    assert "$before = [bool]$control.Enabled" not in mismatch
    # ...and so is the settlement poll that replaced the immediate read.
    assert "$lastEnabled = [bool]$poll.Enabled" in matched
    assert "$lastEnabled = [bool]$poll.Enabled" not in mismatch


def test_200_r8_every_new_com_reference_is_released_on_every_path() -> None:
    """R8. A VBProject RCW may not outlive P5-CMP."""
    block = _compile_gate()
    flat = re.sub(r"[ \t]+", " ", block)
    for variable, label in (("$targetProject", "VBProject(target)"),
                            ("$activeProject", "VBProject(active)")):
        assert f"Release-Transient {variable} '{label}'" in flat, variable
        assert f"{variable} = $null }}" in flat, f"{variable} is not cleared after release"
    # The five accepted handles are all in the last finally. The sixth release
    # is the Run-8 diagnostic collection, and it has a finally of its own.
    finally_at = block.rindex("} finally {")
    assert block.count("Release-Transient") == 8
    assert block.count("Release-Transient", finally_at) == 5
    probe = block[block.index("$probeControls = $null"):finally_at]
    assert "} finally {" in probe, "the diagnostic probe has no finally"
    assert "Release-Transient $probeControls" in re.sub(r"[ \t]+", " ", probe)
    # The handles are declared before the try, so the finally can always see them.
    scenarios = _executable(SCENARIOS)
    declaration = "$targetProject = $null; $activeProject = $null"
    assert declaration in scenarios
    assert scenarios.index(declaration) < scenarios.index("$targetProject = $Workbook.VBProject")


def test_201_r12_the_p5m_history_is_stated_precisely() -> None:
    """R12. Run 7 reported six PASS lines; one of them was borrowed evidence.

    The historical fact that P5-M printed PASS may stay. The conclusion drawn
    from it may not, and "six API procedures callable" states the conclusion.  # retired-authority
    """
    # ASSEMBLED, not written out. A phrase list that bans a phrase and then
    # spells it would find itself - the codebase's established idiom for a
    # self-matching sweep.
    six = "six "  # retired-authority
    stale = (six + "API procedures callable", six + "APIs callable",  # retired-authority
             six + "callable APIs", "all " + six + "are callable",  # retired-authority
             "all " + six + "api procedures are exercised")  # retired-authority
    # ALSO THE UNDERSCORED FORM, so a test or function NAME making the claim is
    # swept by the same rule as prose.
    stale = stale + tuple(phrase.replace(" ", "_") for phrase in stale)
    # BOTH SIDES LOWERCASED. The first version of this sweep compared a phrase
    # containing "API" against `line.lower()`, so it matched nothing at all and
    # reported a clean file while the claim sat in it. That is the same shape as
    # the defect the sweep exists to catch, one level up.
    offenders: list[str] = []
    for path in _authority_scan_files():
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            lowered = line.lower()
            for phrase in stale:
                if phrase.lower() in lowered and AUTHORITY_EXEMPTION_MARKER not in line:
                    offenders.append(f"{path.relative_to(PCCM_ROOT)}:{number}: {phrase!r}")
    assert not offenders, (
        "an unqualified six-callable claim survives:\n  " + "\n  ".join(offenders))
    # AND THE HISTORY BLOCK ITSELF, positively. A sweep for banned phrases
    # passes on a file that simply deleted the history; this requires the
    # correction to be stated where the history is told.
    scenarios = _text(SCENARIOS)
    history = scenarios[scenarios.index("#   P5-M   PASS"):]
    history = history[:900]
    for required in ("under the evidence model", "had never crossed",
                     "six DECLARED and five", "PCCM_Calculate"):
        assert required in history, (
            f"the Run-7 P5-M history does not say {required!r}"
        )

    # And where the history IS recorded, it is recorded precisely.
    assert "reported" in scenarios[scenarios.index("P5-M   PASS"):
                                   scenarios.index("P5-M   PASS") + 400], (
        "the Run-7 history does not distinguish what P5-M reported from what was proved"
    )
    assert "had never crossed" in scenarios, (
        "the borrowed claim is not identified where the history is told"
    )


def test_202_r13_r14_the_evidence_hierarchy_is_what_the_tests_assert() -> None:
    """R13 and R14. No test claims an exercise it has not looked for."""
    own = Path(__file__).read_text()
    retired = "def " + "test_14_all_six_api_procedures" + "_are_exercised"
    assert retired not in own, (
        "the test that claimed six exercises from name presence is back"
    )
    assert "def test_14_each_api_procedure_has_the_evidence_the_hierarchy_gives_it" in own
    # No test NAME anywhere in this suite claims an exercise of all six.
    for name in re.findall(r"^def (test_\w+)", own, re.M):
        assert not ("six" in name and "exercis" in name), name
    # The hierarchy itself, restated once here so a future edit to P5-M has to
    # come past this too.
    source = _executable(SCENARIOS)
    p5m = source[source.index("Add-Phase5Result 'P5-FX'"):
                 source.index("Add-Phase5Result 'P5-M'")]
    assert "is declared in the persisted project'" in p5m       # six
    assert p5m.count("('the API procedure ' + $name + ' is callable')") == 1   # five
    driver = source[source.index("function Invoke-Phase5GateBScenarios"):]
    first_calculate = driver.index("$Excel.Run('PCCM_Calculate')")
    assert driver.index("Add-Phase5Result 'P5-M'") < first_calculate, (
        "PCCM_Calculate runs before P5-M, so P5-M could have claimed it"
    )
    assert first_calculate < driver.index("Add-Phase5Result 'P5-FIX'")


def test_203_r15_the_compile_gate_failure_still_reaches_the_lifecycle() -> None:
    """R15, restated after the identity gate was added."""
    scenarios = _executable(SCENARIOS)
    whole = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    section = whole[:whole.index("Save-Phase5LockedFxSeed")]
    gate_tail = section[section.index("$vbe = $Excel.VBE"):]
    for fatal in ("exit 1", "exit(", "throw", "[Environment]::Exit"):
        assert fatal not in gate_tail, f"the compile gate {fatal}s"
    assert "'SKIP'" not in gate_tail
    assert section.count("Add-Phase5Result 'P5-CMP'") == 2
    assert "Add-Phase5Result 'P5-ALL' 'Phase-5 Gate-B scenarios' 'FAIL'" in section
    harness = _executable(HARNESS)
    order = [harness.index(token) for token in (
        "Invoke-Phase5GateBScenarios -Excel",
        "Add-Result 'Z' 'Excel closed naturally",
        "$transient = @(Get-TransientFailures)",
        "Add-Phase4FinalCompletenessResult -Results $results",
        "Add-Phase5LedgerIntegrityResult")]
    assert order == sorted(order)


# =====================================================================
# RUNTIME RUN 8. THE COMPILER WAS NEVER INVOKED.
# =====================================================================
# Run 8's P5-CMP walked the whole target-project identity chain and passed
# every link of it against real Windows - the active VBE project WAS the PCCM
# Stage-B project, by full path - and then failed on the next line:
#
#     FAIL   the Compile VBAProject command (ID 578) exists
#
# So no compiler ran, and Run 8 licenses no verdict about whether the
# production project compiles. What it proves is about this harness: the
# FindControl call, as written, returned nothing.
#
# CommandBars.FindControl(Type, Id, Tag, Visible) has four optional arguments.
# VBA omits three of them by name; PowerShell cannot, and $null in argument one
# is a supplied criterion, not an omission. The tests below pin the corrected
# call, the identity the returned control has to prove about itself, and the
# wording that may no longer read a discovery failure as a compile defect.


def test_204_r1_r2_r3_r4_the_compile_control_lookup_is_explicit() -> None:
    """R1, R2, R3, R4. Type, Id, and two genuinely omitted arguments."""
    block = _compile_gate()

    # R1. The Run-8 form is gone, in every spelling of it. `$null` in the Type
    # position is a criterion that matches no control, not an omission.
    for retired in ("FindControl($null,", "FindControl($null ,", "FindControls($null,"):
        assert retired not in block, f"the retired lookup {retired} is back"
    assert "$null, 578" not in block, "a $null is still being passed positionally"

    # R2. Type and Id are both explicit, and they are the ONLY argument shape
    # used - the whole list is pinned, not just its first element.
    lookups = re.findall(r"FindControls?\(([^)]*)\)", block)
    assert lookups, "the gate no longer looks the command up at all"
    for arguments in lookups:
        assert arguments == "$msoControlButton, 578, $missing, $missing", arguments

    # R3. 1 is named, not written as a bare literal at the call site. A reader
    # and a mutation both have to go through the name.
    #
    # THE VALUE IS MATCHED WHOLE. `"$msoControlButton = 1" in block` is true of
    # `= 10` as well, and the M2 run proved it: a substring test let the
    # constant change to another MsoControlType with only one detector left
    # standing. The digits are bounded now.
    named = re.search(r"\$msoControlButton = (\d+)\b", block)
    assert named, "the source never says what msoControlButton is"
    assert named.group(1) == "1", (
        f"msoControlButton is {named.group(1)}; the Compile command is a button (1)"
    )
    assert "msoControlButton" in _text(SCENARIOS), "the constant is not named in the source"
    # And it is named ONCE, so the call site and the Type assertion cannot drift.
    assert block.count("$msoControlButton = ") == 1, block.count("$msoControlButton = ")

    # R4. The sentinel is a real Missing, not $null and not an empty string.
    assert "$missing = [System.Reflection.Missing]::Value" in block, (
        "Tag and Visible are not omitted with a Missing sentinel"
    )
    for fake in ("$missing = $null", "$missing = ''", '$missing = ""',
                 "$missing = 0", "[System.DBNull]"):
        assert fake not in block, f"the omission sentinel is faked as {fake}"

    # R5. No caption anywhere - not as the lookup, not as an acceptance
    # predicate, not as a fallback. Captions are localised.
    for caption in (".Caption", "'Compile VBAProject'", '"Compile VBAProject"'):
        assert caption not in block, f"the gate uses {caption} as a lookup or predicate"
    assert "578" in block, "the stable command ID was dropped"

    # AND ID 578 ITSELF IS NOT THE THING THAT CHANGED. Run 8 proved the call
    # failed, not that the ID is wrong; nothing here may quietly try another.
    ids = set(re.findall(r"FindControls?\(\$msoControlButton, (\d+),", block))
    assert ids == {"578"}, ids


def test_205_r6_r7_r10_r11_the_returned_control_proves_its_own_identity() -> None:
    """R6, R7, R10, R11. A non-null return is not an identification."""
    block = _compile_gate()
    # The control is asked what it is...
    assert "$controlId = [int]$control.Id" in block, (
        "the returned control is never asked for its Id"
    )
    assert "$controlType = [int]$control.Type" in block, (
        "the returned control is never asked for its Type"
    )
    # ...and its answers are compared to the criteria that were requested.
    assert "$idOk = ($controlId -eq 578)" in block, "the Id answer is not checked"
    assert "$typeOk = ($controlType -eq $msoControlButton)" in block, (
        "the Type answer is not checked"
    )
    # R10 and R11: both are RECORDED CHECKS, so a wrong Id or a wrong Type
    # fails the checklist that decides P5-CMP rather than passing quietly.
    labels = _check_labels(_text(SCENARIOS))
    assert "the control returned IS command Id 578" in labels, labels
    assert "the control returned IS an msoControlButton (Type 1)" in labels, labels
    # Neither answer may be short-circuited to a constant.
    for constant in ("$idOk = $true", "$typeOk = $true", "$controlProved = $true",
                     "$idOk = $True", "$typeOk = $True", "$controlProved = $True"):
        assert constant not in block, f"the control identity is hard-coded ({constant})"
    # THE CONSTANT AND THE LABEL AGREE. Changing msoControlButton's value in
    # one place would leave the checklist promising Type 1 while the gate asked
    # for something else, and the transcript would read as if nothing moved.
    value = re.search(r"\$msoControlButton = (\d+)", block)
    assert value, "msoControlButton has no value in the source"
    assert value.group(1) == "1", f"msoControlButton is {value.group(1)}, not 1"
    assert f"the control returned IS an msoControlButton (Type {value.group(1)})" in labels
    # And BOTH are required, not either.
    assert "$controlProved = ($idOk -and $typeOk)" in block, (
        "the two answers are not both required"
    )
    for weak in ("($idOk -or $typeOk)", "($typeOk -or $idOk)"):
        assert weak not in block, f"one answer alone proves the control ({weak})"


def test_206_r8_r9_execute_is_unreachable_without_a_proved_control() -> None:
    """R8, R9. No identification, no Enabled read, no Execute, no PASS."""
    block = _compile_gate()
    # R8, first half: the target-project gate still comes first.
    assert block.index("$targetIsActive =") < block.index("FindControl"), (
        "discovery happens before the gate knows whose project this is"
    )
    # R8, second half: Execute and BOTH Enabled reads live inside the branch
    # the control-identity proof guards, and nowhere else.
    guard = block.index("if ($controlProved) {")
    unproved, proved = block[:guard], block[guard:]
    for gated in ("$control.Execute()", "$before = [bool]$control.Enabled",
                  "$lastEnabled = [bool]$poll.Enabled"):
        assert gated in proved, f"{gated} is not on the proved-control path"
        assert gated not in unproved, (
            f"{gated} runs before the control has proved what it is"
        )
    # The old guard may not survive alongside the new one: `$null -ne $control`
    # is exactly the weaker test Run 8's correction replaces.
    assert "if ($null -ne $control) {\n                    $before" not in block, (
        "Execute is still reachable from a merely non-null control"
    )
    # R9: a null control is a FAIL. The existence check is a recorded check,
    # and the checklist is what decides the result.
    assert "'the Compile VBAProject command (ID 578) exists'" in block
    assert "($null -ne $control)" in block
    assert "$compileOk = Test-ChecklistOk $list" in _executable(SCENARIOS)
    # Nothing downgrades a missing control to a skip or a note-only outcome.
    assert "'SKIP'" not in block, "a missing control could be skipped"
    # And $controlProved starts false, so every path that does not prove the
    # control leaves Execute unreached.
    assert "$controlProved = $false" in block, (
        "the control-identity flag does not start closed"
    )
    assert block.index("$controlProved = $false") < guard


def test_207_r12_r13_a_compile_gate_failure_is_not_a_compile_verdict() -> None:
    """R12, R13. P5-CMP can fail for six reasons; one of them is a defect.

    Run 8 failed at command discovery with the compiler never invoked, and the
    dependency-gate line said the project does not compile. That reading was
    unavailable from the evidence.
    """
    scenarios = _text(SCENARIOS)
    # R13: the corrected statement is about the PREREQUISITE, and it sends the
    # reader to the checklist rather than naming a cause.
    assert "the whole-project VBA compile prerequisite was not " in scenarios, (
        "P5-ALL no longer states that the prerequisite was not established"
    )
    assert "See the P5-CMP checklist for the exact reason" in scenarios, (
        "the dependency gate does not point at the evidence"
    )
    # R12: and the reading Run 8 disproved is gone from the whole subtree.
    # ASSEMBLED, NOT WRITTEN OUT. A ban list that contains its own banned
    # phrases matches itself; the marker idiom exists for the lines that must
    # quote a phrase in order to forbid it.
    project = "the VBA project "                    # retired-authority
    banned = (project + "does not compile",         # retired-authority
              project + "doesn't compile",          # retired-authority
              "failing declaration")                # retired-authority
    offenders: list[str] = []
    for path in _authority_scan_files():
        for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            lowered = line.lower()
            for phrase in banned:
                if phrase.lower() in lowered and AUTHORITY_EXEMPTION_MARKER not in line:
                    offenders.append(f"{path.relative_to(PCCM_ROOT)}:{number}: {phrase!r}")
    assert not offenders, (
        "a P5-CMP failure is still read as a production compile defect:\n        "
        + "\n        ".join(offenders)
    )
    # The catch path says the same kind of thing: not completed, not "broken".
    assert "'not attempted: the whole-project compile gate could not be completed'" in scenarios


def test_208_r14_the_run_8_target_project_identity_code_is_unchanged() -> None:
    """R14. Run 8 proved this chain on Windows. It is frozen verbatim."""
    block = _compile_gate()
    for frozen in (
            "$targetProject = $Workbook.VBProject",
            "$activeProject = $vbe.ActiveVBProject",
            "$targetFile = [string]$targetProject.FileName",
            "$activeFile = [string]$activeProject.FileName",
            "$targetFull = [System.IO.Path]::GetFullPath($targetFile)",
            "$activeFull = [System.IO.Path]::GetFullPath($activeFile)",
            "[System.StringComparison]::OrdinalIgnoreCase",
            "$targetIsActive = $haveFiles -and $sameFile",
            "if (-not $targetIsActive) {"):
        assert frozen in block, f"Run-8-proved identity code changed: {frozen}"
    # The four labels Run 8 printed PASS for, still spelled the same way.
    labels = _check_labels(_text(SCENARIOS))
    for label in ("the Stage-B workbook exposes its VBProject",
                  "the VBE reports an active VBProject",
                  "both VBProjects name a file, so identity is comparable at all",
                  "the active VBE project IS the Stage-B workbook project (by file path)",
                  "the two VBProject names are recorded for diagnosis",
                  "the VBE command bars are reachable"):
        assert label in labels, f"a Run-8 PASS label was reworded: {label}"
    # And the run itself is recorded, so the freeze has a stated reason.
    assert "RUNTIME RUN 8" in _text(SCENARIOS), "Run 8 is not recorded in the harness"


def test_209_r16_the_diagnostic_probe_is_diagnostic_only_and_released() -> None:
    """R16. The second probe may inform. It may not decide, and it may not run."""
    block = _compile_gate()
    if "FindControls(" not in block:
        return                      # the probe is optional; nothing to police
    start = block.index("$probeControls = $null")
    end = block.index("$controlProved = $false")
    probe = block[start:end]
    # It only runs when the real lookup came back empty.
    assert "if ($null -eq $control) {" in block[:start + 200], (
        "the diagnostic probe runs even when the command was found"
    )
    # DIAGNOSTIC ONLY. It records notes; it contributes no check, so it cannot
    # move the verdict in either direction.
    assert "Add-Note" in probe, "the probe records nothing"
    assert "Add-Check" not in probe, "the diagnostic probe votes on the verdict"
    assert "Add-Phase5Result" not in probe, "the diagnostic probe emits a result"
    # AND IT IS NEVER A FALLBACK. Nothing it finds is executed, enabled-read,
    # or assigned into the control the gate acts on.
    for fallback in ("Execute()", ".Enabled", "$control = $probeControls",
                     "$probeControls.Item", "$probeControls[", "Item(1)"):
        assert fallback not in probe, f"the diagnostic probe became a fallback ({fallback})"
    # Its own throw is contained: a failed diagnostic is a note, not a FAIL.
    assert "} catch {" in probe, "a throw in the probe would fail the gate"
    # R16: the collection is released, on every path, in its own finally.
    assert "} finally {" in probe, "the probe has no finally"
    flat = re.sub(r"[ \t]+", " ", probe)
    assert "Release-Transient $probeControls 'CommandBarControls(probe)'" in flat, (
        "the diagnostic collection is never released"
    )
    assert "$probeControls = $null" in probe.split("Release-Transient", 1)[1], (
        "the diagnostic collection is not cleared after release"
    )
    # StrictMode 2.0 raises on a member read against $null, and Run 6 is the
    # precedent: the count is only read once the collection is known non-null.
    count_at = probe.index("$probeControls.Count")
    guard_at = probe.index("if ($null -eq $probeControls) {")
    assert guard_at < count_at, "the probe reads .Count without a null guard"


def test_210_r12_no_result_line_reads_a_gate_failure_as_a_compile_defect() -> None:
    """R12, from the other side: every P5-ALL dependency line, checked.

    test_207 bans the phrase Run 8 disproved. This one reads the lines that
    would carry it, so a rephrasing of the same overclaim is caught too.
    """
    scenarios = _text(SCENARIOS)
    lines = scenarios.splitlines()
    dependency_texts: list[str] = []
    for index, line in enumerate(lines):
        if "Add-Phase5Result 'P5-ALL'" not in line:
            continue
        chunk = "\n".join(lines[index:index + 8])
        dependency_texts.append(chunk[:chunk.index("return")] if "return" in chunk else chunk)
    assert len(dependency_texts) >= 3, len(dependency_texts)
    for text in dependency_texts:
        assert "not attempted" in text, text
        # A dependency gate reports that it did not run. It does not diagnose
        # the thing it was waiting for.
        for verdict in ("does not compile", "is broken", "has a compile error",
                        "the production project fails", "compile error in"):
            assert verdict not in text.lower(), (
                f"a dependency line diagnoses its prerequisite: {verdict!r}\n{text}"
            )
    # The compile prerequisite's own line names the alternatives rather than
    # picking one, and command discovery is among them because of Run 8.
    compile_line = next(text for text in dependency_texts
                        if "compile prerequisite" in text)
    for alternative in ("VBE access", "target-project identity",
                        "command discovery", "compiler diagnostic"):
        assert alternative in compile_line, (
            f"the compile prerequisite line does not admit {alternative!r}"
        )


def test_211_the_p5_cmp_evidence_chain_is_visible_and_in_order() -> None:
    """One scenario, one result, and a checklist that reads as a chain.

    P5-CMP answers one question, so it stays one scenario ID. But a reader has
    to be able to see WHERE it stopped, which is exactly what Run 8 needed: the
    identity links passed and the discovery link failed, and the transcript said
    so line by line. Each link is a recorded check, and they are recorded in the
    order they are established.
    """
    scenarios = _text(SCENARIOS)
    labels = _check_labels(scenarios)
    chain = ["the VBE object model is reachable",
             "the Stage-B workbook exposes its VBProject",
             "the VBE reports an active VBProject",
             "both VBProjects name a file, so identity is comparable at all",
             "the active VBE project IS the Stage-B workbook project (by file path)",
             "the two VBProject names are recorded for diagnosis",
             "the VBE command bars are reachable",
             "the Compile VBAProject command (ID 578) exists",
             "the control returned IS command Id 578",
             "the control returned IS an msoControlButton (Type 1)",
             "Compile VBAProject was executed at most once",
             "the compiled state was read by reacquiring the exact Id-578 control",
             "the target PCCM VBProject reached the VBE compiled state"]
    missing = [link for link in chain if link not in labels]
    assert not missing, f"links missing from the P5-CMP evidence chain: {missing}"
    positions = [labels.index(link) for link in chain]
    assert positions == sorted(positions), (
        "the evidence chain is recorded out of order: "
        f"{[chain[i] for i in sorted(range(len(chain)), key=lambda n: positions[n])]}"
    )
    # ONE SCENARIO, not one per link. Internal stages do not get their own IDs.
    ids = _executable(SCENARIOS)
    declared = _procedure(ids, "Get-Phase5ScenarioIds")
    for invented in ("P5-CMP1", "P5-CMPD", "P5-CMP-ID", "P5-DISC", "P5-CMP2"):
        assert invented not in declared, f"the gate split into {invented}"
    assert declared.count("'P5-CMP'") == 1


# =====================================================================
# RUNTIME RUN 9, AND THE MANUAL COMPILE THAT SETTLED IT
# =====================================================================
# Run 9 reached Execute(). Every link before it passed: the right project was
# active, the exact Id-578 msoControlButton was discovered, Enabled read True.
# Then the gate read Enabled again on the SAME handle, in the next statement,
# saw True, and called the project uncompiled.
#
# The retained artifact was afterwards opened by hand, and Debug > Compile
# VBAProject completed with no error and went grey. THE PRODUCTION PROJECT
# COMPILES. What the immediate read measured was the harness's own timing.
#
# The manual compile does not say whether Run 9's programmatic Execute finished
# after the harness stopped looking - that was a different session on a reopened
# file - so the conclusion these tests encode is the narrow one: an immediate
# post-Execute read of a cached handle is not a settlement proof, and settlement
# has to be observed in the same session by reacquiring the control.


def _settlement() -> str:
    """The bounded poll: from the released stale handle to the elapsed read."""
    block = _compile_gate()
    return block[block.index("$settled = $false"):block.index("$settleMs =")]


def test_212_r1_the_compile_command_is_executed_exactly_once() -> None:
    """R1. One Execute, guarded, and none inside the settlement poll."""
    block = _compile_gate()
    assert block.count("$control.Execute()") == 1, block.count("$control.Execute()")
    assert block.count(".Execute()") == 1, (
        "something other than the one proved control is being executed"
    )
    # It is counted at runtime too, so the transcript carries the evidence.
    assert "$executeCount = 0" in block and "$executeCount = 1" in block
    assert "'Compile VBAProject was executed at most once'" in block
    assert "($executeCount -le 1)" in block, "the execution count is never checked"
    # AND THE POLL NEVER RE-EXECUTES. This is the failure mode a naive retry
    # loop would have: compile, look, compile again, and call the second one
    # evidence about the first.
    settle = _settlement()
    for again in ("Execute()", "$executeCount = 2", "$executeCount + 1", "$executeCount++"):
        assert again not in settle, f"the settlement poll re-invokes the command ({again})"


def test_213_r2_the_compiled_state_is_not_read_from_the_cached_handle() -> None:
    """R2. The Run-9 root: `Enabled` one statement after Execute."""
    block = _compile_gate()
    # The retired read is gone in every spelling.
    for retired in ("$after = [bool]$control.Enabled", "$after = $control.Enabled",
                    "(-not $after)"):
        assert retired not in block, f"the immediate post-Execute read is back ({retired})"
    # And the handle it was read from is dropped straight after Execute, before
    # any settlement observation - so there is nothing stale left to read.
    execute_at = block.index("$control.Execute()")
    drop_at = block.index("Release-Transient $control 'CommandBarControl'")
    poll_at = block.index("$settled = $false")
    assert execute_at < drop_at < poll_at, (
        "the stale control is not released between Execute and the settlement poll"
    )
    assert "$control = $null" in block[drop_at:poll_at], (
        "the stale control handle is not cleared after being released"
    )
    # The verdict is the settlement flag, not any cached value.
    assert "'the target PCCM VBProject reached the VBE compiled state' `\n" \
           "                        $settled `" in _text(SCENARIOS), (
        "the compiled-state check is not decided by the settlement observation"
    )


def test_214_r3_the_settlement_poll_is_bounded_in_time_and_terminates() -> None:
    """R3. Five seconds, ~100 ms apart, and it stops on success."""
    settle = _settlement()
    block = _compile_gate()
    # A DEADLINE, computed once, from the clock.
    assert "$settleStarted = Get-Date" in block
    assert "$settleDeadline = $settleStarted.AddSeconds(5)" in block, (
        "the settlement window is not a five-second deadline"
    )
    assert "while ((-not $settled) -and ((Get-Date) -lt $settleDeadline)) {" in settle, (
        "the poll is not bounded by both the deadline and the success flag"
    )
    # THE INTERVAL.
    assert "Start-Sleep -Milliseconds 100" in settle, "the poll has no interval"
    assert settle.count("Start-Sleep") == 1, settle.count("Start-Sleep")
    # NO UNBOUNDED FORM. Any of these would let the gate hang a Windows run.
    for unbounded in ("while ($true)", "while (-not $settled) {", "do {", "for (;;)"):
        assert unbounded not in settle, f"the poll can run forever ({unbounded})"
    # It stops the moment the command goes quiet.
    assert "if (-not $lastEnabled) { $settled = $true }" in settle, (
        "the poll never sets its own success flag"
    )
    assert "$settled = $false" in block, "the settlement flag does not start closed"
    assert block.index("$settled = $false") < block.index("while ((-not $settled)")
    # And every early exit is a break out of the loop, never a swallow.
    assert settle.count("break") == 3, settle.count("break")
    # The elapsed time is measured and reported, so a transcript shows the shape
    # of the wait rather than only its verdict.
    assert "$settleMs = [int]((Get-Date) - $settleStarted).TotalMilliseconds" in block
    assert "'P5-CMP: settlement - '" in block, "the settlement is never reported"


def test_215_r4_r5_every_observation_reacquires_and_releases_the_exact_control() -> None:
    """R4, R5. A fresh handle each time, re-proved, and let go."""
    settle = _settlement()
    # R4: reacquisition through the same explicit criteria, inside the loop.
    assert "$poll = $bars.FindControl($msoControlButton, 578, $missing, $missing)" in settle, (
        "the poll does not reacquire the control through the exact criteria"
    )
    assert settle.index("Start-Sleep") < settle.index("$poll = $bars.FindControl"), (
        "the poll reacquires before waiting, so the first look is the stale one"
    )
    # ...and the reacquired control re-proves what it is, every time.
    assert "$pollId = [int]$poll.Id" in settle
    assert "$pollType = [int]$poll.Type" in settle
    assert "if (($pollId -ne 578) -or ($pollType -ne $msoControlButton)) {" in settle, (
        "a reacquired control's identity is not re-proved"
    )
    assert "$settleIdentityHeld = $false" in settle
    # A vanished control is a stated failure, not a silent end of loop.
    assert "$settleError = 'the Compile VBAProject control could not be reacquired'" in settle
    # R5: released before the next iteration, on every path out of the body,
    # including the two breaks and the throw.
    assert "} finally {" in settle, "the reacquired control has no finally"
    flat = re.sub(r"[ \t]+", " ", settle)
    assert "Release-Transient $poll 'CommandBarControl(settle)'" in flat
    assert settle.count("Release-Transient") == 1, settle.count("Release-Transient")
    assert "$poll = $null" in settle.split("Release-Transient", 1)[1], (
        "the reacquired handle is not cleared after release"
    )
    # It is re-nulled at the TOP of each iteration too, so a failed reacquisition
    # cannot leave the previous iteration's handle in scope.
    body = settle[settle.index("Start-Sleep"):]
    assert body.index("$poll = $null") < body.index("try {"), (
        "the poll handle is not reset before each reacquisition"
    )
    # StrictMode 2.0 raises on a member read against $null; the guard comes first.
    assert settle.index("if ($null -eq $poll) {") < settle.index("$pollId = [int]$poll.Id")
    # And the four accepted long-lived handles are untouched by the poll.
    for outer in ("$bars = $null", "$vbe = $null", "$targetProject = $null",
                  "$activeProject = $null"):
        assert outer not in settle, f"the poll disturbs a long-lived handle ({outer})"


def test_216_r6_r7_only_a_disabled_command_passes_and_a_timeout_diagnoses_nothing() -> None:
    """R6, R7. False is the PASS, and the deadline is not a compile verdict."""
    block = _compile_gate()
    settle = _settlement()
    # R6: PASS requires Enabled False, observed at least once, from a control
    # whose identity held throughout.
    assert "$lastEnabled = [bool]$poll.Enabled" in settle
    assert "$lastEnabled = $true" in block, (
        "the last-seen state does not start at the pessimistic value"
    )
    assert "if (-not $lastEnabled) { $settled = $true }" in settle
    for forced in ("$settled = $true\n", "$settled = $True\n"):
        assert forced not in block.replace(
            "if (-not $lastEnabled) { $settled = $true }", ""), (
            f"the settlement flag is set unconditionally somewhere ({forced!r})"
        )
    # A poll that never observed anything, or lost the control's identity, is
    # not a settlement - that is its own recorded check.
    assert "'the compiled state was read by reacquiring the exact Id-578 control'" in block
    assert "(($observations -gt 0) -and $settleIdentityHeld -and" in block, (
        "an empty or drifting poll could still be read as evidence"
    )

    # R7: the deadline wording says what was not established, and refuses to
    # diagnose the thing it was waiting for. The manual compile of this very
    # artifact is why that distinction is not cosmetic.
    assert "Compile VBAProject did not settle to the disabled/compiled " in block
    assert "state within the bounded observation window" in block
    assert "not a " in block and "compiler diagnostic. " in block, (
        "the timeout detail does not disclaim being a compiler diagnostic"
    )
    tail = block[block.index("'the target PCCM VBProject reached the VBE compiled state'"):]
    for verdict in ("does not compile", "compile error", "is broken",
                    "failing declaration", "undefined"):  # retired-authority
        assert verdict not in tail.lower(), (
            f"the settlement failure diagnoses the production project ({verdict})"
        )
    # And no caption is consulted anywhere in the settlement path.
    assert ".Caption" not in block


def test_217_r8_r9_the_accepted_gates_survive_the_settlement_correction() -> None:
    """R8, R9. Nothing above the poll moved, and nothing below it is lost."""
    block = _compile_gate()
    # R8: the Run-8-proved identity chain, statement by statement.
    for frozen in ("$targetProject = $Workbook.VBProject",
                   "$activeProject = $vbe.ActiveVBProject",
                   "$targetFull = [System.IO.Path]::GetFullPath($targetFile)",
                   "$activeFull = [System.IO.Path]::GetFullPath($activeFile)",
                   "[System.StringComparison]::OrdinalIgnoreCase",
                   "$targetIsActive = $haveFiles -and $sameFile",
                   "if (-not $targetIsActive) {",
                   "$controlProved = ($idOk -and $typeOk)"):
        assert frozen in block, f"an accepted check was disturbed: {frozen}"
    # The whole poll lives inside the identity-gated branch.
    else_at = block.index("} else {")
    assert block.index("$settled = $false") > else_at
    assert "$settled" not in block[:else_at], (
        "the settlement poll escaped the target-project gate"
    )
    # R9: a P5-CMP failure is still a return, and the lifecycle is downstream.
    scenarios = _executable(SCENARIOS)
    whole = scenarios[scenarios.index("function Invoke-Phase5GateBScenarios"):]
    section = whole[:whole.index("Save-Phase5LockedFxSeed")]
    gate_tail = section[section.index("$vbe = $Excel.VBE"):]
    for fatal in ("exit 1", "exit(", "[Environment]::Exit"):
        assert fatal not in gate_tail, f"the compile gate {fatal}s"
    assert "'SKIP'" not in gate_tail
    harness = _executable(HARNESS)
    order = [harness.index(token) for token in (
        "Invoke-Phase5GateBScenarios -Excel",
        "Add-Result 'Z' 'Excel closed naturally",
        "$transient = @(Get-TransientFailures)",
        "Add-Phase4FinalCompletenessResult -Results $results",
        "Add-Phase5LedgerIntegrityResult")]
    assert order == sorted(order)
    # And the manual evidence is recorded, so the freeze has a stated reason.
    text = _text(SCENARIOS)
    assert "RUNTIME RUN 9" in text
    assert "THE PRODUCTION VBA PROJECT COMPILES ON THE REAL TARGET ENVIRONMENT." in text


def test_218_every_wait_loop_in_the_gate_b_harness_is_bounded() -> None:
    """No loop in this harness may run forever on a Windows machine.

    The S2 mutation removed the settlement deadline and only one test noticed.
    A Gate-B run that hangs is worse than one that fails: it holds an Excel
    process open and produces no transcript at all. So the rule is stated once,
    over every loop in the harness, rather than only over the poll that
    happened to prompt it.
    """
    # A loop is bounded when its own head consults a clock or compares against
    # a finite quantity. A head that reads only a boolean flag is not: nothing
    # makes the flag change if the awaited thing never happens.
    bounded_by = ("Get-Date", "-lt ", "-le ", "-gt ", "-ge ")
    problems: list[str] = []
    for path in (SCENARIOS, HARNESS, LIFECYCLE):
        source = _executable(path)
        for match in re.finditer(r"^\s*(?:(do)\s*\{|while\s*\((.*)\)\s*\{)",
                                 source, re.MULTILINE):
            condition = (match.group(2) or "").strip()
            line = source[:match.start()].count("\n") + 1
            if match.group(1):
                problems.append(f"{path.name}:{line}: a do-loop, whose bound is not in its head")
                continue
            if not any(token in condition for token in bounded_by):
                problems.append(
                    f"{path.name}:{line}: while ({condition}) has no clock and no bound")
    assert not problems, "unbounded loops:\n  " + "\n  ".join(problems)
    # And the two that exist name their limit explicitly.
    assert "$settleDeadline = $settleStarted.AddSeconds(5)" in _executable(SCENARIOS)
    assert "$deadline = (Get-Date).AddSeconds($TimeoutSeconds)" in _executable(LIFECYCLE)


def test_219_the_settlement_poll_observes_and_never_acts() -> None:
    """A poll that acts on the thing it is measuring measures itself.

    The S8 mutation put a second `Execute()` inside the loop body. That is the
    shape of a retry masquerading as an observation: it would drive the command
    to disabled and then report the disabled state as evidence about the FIRST
    compile. Only one test caught it, so the read-only rule is stated here in
    its own right.
    """
    settle = _settlement()
    # NOTHING IN THE LOOP INVOKES ANYTHING. The only members touched on the
    # reacquired control are the three the observation needs.
    members = set(re.findall(r"\$poll\.(\w+)", settle))
    assert members == {"Id", "Type", "Enabled"}, members
    for action in ("Execute", "Delete", "Reset", "SetFocus", "Move", "Copy",
                   "$poll.Enabled =", "$poll.Visible ="):
        assert action not in settle, f"the settlement poll acts on the command ({action})"
    # NOR ON ANYTHING ELSE. No Run, no Import, no workbook or VBE mutation can
    # ride along inside a loop whose job is to look.
    for side_effect in ("$Excel.Run", "VBComponents", "$Workbook.", "Set-TableCell",
                        "Add-Phase5Result", "$vbe.", "$targetProject.", "$activeProject."):
        assert side_effect not in settle, f"the settlement poll has a side effect ({side_effect})"
    # The one Execute in the whole gate is upstream of the loop entirely.
    block = _compile_gate()
    assert block.count(".Execute()") == 1
    assert block.index(".Execute()") < block.index("$settled = $false"), (
        "the compile is invoked from inside the settlement window"
    )


# =====================================================================
# RUNTIME RUN 10: FOUR HARNESS DEFECTS AND ONE AUTHORITY DISCREPANCY
# =====================================================================
# Run 10 was the first runtime to clear the compile and fixture boundary and
# reach the broad Phase-5 functional matrix: 69 PASS / 5 FAIL / 0 SKIP, with
# P5-CMP and P5-FIX both passing on real Windows.
#
# Four of the five failures never reached a production predicate at all:
#
#   P5-AN  Set-Phase5Fixture returned System.Object[] - Add-BlankTableRow's row
#          index leaked into the pipeline alongside the Apply result
#   P5-PQ  COM 0x800A03EC from Names.Item - a defined-name MAP was dereferenced,
#          handing Excel a cell address where a name belonged
#   P5-PN  PropertyNotFoundException on 'repeat' - StrictMode 2.0 throws on the
#          read, so the $null guard around it was unreachable
#   P5-AR  PropertyNotFoundException on 'id' - the audit fixture was passed to a
#          checker that demanded a plan-case field it never needed
#
# The fifth, P5-ID's Risk central_basis, is an authority question and is NOT
# decided here. See test_226.


def _mutation_schema_model() -> dict[str, dict[str, object]]:
    """A pure port of Get-Phase5MutationSchema, read OUT OF the PowerShell.

    The table is parsed from the harness source rather than restated here, so a
    kind added or a field reclassified in one place cannot pass this suite while
    the other place still says something different.
    """
    block = _procedure(_executable(SCENARIOS), "Get-Phase5MutationSchema")
    schema: dict[str, dict[str, object]] = {}
    for match in re.finditer(r"'(\w+)'\s*=\s*@\{", block):
        kind = match.group(1)
        index, depth = match.end() - 1, 0
        while index < len(block):
            if block[index] == "{":
                depth += 1
            elif block[index] == "}":
                depth -= 1
                if depth == 0:
                    break
            index += 1
        body = block[match.end():index]
        required = re.search(r"Required = @\(([^)]*)\)", body, re.S)
        nullable = re.search(r"NullMeansBlank = @\(([^)]*)\)", body, re.S)
        optional = re.search(r"Optional = @\{([^}]*)\}", body, re.S)
        schema[kind] = {
            "required": re.findall(r"'(\w+)'", required.group(1)) if required else [],
            "null_ok": re.findall(r"'(\w+)'", nullable.group(1)) if nullable else [],
            "optional": {name: value.rstrip(";").strip()
                         for name, value in re.findall(r"(\w+) = (\S+)", optional.group(1))}
            if optional else {},
        }
    return schema


def _emitted_mutations() -> list[tuple[str, dict]]:
    gate_b = _gate_b()
    found: list[tuple[str, dict]] = []
    for bucket in ("prerequisite_cases", "no_block_cases", "direct_check_cases"):
        for case in gate_b.get(bucket, []) or []:
            where = f"{bucket} {case.get('id', '<no id>')}"
            if isinstance(case.get("mutation"), dict):
                found.append((where, case["mutation"]))
            for entry in case.get("mutations", []) or []:
                found.append((where, entry))
    return found


def test_220_a_the_fixture_boundary_emits_exactly_one_object() -> None:
    """A. P5-AN's root: a PowerShell function emits every uncaptured value."""
    source = _executable(SCENARIOS)
    boundary = _procedure(source, "Set-Phase5Fixture")
    # THE BOUNDARY COUNTS WHAT ITS BODY EMITTED, and refuses anything else.
    assert "$emitted = @(Invoke-Phase5FixtureSteps" in boundary, (
        "the boundary does not capture the steps' output at all"
    )
    assert "if ($emitted.Count -ne 1) {" in boundary, (
        "the boundary does not require exactly one emitted object"
    )
    assert "throw (" in boundary
    # AND IT NAMES WHAT LEAKED. A count alone would have left Run 10's
    # "System.Object[]" exactly as opaque as it was.
    assert "$item.GetType().Name" in boundary, (
        "a contract breach does not report what was in the pipeline"
    )
    # The single object is type- and value-checked before it is returned.
    assert "$applied = $emitted[0]" in boundary
    assert "if ($applied -isnot [string]) {" in boundary
    assert "if ($applied -notlike 'OK|*') {" in boundary
    assert "return $applied" in boundary
    # NOT FIXED AT THE CALLER. Either of these would have made P5-AN green while
    # still selecting one object out of a polluted pipeline.
    for laundering in ("$applied[-1]", "$applied[0] =", "[string]$emitted)",
                       "| Select-Object -Last 1", "$emitted[-1]"):
        assert laundering not in boundary, f"the boundary launders its input ({laundering})"


def test_221_a_no_helper_in_the_fixture_tree_leaks_a_return_value() -> None:
    """A. The leak itself, closed at every site and kept closed.

    Add-BlankTableRow returns the new row's Index. Six call sites ignored it and
    three of those are inside the fixture tree, so the indices joined the Apply
    result in Set-Phase5Fixture's output - the naked 2 / 3 / 2 in Run 10's
    transcript. This walks the whole harness, so a NEW value-returning helper
    called and ignored is caught before a runtime finds it.
    """
    sources = {path.name: _executable(path)
               for path in (SCENARIOS, HARNESS, LIFECYCLE)}
    bodies: dict[str, tuple[str, str]] = {}
    for name, text in sources.items():
        for match in re.finditer(r"^function ([\w-]+) \{", text, re.M):
            index, depth = match.end() - 1, 0
            while index < len(text):
                if text[index] == "{":
                    depth += 1
                elif text[index] == "}":
                    depth -= 1
                    if depth == 0:
                        break
                index += 1
            bodies[match.group(1)] = (name, text[match.start():index + 1])
    returning = {name for name, (_, body) in bodies.items()
                 if re.search(r"^\s*return\s+\S", body, re.M)}
    assert "Add-BlankTableRow" in returning, (
        "the helper whose return value caused the Run-10 leak no longer returns one; "
        "this test is no longer watching what it thinks it is"
    )

    # THE FIXTURE TREE, transitively, from the boundary down.
    tree, queue = set(), ["Set-Phase5Fixture"]
    while queue:
        name = queue.pop()
        if name in tree or name not in bodies:
            continue
        tree.add(name)
        queue.extend(re.findall(r"\b([A-Z][a-z]+-[A-Za-z][\w]*)\b", bodies[name][1]))
    for expected in ("Invoke-Phase5FixtureSteps", "Clear-Phase5Registers",
                     "Reset-Phase5FxTable", "Set-Phase5InflationProfileMaster",
                     "Write-Phase5InflationRates", "Write-Phase5Weights",
                     "Add-BlankTableRow"):
        assert expected in tree, f"{expected} is not reachable from the fixture boundary"

    leaks: list[str] = []
    for name in sorted(tree):
        file_name, body = bodies[name]
        for offset, line in enumerate(body.split("\n"), 1):
            statement = line.strip()
            if not statement:
                continue
            call = re.match(r"^([A-Z][a-z]+-[A-Za-z][\w]*)\b", statement)
            if call and call.group(1) in returning:
                leaks.append(f"{file_name}:{name}: {statement[:70]}")
    assert not leaks, (
        "a value-returning helper is called and ignored inside the fixture tree, "
        "so its result joins the fixture's own:\n  " + "\n  ".join(leaks)
    )


def test_222_a_the_second_fixture_is_the_condition_that_exposed_the_leak() -> None:
    """A. P5-FIX passed and P5-AN did not, and the difference is real.

    The golden fixture appended no FX row and no profile row, so nothing leaked.
    The analytical fixtures append both. A regression that only ever establishes
    the first fixture would not have caught this, which is why the corpus is
    asserted to contain a model that actually appends.
    """
    cases = _emitted()["cases"]
    # The FX seed is one row; a model naming more currencies than the reporting
    # one forces Add-BlankTableRow to run, which is the leak's precondition.
    appending = []
    for case in cases["plan_cases"]:
        model = case.get("model") or {}
        currencies = {entry["currency"] for entry in model.get("fx", []) or []}
        profiles = set((model.get("inflation") or {}).keys())
        if len(currencies) > 1 or len(profiles) > 1:
            appending.append(case["id"])
    assert appending, (
        "no emitted plan case appends an FX or profile row, so no fixture in the "
        "corpus reaches the append path that leaked in Run 10"
    )
    # And the appending models are analytical, so P5-AN really does re-fixture
    # over a workbook that already carries a previous fixture's state.
    analytical = {case["id"] for case in cases["plan_cases"]
                  if case.get("kind") == "analytical"}
    assert analytical.intersection(appending), (
        "the appending fixtures are never applied by the analytical scenario"
    )
    # The audit fixture appends too - P5-AR re-fixtures after P5-AN.
    audit = cases["gate_b"]["audit_reconstruction"]["model"]
    audit_currencies = {entry["currency"] for entry in audit["fx"]}
    assert len(audit_currencies) > 1 or len(audit["inflation"]) > 1, (
        "the audit fixture appends nothing, so P5-AR would not have re-exposed the leak"
    )


def test_223_b_every_emitted_mutation_is_well_formed_against_the_schema() -> None:
    """B. The schema table is real, and the emitted corpus satisfies it.

    The table is parsed out of the PowerShell rather than restated, so this
    cannot drift into asserting a schema the harness does not use.
    """
    schema = _mutation_schema_model()
    assert schema, "the mutation schema table could not be read from the harness"
    mutations = _emitted_mutations()
    assert len(mutations) >= 25, f"only {len(mutations)} mutations were found"

    problems: list[str] = []
    for where, mutation in mutations:
        kind = mutation.get("kind")
        if kind not in schema:
            problems.append(f"{where}: unknown kind {kind!r}")
            continue
        rule = schema[kind]
        for field in rule["required"]:
            if field not in mutation:
                problems.append(f"{where} [{kind}]: required {field!r} absent")
            elif mutation[field] is None and field not in rule["null_ok"]:
                problems.append(f"{where} [{kind}]: required {field!r} is null")
        for field in mutation:
            if field == "kind":
                continue
            if field not in rule["required"] and field not in rule["optional"]:
                problems.append(f"{where} [{kind}]: {field!r} is in no schema class")
    assert not problems, "the emitted mutation corpus is malformed:\n  " + "\n  ".join(problems)

    # EVERY KIND THE APPLIER HANDLES IS IN THE TABLE, and vice versa: a kind the
    # switch accepts but the schema does not describe would reach Excel unchecked.
    applier = _procedure(_executable(SCENARIOS), "Invoke-Phase5Mutation")
    handled = set(re.findall(r"^        '(\w+)' \{", applier, re.M))
    assert handled == set(schema), f"applier {sorted(handled)} vs schema {sorted(schema)}"


def test_224_b_every_named_and_entered_target_resolves_without_excel() -> None:
    """B. P5-PQ's root, proved against the emitted authorities.

    The harness dereferenced the manifest's defined_names MAP, which is keyed by
    name with the cell ADDRESS as its value, and handed Excel `'Setup'!$C$48`
    where `nmDuration_Entered` belonged. Names.Item raised COM 0x800A03EC and
    PQ-02 produced no production verdict at all.
    """
    source = _executable(SCENARIOS)
    applier = _procedure(source, "Invoke-Phase5Mutation")
    # THE RETIRED DEREFERENCE IS GONE, everywhere in the harness.
    for path in (SCENARIOS, HARNESS, LIFECYCLE):
        assert "defined_names.(" not in _executable(path), (  # retired-authority
            f"{path.name} still dereferences the defined_names map to build a name"
        )
    assert "$Manifest.defined_names" not in applier, (
        "the applier still reads a defined name out of the manifest's address map"
    )
    # The names come from the one table that holds NAMES.
    assert "$names = Get-Phase5EnteredStructureNames" in applier
    assert "$definedName = [string]$names[$target]" in applier
    assert "Set-NamedValue -Workbook $Workbook -DefinedName $definedName" in applier

    # AND THE TABLE IS RIGHT, checked against the emitted manifest: each value is
    # a declared defined NAME, and no value is an address.
    entered = _procedure(source, "Get-Phase5EnteredStructureNames")
    pairs = dict(re.findall(r"^\s+(\w+)\s*=\s*'([^']+)'", entered, re.M))
    assert set(pairs) == {"base_year", "start_year", "duration"}, pairs
    declared = set(_emitted()["manifest"]["defined_names"])
    for target, name in pairs.items():
        assert name in declared, f"{target} -> {name!r} is not a declared defined name"
        assert "!" not in name and "$" not in name, (
            f"{target} -> {name!r} looks like a cell address, not a name"
        )

    # THE PREFLIGHT CATCHES IT WITHOUT EXCEL. Same table, run in P5-PRE.
    preflight = _procedure(source, "Invoke-Phase5CoveragePreflight")
    assert "Test-Phase5MutationSchema" in preflight, (
        "a malformed mutation still has to reach Excel to be discovered"
    )
    assert "'every emitted mutation is well formed and every target resolves (no Excel)'" \
        in preflight
    assert "stage_b_manifest.json" in preflight, (
        "the preflight cannot check entered-structure names without the manifest"
    )
    # And the schema really does resolve the two target classes.
    checker = _procedure(source, "Test-Phase5MutationSchema")
    assert "$kind -like 'named_*'" in checker
    assert "$kind -eq 'entered_structure'" in checker
    assert "names no declared input" in checker
    assert "carries no defined_name" in checker
    # A required target that does not resolve is an ERROR, never a default.
    assert "-Default" not in checker.split("$kind -like 'named_*'")[1].split("}")[0]


def test_225_c_optional_properties_are_read_through_the_property_collection() -> None:
    """C. P5-PN's root: under StrictMode 2.0 the READ throws, not the guard."""
    source = _executable(SCENARIOS)
    applier = _procedure(source, "Invoke-Phase5Mutation")
    value = _procedure(source, "Get-MutationValue")

    # THE RETIRED IDIOM IS GONE. `$null -ne $Mutation.<optional>` cannot run.
    for retired in ("$null -ne $Mutation.repeat", "$null -ne $Mutation.append",  # retired-authority
                    "if ($Mutation.append)", "if ($Mutation.apply_timeline)",  # retired-authority
                    "if ($Mutation.require_clean_structure)", "$raw = $Mutation.$Property"):  # retired-authority
        assert retired not in applier + value, f"the StrictMode hazard is back ({retired})"

    # The two genuinely optional properties go through the collection, with the
    # defaults the schema states.
    assert "Get-Phase5OptionalProperty -Object $Mutation -Name 'repeat' -Default 1" in applier
    assert "Get-Phase5OptionalProperty -Object $Mutation -Name 'append' -Default $false" in applier
    assert "$raw = Get-Phase5OptionalProperty -Object $Mutation -Name $Property" in value

    # ABSENT AND PRESENT-WITH-NULL STAY DIFFERENT. The accessor returns the
    # default only when the property is not there; a stated null is returned.
    accessor = _procedure(source, "Get-Phase5OptionalProperty")
    assert "if (-not (Test-Phase5HasProperty -Object $Object -Name $Name)) { return $Default }" \
        in accessor
    assert "return $Object.PSObject.Properties[$Name].Value" in accessor
    has = _procedure(source, "Test-Phase5HasProperty")
    assert "$Object.PSObject.Properties[$Name]" in has
    assert "$Object.$Name" not in has, "the presence test itself reads the property"

    # REQUIRED FIELDS NEVER DEFAULT.
    required = _procedure(source, "Get-Phase5RequiredProperty")
    assert "throw (" in required
    assert "A required field is never defaulted." in required
    assert "-Default" not in required
    # And the required flags are read through it, not directly.
    assert applier.count("Get-Phase5RequiredProperty -Object $Mutation -Name 'apply_timeline'") == 2
    assert "Get-Phase5RequiredProperty -Object $Mutation -Name 'require_clean_structure'" in applier

    # STRICTMODE IS NOT WEAKENED ANYWHERE.
    for path in (SCENARIOS, HARNESS, LIFECYCLE):
        text = _executable(path)
        for weakening in ("Set-StrictMode -Off", "-Version 1.0", "-Version Latest"):
            assert weakening not in text, f"{path.name} weakens StrictMode ({weakening})"
    assert "Set-StrictMode -Version 2.0" in _executable(SCENARIOS)

    # THE CORPUS REALLY EXERCISES BOTH SHAPES, so this is not a hypothetical.
    fx_rows = [m for _, m in _emitted_mutations() if m.get("kind") == "fx_row"]
    assert any("repeat" not in m for m in fx_rows), "no fx_row mutation omits repeat"
    assert any("repeat" in m for m in fx_rows), "no fx_row mutation states repeat"
    assert any("append" not in m for m in fx_rows), "no fx_row mutation omits append"
    assert any(m.get("rate") is None for m in fx_rows), (
        "no fx_row mutation states a null rate, so the blank-cell semantic is untested"
    )


def test_226_d_the_audit_fixture_is_checked_against_its_own_expected_block() -> None:
    """D. P5-AR's root: a display label demanded a plan-case field."""
    source = _executable(SCENARIOS)
    checker = _procedure(source, "Add-Phase5AnalyticalChecks")
    # THE LABEL IS THE CALLER'S, and `id` is not read at all.
    assert "[string]$Label" in checker
    assert "$Case.id" not in checker, (
        "the analytical checker still demands a plan-case id"
    )
    assert "$Label" in checker
    # A caller that forgets, or a fixture with no expected block, fails loudly
    # rather than asserting nothing.
    assert "if ([string]::IsNullOrWhiteSpace($Label)) {" in checker
    assert "if ($null -eq $expected) {" in checker
    assert checker.count("throw (") >= 2

    # EVERY CALL SITE NAMES ITS FIXTURE.
    joined = re.sub(r"`\s*\n\s*", " ", source)
    calls = re.findall(r"Add-Phase5AnalyticalChecks -List[^\n]*", joined)
    assert len(calls) == 5, f"{len(calls)} call sites"
    for call in calls:
        assert "-Label " in call, call[:120]
    audit_calls = [call for call in calls if "-Case $audit" in call]
    assert len(audit_calls) == 1
    assert "$audit.title" in audit_calls[0], (
        "the audit fixture's label does not come from the emitted fixture itself"
    )
    assert "$audit.id" not in source, "an id is invented for the audit fixture"

    # NO EXPECTED VALUE IS CONSTRUCTED IN POWERSHELL. The audit fixture carries
    # the same emitted `expected` shape a plan case does, from the same oracle.
    cases = _emitted()["cases"]
    audit = cases["gate_b"]["audit_reconstruction"]
    plan = next(case for case in cases["plan_cases"] if case.get("kind") == "analytical")
    assert set(audit["expected"]) == set(plan["expected"]), (
        f"the audit expected block has a different shape: "
        f"{sorted(set(audit['expected']) ^ set(plan['expected']))}"
    )
    assert set(audit["model"]) == set(plan["model"])
    assert "id" not in audit, (
        "the corpus now emits an id for the audit fixture; the harness must not "
        "start depending on it without the schema saying so"
    )


# ---------------------------------------------------------------------
# E. P5-ID: Risk central_basis. AN AUTHORITY QUESTION, NOT A HARNESS BUG
# ---------------------------------------------------------------------
# P5-ID ran deeply. Cases 3, 9 and 30 all committed, and the reconciliation,
# annual, totals and weights checks passed. One published value disagreed:
#
#     case 9, R-001.central_basis, actual BLANK, expected 'ML'
#
# That is not arithmetic. It is a question about which side owns the field, and
# these two tests answer only what the evidence answers: the three declared
# authorities agree with each other, and production is the one that differs.
# NOTHING IS CHANGED IN EITHER PRODUCTION OR THE ORACLE on this round.

CALC_DRIVERS_CONTRACT_PATH = PCCM_ROOT / "spec" / "calc_contract.yaml"
DRIVERS_BLOCK_PATH = PCCM_ROOT / "src" / "vba" / "modCalcReport.bas"


def _calc_drivers_applies_to() -> dict[str, list[str]]:
    """`applies_to` for every tblCalcDrivers column, from the accepted contract."""
    text = CALC_DRIVERS_CONTRACT_PATH.read_text(encoding="utf-8")
    start = text.rfind("columns:", 0, text.index('- key: "central_basis"'))
    segment = text[start:]
    found: dict[str, list[str]] = {}
    for name, body in re.findall(
            r'- key: "(\w+)"(.*?)(?=\n      - key: "|\Z)', segment, re.S):
        applies = re.search(r"applies_to: \[([^\]]*)\]", body)
        if applies and name not in found:
            found[name] = re.findall(r'"(\w+)"', applies.group(1))
    return found


def _drivers_block_branches() -> tuple[set[str], set[str]]:
    """The calc_drivers columns production leaves Empty, per driver kind."""
    text = DRIVERS_BLOCK_PATH.read_text(encoding="utf-8")
    start = text.index("Private Function DriversBlock(")
    body = text[start:text.index("\nEnd Function", start)]
    risk_at = body.index("If package.Model.Drivers(index).IsRisk Then")
    else_at = body.index("\n        Else\n", risk_at)
    def blanked(chunk: str) -> set[str]:
        return {name.lower() for name in re.findall(
            r"COL_CALC_DRIVERS_(\w+)\) = Empty", chunk)}
    return blanked(body[risk_at:else_at]), blanked(body[else_at:])


def test_227_e_the_three_declared_authorities_agree_on_central_basis() -> None:
    """E. Contract, plan and oracle say the same thing. They must keep doing so.

    This is the test that stops `central_basis` being quietly set to null in the
    Python oracle to make P5-ID green. Doing that would put the oracle at odds
    with both the accepted contract and the accepted plan, which is a bigger
    problem than the one it would appear to solve.
    """
    # 1. THE CONTRACT.
    applies = _calc_drivers_applies_to()
    assert applies["central_basis"] == ["cost_line", "risk"], applies["central_basis"]
    # It is used discriminatingly elsewhere, so it is a real applicability
    # statement rather than boilerplate on every column.
    assert applies["quantity"] == ["cost_line"], applies["quantity"]
    assert applies["probability"] == ["risk"], applies["probability"]
    assert applies["central_value"] == ["cost_line"], applies["central_value"]
    assert applies["expected_risk_nominal"] == ["risk"]

    # 2. THE PLAN's own column table, which spells out blank where it means it.
    plan = (PCCM_ROOT / "docs" / "phase5_plan.md").read_text(encoding="utf-8")
    rows = {}
    for line in plan.splitlines():
        cells = [cell.strip() for cell in line.split("|")]
        if len(cells) >= 9 and cells[1].isdigit():
            rows[cells[2]] = (cells[6], cells[7])
    assert rows.get("Central Basis") == ("yes", "yes"), rows.get("Central Basis")
    assert rows.get("Quantity") == ("yes", "**blank**"), rows.get("Quantity")
    assert rows.get("Central Value") == ("yes", "**blank**"), rows.get("Central Value")

    # 3. THE ORACLE. Every emitted Risk driver row carries a basis, and every
    # Cost Line row does too - the sweep the review asked for, over every case
    # rather than only case 9.
    risk_rows, risk_with_basis, cost_rows, cost_missing = 0, 0, 0, 0
    offenders: list[str] = []
    cases = _emitted()["cases"]
    fixtures = [(f"plan {case['id']}", case) for case in cases["plan_cases"]]
    fixtures.append(("audit_reconstruction", cases["gate_b"]["audit_reconstruction"]))
    for where, fixture in fixtures:
        for row in (fixture.get("expected") or {}).get("drivers", []):
            if row.get("driver_kind") == "Risk":
                risk_rows += 1
                if row.get("central_basis") is not None:
                    risk_with_basis += 1
                else:
                    offenders.append(f"{where}: {row.get('permanent_id')}")
            else:
                cost_rows += 1
                if row.get("central_basis") is None:
                    cost_missing += 1
    assert risk_rows >= 4, f"only {risk_rows} Risk rows are emitted at all"
    assert risk_with_basis == risk_rows, (
        "the oracle no longer publishes central_basis for every Risk row, which "
        "puts it at odds with the accepted contract and plan: " + "; ".join(offenders)
    )
    assert cost_missing == 0, f"{cost_missing} Cost Line rows carry no central_basis"
    assert cost_rows >= 20


def test_228_e_production_matches_the_contract_applicability_exactly() -> None:
    """E. RESOLVED after Run 10: zero deviations, in both directions.

    This test used to allow one - `risk_deviation <= {"central_basis"}` - while
    the authority question was open. The decision came back: Central Basis
    applies to Cost Line AND Risk, per the contract's applies_to, the accepted
    plan's own column table and the Python oracle. Production was the defect and
    has been corrected, so the exception is retired and the requirement is now
    exact agreement.
    """
    applies = _calc_drivers_applies_to()
    risk_blank, cost_blank = _drivers_block_branches()
    assert risk_blank, "the Risk branch of DriversBlock blanks nothing at all"
    assert cost_blank, "the Cost Line branch of DriversBlock blanks nothing at all"

    risk_deviation = {name for name in risk_blank
                      if "risk" in applies.get(name, [])}
    cost_deviation = {name for name in cost_blank
                      if "cost_line" in applies.get(name, [])}
    assert not risk_deviation, (
        "production blanks a column the contract says applies to Risk: "
        f"{sorted(risk_deviation)}"
    )
    assert not cost_deviation, (
        "production blanks a column the contract says applies to Cost Line: "
        f"{sorted(cost_deviation)}"
    )
    # central_basis specifically: it is published for Risk now, not blanked.
    assert "central_basis" not in risk_blank, (
        "the P5-ID defect is back: the Risk branch publishes Central Basis blank"
    )
    # AND THE OTHER DIRECTION: production must never publish a column the
    # contract excludes for that kind.
    for name, kinds in applies.items():
        if "risk" not in kinds:
            assert name in risk_blank, (
                f"the contract excludes {name} for Risk but production does not blank it"
            )
        if "cost_line" not in kinds:
            assert name in cost_blank, (
                f"the contract excludes {name} for Cost Line but production does not "
                "blank it"
            )
    # The decision is recorded where a reader will find it, as a decision.
    record = (PCCM_ROOT / "docs" / "phase5_gate_b_harness.md").read_text(encoding="utf-8")
    assert "central_basis" in record
    assert "RESOLVED AFTER RUN 10" in record, (
        "the record does not mark the P5-ID authority question as decided"
    )


def test_229_the_row_index_of_every_added_table_row_is_accounted_for() -> None:
    """A, harness-wide. The specific helper whose index leaked in Run 10.

    test_221 walks the fixture tree and bans ANY ignored value-returning helper.
    This one follows Add-BlankTableRow everywhere, including the three call sites
    outside that tree - a leak in a mutation applier or in P5-NS pollutes
    whatever function encloses it just as surely, and the naked 2 / 3 / 2 in
    Run 10's transcript came from exactly that class of site.
    """
    unaccounted: list[str] = []
    for path in (SCENARIOS, HARNESS, LIFECYCLE):
        source = _executable(path)
        for number, line in enumerate(source.split("\n"), 1):
            statement = line.strip()
            if "Add-BlankTableRow" not in statement:
                continue
            if statement.startswith("$null = Add-BlankTableRow"):
                continue
            if re.match(r"^\$\w+ = Add-BlankTableRow", statement):
                continue
            if "Add-BlankTableRow" not in statement.split("#")[0]:
                continue
            if not statement.startswith("Add-BlankTableRow"):
                continue
            unaccounted.append(f"{path.name}:{number}: {statement[:70]}")
    assert not unaccounted, (
        "Add-BlankTableRow returns the new row's Index and it is being dropped "
        "into the enclosing function's output:\n  " + "\n  ".join(unaccounted)
    )
    # The helper still returns what this test thinks it returns.
    assert "return [int]$added.Index" in _executable(HARNESS)
    # And every site is suppressed rather than captured-and-ignored, so a reader
    # sees the intent at the call.
    suppressed = _executable(SCENARIOS).count("$null = Add-BlankTableRow")
    assert suppressed == 6, f"{suppressed} suppressed Add-BlankTableRow calls, expected 6"


def test_230_the_four_run_10_idioms_stay_retired_across_the_subtree() -> None:
    """A second, independent reading of all four Run-10 roots.

    The structural tests above prove the corrected code is present. This one
    proves the DEFECTIVE code is absent, everywhere, by the text that made it
    defective - so a correction reverted in one file while the structure stays
    plausible elsewhere is still caught.
    """
    mutation = "$Mutation."  # retired-authority
    retired = (
        # P5-AN: the leak, at the site that caused it
        "\nAdd-BlankTableRow -Workbook",                         # retired-authority
        # P5-PQ: a defined-name MAP dereferenced into a name argument
        "defined_names.(",                                       # retired-authority
        # P5-PN: a direct read of an optional property under StrictMode 2.0
        "$null -ne " + mutation + "repeat",                      # retired-authority
        "$null -ne " + mutation + "append",                      # retired-authority
        "if (" + mutation + "append)",                           # retired-authority
        "if (" + mutation + "apply_timeline)",                   # retired-authority
        "if (" + mutation + "require_clean_structure)",          # retired-authority
        "$raw = " + mutation + "$Property",                      # retired-authority
        "[int]" + mutation + "repeat",                           # retired-authority
        "[bool]" + mutation + "append",                          # retired-authority
        # P5-AR: a plan-case field demanded of a fixture that has none
        "'case ' + [string]$Case.id",                            # retired-authority
        "[string]$Case.id",                                      # retired-authority
    )
    offenders: list[str] = []
    for path in _authority_scan_files():
        text = path.read_text(encoding="utf-8")
        for number, line in enumerate(text.splitlines(), 1):
            if AUTHORITY_EXEMPTION_MARKER in line:
                continue
            for phrase in retired:
                needle = phrase.lstrip("\n")
                if phrase.startswith("\n"):
                    if line.startswith(needle) or line.strip().startswith(needle):
                        offenders.append(f"{path.relative_to(PCCM_ROOT)}:{number}: {phrase!r}")
                elif needle in line:
                    offenders.append(f"{path.relative_to(PCCM_ROOT)}:{number}: {phrase!r}")
    assert not offenders, (
        "a Run-10 defect idiom is back:\n  " + "\n  ".join(offenders)
    )


# ---------------------------------------------------------------------
# SECOND DETECTORS. The A2/A3/B2/C2/C3/C4/D2/D3 mutations each tripped
# exactly one test, because the tests above bundle several claims each.
# These split those claims along their real seams: the tests below ask
# the same questions from a different angle, not the same assertion twice.
# ---------------------------------------------------------------------


def test_231_the_fixture_boundary_decision_table_is_exactly_three_refusals() -> None:
    """A. WHEN the boundary refuses, enumerated - not merely THAT it refuses.

    test_220 reads the boundary structurally. This reads it as a decision table,
    so widening a guard (`-ne 1` to `-lt 1`, which would wave an array through)
    or changing which object is taken is a change to the table.
    """
    boundary = _procedure(_executable(SCENARIOS), "Set-Phase5Fixture")
    conditions: list[str] = []
    for match in re.finditer(r"if \((.*?)\) \{", boundary):
        index, depth = boundary.index("{", match.start()), 0
        while index < len(boundary):
            if boundary[index] == "{":
                depth += 1
            elif boundary[index] == "}":
                depth -= 1
                if depth == 0:
                    break
            index += 1
        if "throw (" in boundary[match.end():index]:
            conditions.append(re.sub(r"\s+", " ", match.group(1)).strip())
    assert conditions == [
        "$emitted.Count -ne 1",
        "$applied -isnot [string]",
        "$applied -notlike 'OK|*'",
    ], conditions
    # EXACTLY ONE means exactly one. A lower bound would accept the array that
    # Run 10 actually received.
    assert "-lt 1" not in boundary and "-gt 1" not in boundary and "-ge 1" not in boundary, (
        "the count guard is a bound, not an equality, so a polluted pipeline passes"
    )
    # AND THE OBJECT TAKEN IS INDEX 0 - the only index that means anything once
    # the count is known to be exactly one. Any other index is a selection out
    # of a pipeline the boundary has just refused to accept.
    taken = re.search(r"\$applied = \$emitted\[([^\]]*)\]", boundary)
    assert taken and taken.group(1) == "0", taken.group(1) if taken else None
    assert boundary.count("$emitted[") == 1, boundary.count("$emitted[")


def test_232_the_preflight_really_walks_the_corpus_and_records_what_it_finds() -> None:
    """B. The preflight's mutation loop, as a loop - not as a call that exists.

    Emptying the body while leaving the check label in place would report
    "every emitted mutation is well formed" over nothing at all.
    """
    preflight = _procedure(_executable(SCENARIOS), "Invoke-Phase5CoveragePreflight")
    # ALL THREE BUCKETS, by name.
    assert "foreach ($bucket in 'prerequisite_cases', 'no_block_cases', 'direct_check_cases')" \
        in preflight, "the preflight does not walk every bucket of mutations"
    # BOTH SHAPES: a case with one `mutation` and a case with a `mutations` list.
    assert "-Name 'mutation'" in preflight and "-Name 'mutations'" in preflight
    # THE RESULT IS ACCUMULATED FROM THE CHECKER, not assigned a constant.
    accumulation = re.search(
        r"\$malformed \+= @\(([^)]*)", preflight)
    assert accumulation, "the preflight never accumulates a schema result"
    assert "Test-Phase5MutationSchema" in accumulation.group(1), accumulation.group(1)
    assert "$malformed += @()" not in preflight, "the accumulation was emptied"
    # AND IT IS COUNTED, so a walk that found nothing to check is itself a FAIL.
    assert "$mutationCount++" in preflight
    assert "'the corpus emitted at least one Gate-B mutation'" in preflight
    assert "($mutationCount -gt 0)" in preflight
    assert "($malformed.Count -eq 0)" in preflight


def test_233_the_optional_accessor_defaults_on_absence_only() -> None:
    """C. The semantics, stated as semantics.

    test_225 proves the retired reads are gone and the accessors are called.
    This proves the accessors MEAN what the schema assumes: absence takes the
    default, a stated null is returned as null, and a required field throws.
    """
    source = _executable(SCENARIOS)
    accessor = _procedure(source, "Get-Phase5OptionalProperty")
    # THE DEFAULT IS REACHED FROM THE PRESENCE TEST, never from the value.
    assert "if (-not (Test-Phase5HasProperty -Object $Object -Name $Name)) { return $Default }" \
        in accessor, "the accessor does not default on absence"
    assert "$null -eq $value" not in accessor and "$null -eq $raw" not in accessor, (
        "the accessor defaults on a NULL VALUE, which collapses 'absent' into "
        "'present and deliberately null' - the corpus distinguishes them"
    )
    assert accessor.count("return $Default") == 1, accessor.count("return $Default")
    assert "return $Object.PSObject.Properties[$Name].Value" in accessor

    # THE REQUIRED ACCESSOR HAS NO DEFAULT AT ALL.
    required = _procedure(source, "Get-Phase5RequiredProperty")
    assert "-Default" not in required, "the required accessor can default"
    assert "return $null" not in required, "a missing required field returns null"
    assert "A required field is never defaulted." in required

    # AND THE DEFAULTS AT THE CALL SITES ARE THE ONES THE SCHEMA DECLARES.
    schema = _mutation_schema_model()
    declared = schema["fx_row"]["optional"]
    assert declared == {"repeat": "1", "append": "$false"}, declared
    applier = _procedure(source, "Invoke-Phase5Mutation")
    for field, value in declared.items():
        assert f"-Name '{field}' -Default {value}" in applier, (
            f"the applier's default for {field} does not match the schema's {value}"
        )


def test_234_a_corpus_null_still_writes_a_blank_cell() -> None:
    """C. The distinction the accessor exists to protect, at its consumer.

    Several locked prerequisites ARE the blank. If the accessor defaulted on a
    null value, `Get-MutationValue` would stop returning $null and the writers
    would put a real number where the predicate needs an absence.
    """
    source = _executable(SCENARIOS)
    value = _procedure(source, "Get-MutationValue")
    assert "$raw = Get-Phase5OptionalProperty -Object $Mutation -Name $Property" in value
    # No default is supplied here, so a stated null stays null...
    assert "-Name $Property -Default" not in value, (
        "Get-MutationValue supplies a default, so a corpus null becomes a value"
    )
    assert "if ($null -eq $raw) { return $null }" in value
    # ...and the accessor it relies on must not default on a value either.
    accessor = _procedure(source, "Get-Phase5OptionalProperty")
    assert "$null -eq" not in accessor.split("return $Object.PSObject")[0].split(
        "Test-Phase5HasProperty")[-1], "the accessor inspects the value before defaulting"
    # THE SCHEMA NAMES WHICH NULLS ARE MEANINGFUL, and the corpus really has them.
    schema = _mutation_schema_model()
    assert "rate" in schema["fx_row"]["null_ok"]
    assert "value" in schema["register_cell"]["null_ok"]
    nulls = [(where, mutation) for where, mutation in _emitted_mutations()
             if any(mutation.get(field) is None for field in mutation)]
    assert nulls, "no emitted mutation states a null, so the blank path is untested"


def test_235_the_labelling_contract_is_enforced_at_both_ends() -> None:
    """D. Every call names its fixture, and the checker refuses one that does not."""
    source = _executable(SCENARIOS)
    joined = re.sub(r"`\s*\n\s*", " ", source)
    calls = re.findall(r"Add-Phase5AnalyticalChecks -List[^\n]*", joined)
    unlabelled = [call for call in calls if "-Label " not in call]
    assert not unlabelled, (
        "an analytical call site does not name the fixture it is about:\n  "
        + "\n  ".join(call[:110] for call in unlabelled)
    )
    # Each label is BUILT, not a bare constant repeated at every site.
    for call in calls:
        label = call.split("-Label ", 1)[1]
        assert label.startswith("("), label[:80]
        assert "$" in label.split(")")[0], f"a label names no fixture: {label[:80]}"
    # THE OTHER END. A caller that forgets is refused rather than labelled
    # 'case ' with nothing after it.
    checker = _procedure(source, "Add-Phase5AnalyticalChecks")
    assert "if ([string]::IsNullOrWhiteSpace($Label)) {" in checker, (
        "the checker accepts a call with no label"
    )
    assert "if ($false)" not in checker and "if ($true)" not in checker
    guard = checker[checker.index("IsNullOrWhiteSpace($Label)"):]
    assert "throw (" in guard[:400], "an unlabelled call does not throw"


# =====================================================================
# RUNTIME RUN 11: A WILDCARD OPERATOR WHERE A SUBSTRING TEST BELONGED
# =====================================================================
# Run 11 returned 73 PASS / 1 FAIL / 0 SKIP. The single failure was P5-PQ, and
# within it only the detail-token subchecks of PQ-25 and PQ-26. Production had
# returned the accepted message verbatim:
#
#     risk R-001: Probability must be a fraction in [0, 1]
#
# against the accepted token `fraction in [0, 1]` - the token is visibly inside
# the detail - and the check still failed.
#
# PowerShell's -like is a WILDCARD operator. `[0, 1]` is a character class
# matching ONE character from {0 , space 1}, and the character after
# "fraction in " is "[", which is not one of them. The predicate had fired
# correctly; the checker was wrong.
#
# These tests are about SEMANTICS, not spelling. They model both matchers and
# assert what each decides, so a future rewrite that happens to avoid the exact
# retired line but reintroduces pattern matching is still caught.

RUN11_DETAIL = "risk R-001: Probability must be a fraction in [0, 1]"
RUN11_TOKEN = "fraction in [0, 1]"


def _wildcard_like(text: str, pattern: str) -> bool:
    """PowerShell `-like`, modelled: * ? and [...] are pattern syntax.

    fnmatch implements the same three constructs with the same meanings, which
    is all this needs: the point is that a bracket expression is a character
    class in both, and that is the whole of the Run-11 root.
    """
    import fnmatch

    return fnmatch.fnmatchcase(text.lower(), pattern.lower())


def _literal_contains(text: str, token: str) -> bool:
    """The accepted contract: the token occurs as literal text, case-insensitively."""
    if not token.strip():
        return False
    return token.lower() in text.lower()


def test_236_a_the_run_11_reproducer_matches_literally() -> None:
    """A. The exact detail and token Run 11 rejected."""
    # THE OLD MATCHER REJECTED IT. If this stops being true the reproducer has
    # gone stale and the rest of this test proves nothing.
    assert not _wildcard_like(RUN11_DETAIL, "*" + RUN11_TOKEN + "*"), (
        "the wildcard matcher no longer fails the Run-11 case; the modelled "
        "semantics have drifted from PowerShell's"
    )
    # THE ACCEPTED CONTRACT ACCEPTS IT.
    assert _literal_contains(RUN11_DETAIL, RUN11_TOKEN)
    assert RUN11_TOKEN in RUN11_DETAIL, "the token is not literally in the detail at all"

    # AND THE HARNESS IMPLEMENTS THE LITERAL RULE.
    body = _procedure(_executable(SCENARIOS), "Add-Phase5DetailTokenChecks")
    assert "IndexOf(" in body, "the matcher does not perform a substring search"
    assert "[System.StringComparison]::OrdinalIgnoreCase" in body, (
        "the substring search is not case-insensitive; -like was"
    )
    assert "-ge 0" in body, "the IndexOf result is not tested for a hit"


def test_237_b_a_wildcard_match_is_not_accepted_as_a_literal_one() -> None:
    """B. The divergence, in the direction that matters for evidence.

    A pattern match that is not a substring match would let a token "find" a
    predicate name that was never printed. That is a false PASS in an evidence
    harness, which is worse than Run 11's false FAIL.
    """
    detail, token = "the value is 7 units", "is [0-9] units"
    assert _wildcard_like(detail, "*" + token + "*"), (
        "the control pair no longer diverges; pick another"
    )
    assert not _literal_contains(detail, token), (
        "the literal rule accepts a token the detail does not contain"
    )
    # The same shape for the other two metacharacters.
    for detail, token in (("refused: rate 3 missing", "rate ? missing"),
                          ("refused: quantity absent", "quantity*absent")):
        assert _wildcard_like(detail, "*" + token + "*"), (detail, token)
        assert not _literal_contains(detail, token), (detail, token)


def test_238_c_every_emitted_token_is_found_by_the_literal_rule() -> None:
    """C. The whole corpus, swept through the accepted semantics.

    For each emitted token, a detail that literally contains it must match. The
    tokens carrying wildcard metacharacters are the ones the old matcher could
    have rejected, and they are named so the sweep cannot quietly become empty.
    """
    gate_b = _gate_b()
    tokens: set[str] = set()
    for bucket in ("prerequisite_cases", "direct_check_cases", "no_block_cases"):
        for case in gate_b.get(bucket, []) or []:
            tokens.update(case.get("detail_tokens", []) or [])
    for values in (gate_b.get("plan_refusal_tokens") or {}).values():
        tokens.update(values or [])
    assert len(tokens) >= 30, f"only {len(tokens)} detail tokens were found"

    # NO EMPTY TOKEN. One would be found at index 0 by any substring search and
    # would prove nothing; the matcher fails it, and the corpus must not emit one.
    assert not [token for token in tokens if not token.strip()], "an empty token is emitted"

    wildcard_bearing = sorted(token for token in tokens
                              if any(ch in token for ch in "[]*?"))
    assert RUN11_TOKEN in wildcard_bearing, (
        "the Run-11 token is no longer emitted, or no longer contains brackets"
    )
    for token in tokens:
        detail = f"production refused: {token} and nothing else"
        assert _literal_contains(detail, token), token
    # ...and for the metacharacter-bearing ones, the OLD matcher would have
    # rejected that very detail. That is the blast radius, measured.
    for token in wildcard_bearing:
        detail = f"production refused: {token} and nothing else"
        assert not _wildcard_like(detail, "*" + token + "*"), (
            f"{token!r} no longer demonstrates the wildcard defect"
        )


def test_239_d_the_wildcard_anti_pattern_cannot_return() -> None:
    """D. Source-locked, across the harness, in every spelling."""
    body = _procedure(_executable(SCENARIOS), "Add-Phase5DetailTokenChecks")
    for retired in ("$Detail -like", "-like ('*'", "-match", "-notlike",
                    "[regex]", "Select-String"):
        assert retired not in body, f"the token matcher uses {retired}"
    # The token never becomes part of a pattern anywhere in the function.
    assert "'*' + $token" not in body and "$token + '*'" not in body, (
        "the token is still being wrapped in wildcards"
    )
    # EVERY CALLER GOES THROUGH THE SHARED MATCHER. A caller that rolled its own
    # comparison would not be fixed by fixing this one.
    scenarios = _executable(SCENARIOS)
    callers = re.findall(r"Add-Phase5DetailTokenChecks -List", scenarios)
    assert len(callers) == 3, f"{len(callers)} callers; P5-DC, P5-RF and P5-PQ expected"
    for scenario in ("P5-DC", "P5-RF", "P5-PQ"):
        assert f"'{scenario}'" in scenarios
    # No scenario compares a detail token itself.
    stray = [line.strip() for line in scenarios.splitlines()
             if "detail_tokens" in line and "-like" in line]
    assert not stray, stray


def test_240_the_matcher_fails_an_empty_token_rather_than_passing_it() -> None:
    """An empty discriminator is found at index 0 by any substring search."""
    body = _procedure(_executable(SCENARIOS), "Add-Phase5DetailTokenChecks")
    assert "$found = $false" in body, "the result does not start closed"
    assert "if (-not [string]::IsNullOrWhiteSpace($literal)) {" in body, (
        "an empty token is passed straight to IndexOf, which finds it at 0"
    )
    assert body.index("$found = $false") < body.index("IndexOf(")
    # And the modelled rule agrees.
    assert not _literal_contains("anything at all", "")
    assert not _literal_contains("anything at all", "   ")
    # The old matcher passed it vacuously, which is what this closes.
    assert _wildcard_like("anything at all", "**")


def test_241_the_index_of_result_is_tested_as_a_found_threshold() -> None:
    """W5. `IndexOf` returns -1 for absent and >= 0 for present.

    test_236 asserts the search exists. This asserts the COMPARISON decides the
    right way round, by evaluating it against the two results IndexOf can
    actually produce - so a threshold that accepts -1 (a token that was never
    printed, reported as found) fails here whatever it is spelled.
    """
    body = _procedure(_executable(SCENARIOS), "Add-Phase5DetailTokenChecks")
    match = re.search(r"::OrdinalIgnoreCase\)\s*(-\w+)\s*(-?\d+)\)", body)
    assert match, f"the IndexOf result is not compared to a number:\n{body}"
    operator, operand = match.group(1), int(match.group(2))

    def decides(result: int) -> bool:
        table = {"-ge": result >= operand, "-gt": result > operand,
                 "-ne": result != operand, "-eq": result == operand,
                 "-lt": result < operand, "-le": result <= operand}
        assert operator in table, f"unmodelled comparison operator {operator}"
        return table[operator]

    # NOT FOUND must be rejected. This is the half that matters: accepting -1
    # would report every token as present, including one production never wrote.
    assert not decides(-1), (
        f"`{operator} {operand}` treats IndexOf's not-found result as a hit"
    )
    # FOUND, at any position, must be accepted.
    for position in (0, 1, 12, 4096):
        assert decides(position), (
            f"`{operator} {operand}` rejects a token found at index {position}"
        )


# ===========================================================================
# 31. PHASE-6 STEP 12: the Gate-B temporary-directory leak, closed
# ===========================================================================
# `_emitted()` is called by more than fifty tests in this file and used to leave
# its `mkdtemp` tree behind on every one of them. Repeated runs accumulated tens
# of thousands of `pccm-gateb-*` directories and exhausted the writable
# filesystem. The debt has been carried OPEN since Phase-6 Step 4; Step 12 is
# where it closes, because Step 13 runs on Windows and cannot afford it.
def _gateb_temp_dirs() -> set[str]:
    """The `pccm-gateb-*` directories that exist RIGHT NOW, by name.

    Names only, and only this prefix. Nothing here deletes anything: a
    concurrent pytest process, or a developer's own run, may own directories
    with the same prefix, and removing those would be a far worse defect than
    the leak this closes.
    """
    root = Path(tempfile.gettempdir())
    return {entry.name for entry in root.glob("pccm-gateb-*")}


def test_224_the_emitted_helper_leaves_no_temporary_directory_behind() -> None:
    """THE REGRESSION. The real helper, called repeatedly, leaks nothing.

    This drives `_emitted()` itself - not a copy of it, not a model of it - and
    compares the `pccm-gateb-*` set before and after. Any directory that appears
    and stays is a leak.
    """
    before = _gateb_temp_dirs()
    for _ in range(3):
        emitted = _emitted()
        # AND THE RESULT IS STILL USABLE AFTER THE TREE IS GONE, which is the
        # half a naive cleanup gets wrong: everything was read into memory
        # before the directory closed.
        assert emitted["manifest"]["vba"]["modules"], "the manifest came back empty"
        assert emitted["cases"]["plan_cases"], "the plan cases came back empty"
        assert emitted["inspection"], "the inspection came back empty"
        assert emitted["calc_module"].startswith("Attribute VB_Name")
        assert emitted["constants"].startswith("Attribute VB_Name")
    after = _gateb_temp_dirs()
    leaked = after - before
    assert leaked == set(), f"the helper leaked {len(leaked)} directories: {sorted(leaked)}"
    # NO DANGLING HANDLE. A path returned out of the helper would be invalid the
    # moment the tree closed, so the key that used to carry one is gone.
    assert "dir" not in emitted, (
        "the helper hands back a path into a directory it has already deleted"
    )


def test_225_the_helper_uses_a_self_cleaning_temporary_directory() -> None:
    """The source shape, so a future edit cannot quietly reopen the debt."""
    text = _text(Path(__file__))
    body = text[text.index("def _emitted() -> dict:"):]
    body = body[:body.index("\ndef _ps_string_literals")]
    assert 'with tempfile.TemporaryDirectory(prefix="pccm-gateb-") as name:' in body, (
        "the emitted tree is not created inside a self-cleaning context"
    )
    assert "tempfile.mkdtemp" not in body, (
        "a bare mkdtemp is back in the helper; it never removes its directory"
    )
    # The reads happen INSIDE the context, so nothing is read after cleanup.
    context_at = body.index("with tempfile.TemporaryDirectory")
    for read in ("json.loads(_text(stage_b.manifest_path))",
                 "json.loads(_text(calc_artifacts.cases_path))",
                 "json.loads(_text(inspection.path))",
                 "_text(calc_artifacts.module_path)",
                 "_text(stage_b.module_path)"):
        assert body.index(read) > context_at, read
    assert '"dir": tmp' not in body, "the dangling path handle is back"


def test_nc_110_a_leaking_temporary_directory_helper_is_caught() -> None:
    """MUTATION CONTROL, run for real: restore the bare mkdtemp and leak.

    The mutated helper is executed - it really does create a directory and
    really does fail to remove it - and the detector this round added is run
    against it. A control that only inspected text would not prove the check
    can see an actual leak.
    """
    before = _gateb_temp_dirs()

    def leaking_emitted() -> dict:
        # THE PRE-STEP-12 SHAPE, verbatim in behaviour: mkdtemp, no cleanup.
        tmp = Path(tempfile.mkdtemp(prefix="pccm-gateb-"))
        return {"dir": tmp}

    damaged = leaking_emitted()
    try:
        after = _gateb_temp_dirs()
        leaked = after - before
        assert leaked, "the mutation did not actually leak a directory"
        assert damaged["dir"].name in leaked
        # THE DETECTOR MUST SEE IT. This is the same comparison test_224 makes.
        try:
            assert after - before == set(), "leak"
        except AssertionError:
            pass
        else:
            raise AssertionError("the leak survived the before/after comparison")
    finally:
        # The control cleans up after ITSELF, and only after itself: it removes
        # the one directory it created, by name, and touches nothing else.
        shutil.rmtree(damaged["dir"], ignore_errors=True)
    assert not damaged["dir"].exists()
    assert _gateb_temp_dirs() - before == set(), "the control leaked its own mutation"


# ===========================================================================
# 32. PHASE-6 STEP 12: no fixed production-module count in active wording
# ===========================================================================
# P5-M and P5-D8 are manifest-driven in the scenarios file and always were. The
# DESCRIPTIONS were not: the driver banner and the harness record both said
# "15 modules" and "inventory back to 15", which was true when Phase 5 closed
# and stopped being true the moment Phase 6 added its first module. A number
# that has to be edited by hand every time the manifest grows is a second
# inventory authority, and this refuses one.
_MANIFEST_LANGUAGE = ("manifest module set", "manifest")

# Wording that pins the CURRENT inventory to a literal, in any of the forms this
# document has used for it. These are refused inside ACTIVE description blocks
# only: a historical Run table that records `present 30 of 15` is reporting what
# that run actually said, and rewriting it would destroy evidence.
_STALE_INVENTORY_WORDING = ("back to 15", "at 15", "re-asserted at 15",
                            "inventory of 15", "15 by name")

HARNESS_RECORD = PCCM_ROOT / "docs" / "phase5_gate_b_harness.md"


def _active_block(doc: str, heading: str) -> str:
    """One active description block: from its heading to the next `## ` one.

    The `## ` boundary is what separates the current architectural description
    from the Runtime Run sections below it, which are historical evidence.
    """
    start = doc.index(heading)
    tail = doc[start + len(heading):]
    end = tail.index("\n## ") if "\n## " in tail else len(tail)
    return heading + tail[:end]


def _assert_no_fixed_inventory(block: str, where: str) -> None:
    """THE SEMANTIC DETECTOR, shared by the conformance test and the control."""
    counts = re.findall(r"\b(\d+)\s+modules?\b", block)
    assert counts == [], f"{where} states a fixed module count: {counts}"
    lowered = block.lower()
    for stale in _STALE_INVENTORY_WORDING:
        assert stale not in lowered, f"{where} pins the inventory: {stale!r}"


def test_226_the_active_p5m_and_p5d8_descriptions_name_no_module_count() -> None:
    banner = _text(HARNESS)
    banner = banner[banner.index("      P5-M "):banner.index("      P5-AN ")]
    assert "P5-D8" in banner, "the banner slice lost P5-D8"
    _assert_no_fixed_inventory(banner, "the driver banner")
    for description in ("P5-M", "P5-D8"):
        line = banner[banner.index(description):]
        line = line[:line.index("\n      P5-")] if "\n      P5-" in line else line
        assert any(word in line for word in _MANIFEST_LANGUAGE), (
            f"{description} no longer describes itself in manifest terms"
        )
    # THE OTHER ACTIVE DESCRIPTION: the harness record's scenario table, which
    # says what each scenario ESTABLISHES. The Run-by-Run tables further down
    # are historical evidence and keep the numbers that were true when they were
    # written - Step 12 corrects active wording, it does not rewrite history.
    doc = HARNESS_RECORD.read_text(encoding="utf-8")
    table = _active_block(doc, "## The Windows scenarios")
    rows = [line for line in table.splitlines()
            if line.startswith("| `P5-M` |") or line.startswith("| `P5-D8` |")]
    assert len(rows) == 2, rows
    for row in rows:
        _assert_no_fixed_inventory(row, "the Windows-scenario table")
        assert any(word in row for word in _MANIFEST_LANGUAGE), row
    # AND NO ACTIVE ROW ANYWHERE STATES A DIGIT-FORM MODULE COUNT.
    for line in doc.splitlines():
        if line.startswith("| `P5-M` |") or line.startswith("| `P5-D8` |"):
            assert re.search(r"\b\d+\s+modules?\b", line) is None, line
    # THE THIRD ACTIVE DESCRIPTION, missed by the first pass of this detector
    # and found by independent review: the `### Lifecycle` block is a CURRENT
    # architectural description of the harness sequence, not a Run record, and
    # it said "inventory re-asserted at 15".
    lifecycle = _active_block(doc, "### Lifecycle")
    assert "P5-D8" in lifecycle, "the Lifecycle slice lost P5-D8"
    assert "P5-D0" in lifecycle and "P5-AN" in lifecycle, "the slice is too narrow"
    _assert_no_fixed_inventory(lifecycle, "the active Lifecycle block")
    d8 = lifecycle[lifecycle.index("P5-D8"):]
    d8 = d8[:d8.index("P5-AN")]
    assert any(word in d8 for word in _MANIFEST_LANGUAGE), (
        f"the Lifecycle P5-D8 line is not manifest-owned: {d8!r}"
    )
    # THE SLICE STOPS AT THE HISTORY BOUNDARY. Everything from the next `## `
    # heading onward is Runtime Run evidence and is not inspected. The block's
    # own prose may CITE a run - "Runtime Run 7 passed it and then met a VBE
    # compile error" is a cross-reference, not a Run record - so the boundary is
    # the heading, not the words.
    assert not re.search(r"^## ", lifecycle, re.M), (
        "the active slice ran past a section heading"
    )
    assert "present 30 of 15" not in lifecycle, (
        "the slice reached the historical Run-2 inventory row"
    )
    # AND THE HISTORY IT STOPS SHORT OF IS STILL THERE, untouched.
    assert "`present 30 of 15`" in doc, (
        "the historical Run-2 inventory evidence was rewritten"
    )


def test_227_the_executable_inventory_logic_is_still_manifest_driven() -> None:
    """Only the WORDING moved. The checks were manifest-owned already."""
    source = _executable(SCENARIOS)
    assert "'Persisted project: manifest module set by name, 5 buttons, 6 API procedures'" \
        in source, "the P5-M result name is no longer manifest-owned"
    assert "'Transient diagnostic module removed; inventory back to the manifest module set'" \
        in source, "the P5-D8 result name is no longer manifest-owned"
    # AND NO NUMERIC MODULE COUNT ANYWHERE IN THE EXECUTABLE SCENARIOS.
    counts = re.findall(r"\b(\d+)\s+modules?\b", source)
    assert counts == [], f"the scenarios compare against a hardcoded count: {counts}"


def test_nc_111_a_fixed_module_count_in_the_active_wording_is_caught() -> None:
    """MUTATION CONTROL. Put "15 modules" back and require the detector to see it."""
    banner = _text(HARNESS)
    banner = banner[banner.index("      P5-M "):banner.index("      P5-AN ")]
    assert re.findall(r"\b(\d+)\s+modules?\b", banner) == [], "the accepted text is clean"
    for planted in ("The persisted project: 15 modules BY NAME",
                    "the inventory back to 15 modules",
                    "inventory returned to the 16 modules"):
        damaged = banner + "\n      " + planted + "\n"
        assert re.findall(r"\b(\d+)\s+modules?\b", damaged), (
            f"the detector cannot see a fixed count in {planted!r}"
        )
    # The row form too.
    damaged_row = "| `P5-M` | 15 modules **by name**, exactly 5 buttons |"
    assert re.search(r"\b\d+\s+modules?\b", damaged_row), (
        "the detector cannot see a fixed count in a table row"
    )


def test_nc_112_the_stale_lifecycle_inventory_wording_is_caught() -> None:
    """MUTATION CONTROL for the description independent review found.

    The accepted Lifecycle block is taken, its P5-D8 line is mutated back to the
    exact wording that shipped in 15706c5, and the REAL detector is run against
    it. The second half is the point: the same detector must NOT reject the
    historical Run-2 evidence row, so what is protected is the active/history
    boundary rather than a global ban on the digits.
    """
    doc = HARNESS_RECORD.read_text(encoding="utf-8")
    lifecycle = _active_block(doc, "### Lifecycle")
    # The accepted block passes.
    _assert_no_fixed_inventory(lifecycle, "the accepted Lifecycle block")

    accepted_line = ("P5-D8  the diagnostic module is REMOVED; inventory re-asserted "
                     "against the\n       manifest module set")
    assert lifecycle.count(accepted_line) == 1, lifecycle
    for planted in (
        "P5-D8  the diagnostic module is REMOVED; inventory re-asserted at 15",
        "P5-D8  the diagnostic module is REMOVED; inventory back to 15",
        "P5-D8  the diagnostic module is REMOVED; 15 modules re-asserted",
        "P5-D8  the diagnostic module is REMOVED; inventory of 15 restored",
    ):
        damaged = lifecycle.replace(accepted_line, planted)
        assert damaged != lifecycle, planted
        try:
            _assert_no_fixed_inventory(damaged, "damaged")
        except AssertionError:
            pass
        else:
            raise AssertionError(f"the detector cannot see {planted!r}")

    # AND THE HISTORY IS NOT COLLATERAL DAMAGE, on two counts.
    #
    # FIRST: the Run-2 row records what that run actually reported. It survives
    # the detector on its own terms - it names no module count and pins no
    # current inventory - so the number in it is safe even from a global scan.
    historical = ("| R1 inventory semantics | P5-M, P5-D8 | `present 30 of 15`, "
                  "`extra: ThisWorkbook, shDashboard, …` |")
    assert historical in doc, "the historical Run-2 evidence row was rewritten"
    _assert_no_fixed_inventory(historical, "the historical Run-2 row")

    # SECOND, and this is what the slice is actually for: wording the detector
    # DOES reject stays legal below the boundary. Historical sections are full
    # of sentences describing what an old run asserted, and the active-block
    # scope is what keeps them out of scope.
    below_the_line = "Run 2 reported the inventory back to 15 modules by name."
    try:
        _assert_no_fixed_inventory(below_the_line, "planted")
    except AssertionError:
        pass
    else:
        raise AssertionError("the detector is blind to the planted history wording")
    boundary = "## Analytical and refusal coverage"
    assert doc.count(boundary) == 1, doc.count(boundary)
    planted_doc = doc.replace(boundary, boundary + "\n\n" + below_the_line, 1)
    assert below_the_line in planted_doc
    # The slice is unchanged: the sentence landed after the boundary heading, so
    # the active Lifecycle block never sees it.
    assert _active_block(planted_doc, "### Lifecycle") == lifecycle, (
        "the active slice absorbed text from below its boundary"
    )
    _assert_no_fixed_inventory(_active_block(planted_doc, "### Lifecycle"),
                               "the slice under planted history")


# ===========================================================================
# REVIEW ROUND 4A - THE WITHDRAWN `-is [double]` RATIONALE
# ===========================================================================
def test_the_is_double_rationale_is_withdrawn_not_merely_edited_away() -> None:
    """A false rationale is a defect even when the implementation is right.

    The exact-type setter is correct and stays. What was wrong was the reason
    given for it: that `-is [double]` "would be true for a boxed Int32 under
    PowerShell's numeric conversions". That claim is FALSE and is withdrawn.
    `-is` is a .NET instance test - it asks
    whether the object IS of that type and performs no numeric conversion - so
    `1 -is [double]` is $false. The old setter normalised Int32 through its
    explicit `-is [int]` branch and the `[double]$Value` cast, not through the
    `-is [double]` test.

    This control requires the claim to appear ONLY inside an explicit
    withdrawal, never as a live rationale, in both the test file and the
    active record.
    """
    # ASSEMBLED FROM PARTS so this control's own source does not contain the
    # literal it hunts for - otherwise it would only ever be finding itself.
    false_claim = "would be true " + "for a boxed"
    for path in (Path(__file__), HARNESS_RECORD):
        lines = path.read_text(encoding="utf-8").splitlines()
        hits = [i for i, line in enumerate(lines) if false_claim in line]
        assert hits, (
            f"{path.name} no longer mentions the withdrawn rationale at all; "
            "this control can no longer prove it was retracted rather than "
            "quietly deleted"
        )
        for index in hits:
            # Every surviving occurrence must sit inside a retraction: the
            # withdrawal words appear within a few lines either side.
            window = "\n".join(lines[max(0, index - 8): index + 9]).lower()
            assert "withdrawn" in window and "false" in window, (
                f"{path.name}:{index + 1} states the `-is [double]` rationale "
                "without withdrawing it"
            )

    # THE ACCURATE ACCOUNT IS PRESENT, and names the real cause.
    #
    # WHITESPACE IS NORMALISED FIRST. Markdown wraps prose and prefixes block
    # quotes with `> `, so a phrase that reads as one sentence is not one
    # contiguous string in the file. Searching the raw text for it would fail
    # on formatting and pass on nothing.
    record = HARNESS_RECORD.read_text(encoding="utf-8")
    flat = " ".join(record.replace("\n> ", " ").replace("\n", " ").split())
    assert "performs no numeric conversion" in flat
    assert "`1 -is [double]` is `$false`" in flat
    for real_cause in ("-is [int]", "[double]$Value"):
        assert real_cause in flat, real_cause
    # And the correction really is attached to the setter paragraph, not filed
    # somewhere unrelated.
    assert "Correction (review round 4A)" in flat

    # AND THE IMPLEMENTATION IS UNTOUCHED BY THE CORRECTION: exactly the four
    # approved captured types, matched by exact CLR type name.
    setter = _typed_setter()
    assert "$Value.GetType().FullName -ceq 'System.Double'" in setter
    assert "if ($null -eq $Value) {" in setter
    assert "$cell.Value2 = [string]$Value" in setter
    assert "$cell.Value2 = [bool]$Value" in setter
    assert "$cell.Value2 = [double]$Value" in setter
    assert setter.count("$cell.Value2") == 3
    for retired in ("$Value -is [single]", "$Value -is [int]", "$Value -is [long]",
                    "$Value -is [decimal]", "$Value -is [int16]", "$Value -is [byte]",
                    "$Value -is [datetime]"):
        assert retired not in setter, f"{retired} is back"


def test_the_ledger_verdict_has_exactly_one_owner_and_one_invocation() -> None:
    """REVIEW ROUND 4A §6. P5-LDG bypasses the guarded reporter by design.

    Because it is emitted through `Add-Result` rather than `Add-Phase5Result`,
    the ledger cannot suppress the result that reports ON the ledger. The price
    of that exemption is that nothing else guards it, so the exclusivity has to
    be structural: one owning procedure, one driver call, and no second emitter
    anywhere in the scenario source.
    """
    source = _executable(SCENARIOS)
    harness = _executable(HARNESS)

    # ONE OWNING PROCEDURE.
    assert source.count("function Add-Phase5LedgerIntegrityResult") == 1
    integrity = _procedure(source, "Add-Phase5LedgerIntegrityResult")

    # AND NO EMITTER OUTSIDE IT. The two occurrences inside are the PASS and
    # FAIL arms of the one verdict, which is why this counts rather than
    # asserting presence.
    inside = integrity.count("Add-Result 'P5-LDG'")
    assert inside == 2, (inside, "the verdict no longer has exactly a PASS and a FAIL arm")
    assert source.count("Add-Result 'P5-LDG'") == inside, (
        "P5-LDG is emitted from somewhere other than its owning procedure"
    )
    assert "Add-Result 'P5-LDG'" not in harness, (
        "the driver emits the verdict directly, bypassing the emitted-once flag"
    )

    # ONE DRIVER INVOCATION.
    assert harness.count("Add-Phase5LedgerIntegrityResult") == 1

    # DELIBERATELY OUTSIDE THE GUARDED REPORTER, and still emitted once.
    assert "Add-Phase5Result 'P5-LDG'" not in source, (
        "the verdict was routed through the ledger it reports on"
    )
    assert "if ($script:Phase5LedgerReported) { return }" in integrity
    assert integrity.index("if ($script:Phase5LedgerReported) { return }") < \
        integrity.index("Add-Result 'P5-LDG'"), (
        "the emitted-once flag is checked after the verdict is already emitted"
    )



# ===========================================================================
# THE STALE SCOPED-GRANT ASSERTION THAT BLOCKED PHASE 6 IN RUNTIME RUN 3
# ===========================================================================
# P5-EV asserted `RunSimulation is still forbidden in every module`. That was
# true when it was written and stopped being true the moment Step 11 granted the
# endpoint to modSimReport. In Run 3 it was the ONLY failed check in the only
# failed Phase-5 scenario, P6-PRE then correctly failed closed on it, and the
# entire Phase-6 behavioural matrix went unexecuted - because the harness
# disagreed with the accepted contract about a grant the contract had made.
#
# The generic module-aware executable scan passed on the real persisted project
# in the same run. The defect was a SECOND, hand-written assertion that had not
# moved with the contract, and these controls exist so that class cannot recur.
def _scoped_grants() -> dict[str, list[str]]:
    """Every scoped forbidden-construct rule, from the accepted contract."""
    import yaml
    contract = yaml.safe_load(_text(SPEC / "structure_contract.yaml"))
    return {
        rule["construct"]: list(rule["allowed_in"])
        for rule in contract["vba"]["forbidden_constructs"]
        if isinstance(rule, dict)
    }


def test_ev_01_the_contract_grants_run_simulation_to_exactly_one_owner() -> None:
    """One construct, one owner, named. Not "at least" and not "including"."""
    grants = _scoped_grants()
    assert grants.get("RunSimulation") == ["modSimReport"], grants.get("RunSimulation")
    assert grants.get("MRG32k3a") == ["modSimRng"], grants.get("MRG32k3a")
    # AND THE EMITTED MANIFEST AGREES, since that is what the harness reads.
    emitted = {
        rule["construct"]: list(rule["allowed_in"])
        for rule in _emitted()["manifest"]["vba"]["forbidden_construct_rules"]
    }
    for construct, owners in grants.items():
        assert emitted.get(construct) == owners, (construct, emitted.get(construct))


def test_ev_02_every_scoped_grant_is_checked_as_a_grant() -> None:
    """A granted construct asserted as globally forbidden is a contradiction.

    The predicate is chosen from the CONTRACT, not from memory: a construct with
    owners must be checked with `Test-ConstructScopedTo` naming that owner, and
    a construct with none with `Test-ConstructForbiddenGlobally`.
    """
    source = _executable(SCENARIOS)
    for construct, owners in _scoped_grants().items():
        if not owners:
            continue
        assert len(owners) == 1, (construct, owners)
        scoped = (
            f"(Test-ConstructScopedTo -Manifest $Manifest -Construct '{construct}' "
            f"-ModuleName '{owners[0]}')"
        )
        assert scoped in source, (
            f"P5-EV does not check the scoped grant for {construct} against "
            f"{owners[0]}"
        )
        stale = (
            f"(Test-ConstructForbiddenGlobally -Manifest $Manifest "
            f"-Construct '{construct}')"
        )
        assert stale not in source, (
            f"P5-EV still asserts that {construct} is forbidden globally, which "
            f"the contract contradicts: it is granted to {owners[0]}"
        )
        # And the check's WORDING must not claim the opposite of what it tests.
        for line in source.splitlines():
            if f"'{construct} " in line and "Add-Check" in line:
                assert "forbidden in every module" not in line, line


def test_ev_03_a_globally_forbidden_construct_is_still_checked_globally() -> None:
    """The correction must not turn every rule into a grant."""
    source = _executable(SCENARIOS)
    grants = _scoped_grants()
    for handler in ("Worksheet_Change", "Workbook_SheetChange"):
        assert grants.get(handler, []) == [], (handler, grants.get(handler))
    assert "foreach ($handler in 'Worksheet_Change', 'Workbook_SheetChange')" in source
    assert "Test-ConstructForbiddenGlobally -Manifest $Manifest -Construct $handler" in source


def test_ev_04_the_generic_module_aware_scan_stays_load_bearing() -> None:
    """The two explicit scoped checks are additional, never a replacement.

    The generic scan is what reads the REAL persisted project; the scoped checks
    read the manifest. Losing the first would trade evidence about what Excel
    holds for evidence about what the build said.
    """
    source = _executable(SCENARIOS)
    for required in ("Get-ForbiddenConstructRules",
                     "Test-ConstructForbiddenIn",
                     "no forbidden construct exists in the EXECUTABLE code of the "
                     "real Stage-B project"):
        assert required in source, required
    # It walks components and their code, not the manifest alone.
    scan = source.split("Get-Phase5VbComponentInventory")[0]
    assert "$offenders" in source
    assert "Release-Transient $module 'CodeModule'" in source


@contextmanager
def _aimed(scenarios: str | None = None, spec_contract: str | None = None):
    """Point the P5-EV controls at damaged copies for one mutation.

    Module globals, restored on the exception path too. Nothing is written to
    the repository: the copies live in a temporary directory that is removed on
    the way out.
    """
    global SCENARIOS, SPEC
    saved = (SCENARIOS, SPEC)
    with tempfile.TemporaryDirectory(prefix="pccm-p5ev-mutation-") as name:
        temp = Path(name)
        try:
            if scenarios is not None:
                assert scenarios != _text(saved[0]), "the mutation changed nothing"
                target = temp / saved[0].name
                target.write_text(scenarios, encoding="utf-8")
                SCENARIOS = target
            if spec_contract is not None:
                spec_dir = temp / "spec"
                shutil.copytree(saved[1], spec_dir)
                path = spec_dir / "structure_contract.yaml"
                assert spec_contract != path.read_text(encoding="utf-8"), (
                    "the mutation changed nothing"
                )
                path.write_text(spec_contract, encoding="utf-8")
                SPEC = spec_dir
            yield
        finally:
            SCENARIOS, SPEC = saved


def _refuses(control, **damage) -> str:
    with _aimed(**damage):
        try:
            control()
        except AssertionError as error:
            return str(error)
    raise AssertionError(f"the mutation survived {control.__name__}")


def test_ev_05_the_stale_global_assertion_is_refused() -> None:
    """The exact line Runtime Run 3 failed on."""
    damaged = _text(SCENARIOS).replace(
        "$null = Add-Check $list 'RunSimulation is permitted in modSimReport and nowhere else' `\n"
        "                (Test-ConstructScopedTo -Manifest $Manifest -Construct 'RunSimulation' "
        "-ModuleName 'modSimReport')",
        "$null = Add-Check $list 'RunSimulation is still forbidden in every module' `\n"
        "                (Test-ConstructForbiddenGlobally -Manifest $Manifest "
        "-Construct 'RunSimulation')", 1)
    message = _refuses(test_ev_02_every_scoped_grant_is_checked_as_a_grant,
                       scenarios=damaged)
    # Either arm is a correct refusal: the scoped grant is no longer checked,
    # and/or the contradicting global assertion is back.
    assert ("scoped grant" in message) or ("forbidden globally" in message), message


def test_ev_06_a_wrong_owner_in_the_scoped_check_is_refused() -> None:
    """Naming the wrong module would grant the endpoint somewhere it is banned."""
    damaged = _text(SCENARIOS).replace(
        "-Construct 'RunSimulation' -ModuleName 'modSimReport')",
        "-Construct 'RunSimulation' -ModuleName 'modSimNonce')", 1)
    message = _refuses(test_ev_02_every_scoped_grant_is_checked_as_a_grant,
                       scenarios=damaged)
    assert "scoped grant" in message, message


def test_ev_07_a_check_whose_wording_contradicts_its_predicate_is_refused() -> None:
    """The predicate was corrected once and the sentence beside it was not; a
    reader trusts the sentence."""
    damaged = _text(SCENARIOS).replace(
        "'RunSimulation is permitted in modSimReport and nowhere else'",
        "'RunSimulation is still forbidden in every module'", 1)
    _refuses(test_ev_02_every_scoped_grant_is_checked_as_a_grant, scenarios=damaged)


def test_ev_08_widening_the_accepted_owner_set_is_refused() -> None:
    """One construct, one owner. A second owner is a contract change."""
    contract = _text(SPEC / "structure_contract.yaml")
    anchor = ('    - construct: "RunSimulation"\n'
              '      allowed_in:\n'
              '        - "modSimReport"\n')
    assert contract.count(anchor) == 1, "the contract's grant shape moved"
    damaged = contract.replace(anchor, anchor + '        - "modSimNonce"\n', 1)
    _refuses(test_ev_01_the_contract_grants_run_simulation_to_exactly_one_owner,
             spec_contract=damaged)


def test_ev_09_dropping_the_generic_scan_for_the_scoped_checks_is_refused() -> None:
    """The scoped checks read the manifest; the generic scan reads the project."""
    damaged = _text(SCENARIOS).replace(
        "'no forbidden construct exists in the EXECUTABLE code of the real Stage-B project'",
        "'the manifest declares its forbidden constructs'", 1)
    _refuses(test_ev_04_the_generic_module_aware_scan_stays_load_bearing,
             scenarios=damaged)


def test_ev_10_a_globally_forbidden_handler_becomes_a_grant() -> None:
    """The correction must not turn every rule into a scoped grant."""
    damaged = _text(SCENARIOS).replace(
        "(Test-ConstructForbiddenGlobally -Manifest $Manifest -Construct $handler)",
        "($true)", 1)
    _refuses(test_ev_03_a_globally_forbidden_construct_is_still_checked_globally,
             scenarios=damaged)
