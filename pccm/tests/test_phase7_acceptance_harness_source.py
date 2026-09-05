#!/usr/bin/env python3
"""P7-7: the Windows acceptance harness, proved on Linux before it is run.

WHAT CAN AND CANNOT BE PROVED HERE. There is no PowerShell and no Excel on this
side, so nothing in this file claims the harness RAN. What it proves is
everything a reader would otherwise have to take on trust before spending a
Windows session on it: that the matrix is the authorised one, that each scenario
reaches the property it claims through the path it claims, that every address
comes from the projection rather than from a letter typed into the script, and
that the fixtures the two inherited-behaviour scenarios use are genuinely
independent in the dimension each is for.

THE HARNESS IS NOT PRODUCTION. A defect here wastes a Windows session; a defect
in production is shipped. So these controls are proportionate: they check the
material flow properties and the authority boundaries, and they do not attempt
to be a PowerShell type checker.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import load_sim_contract  # noqa: E402
from pccm_builder.phase7_acceptance import (  # noqa: E402
    ALLOWED_CASE_KEYS,
    ALLOWED_INSPECTION_KEYS,
    CASES_FILENAME,
    INSPECTION_FILENAME,
)

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
HARNESS = WINDOWS / "phase7_acceptance_scenarios.ps1"
TIMING = WINDOWS / "phase7_timing_scenarios.ps1"
BUILD = PCCM_ROOT / "build"
SPEC = PCCM_ROOT / "spec"

_CACHE: dict = {}

SCENARIO_IDS = ("W1", "W2", "W3", "W4", "W5", "W6", "W7", "W8")


def _text() -> str:
    if "text" not in _CACHE:
        _CACHE["text"] = HARNESS.read_text(encoding="utf-8")
    return _CACHE["text"]


def _executable() -> str:
    """The harness with comments and here-doc prose removed.

    Every claim about what the harness DOES has to be made about code, not about
    the paragraphs explaining it - a rule stated in a comment would otherwise
    satisfy a test that the code violates.
    """
    if "exec" not in _CACHE:
        # THE BLOCK COMMENT GOES FIRST. The .SYNOPSIS header explains what the
        # harness refuses to do, naming the very procedures the controls below
        # forbid - so a stripper that left it in would convict the file of its
        # own documentation.
        body = re.sub(r"^<#.*?#>\n", "", _text(), flags=re.S)
        lines = []
        for line in body.splitlines():
            if line.strip().startswith("#"):
                continue
            lines.append(re.sub(r"\s+#(?![^'\"]*['\"]).*$", "", line))
        _CACHE["exec"] = "\n".join(lines)
    return _CACHE["exec"]


def _function(name: str) -> str:
    match = re.search(rf"^function {re.escape(name)} \{{(.*?)^\}}", _text(), re.M | re.S)
    assert match, f"{name} is not defined in the harness"
    return match.group(1)


def _inspection() -> dict:
    if "inspection" not in _CACHE:
        _CACHE["inspection"] = json.loads(
            (BUILD / INSPECTION_FILENAME).read_text(encoding="utf-8"))
    return _CACHE["inspection"]


def _cases() -> dict:
    if "cases" not in _CACHE:
        _CACHE["cases"] = json.loads((BUILD / CASES_FILENAME).read_text(encoding="utf-8"))
    return _CACHE["cases"]


def _case(scenario: str) -> dict:
    for entry in _cases()["scenarios"]:
        if entry["id"] == scenario:
            return entry
    raise AssertionError(f"{scenario} is not in {CASES_FILENAME}")


# ===========================================================================
# A. THE MATRIX IS THE AUTHORISED ONE
# ===========================================================================
def test_01_the_harness_declares_exactly_the_authorised_eight() -> None:
    declared = re.findall(r"Id = '(W\d)'", _text())
    assert declared == list(SCENARIO_IDS), declared
    validated = re.search(r"ValidateSet\((.*?)\)", _text(), re.S).group(1)
    assert re.findall(r"'(W\d)'", validated) == list(SCENARIO_IDS)
    # NO 'All'. The step-at-a-time discipline is structural, not a convention:
    # there is no way to ask this script for the whole matrix in one command.
    assert "'All'" not in validated, "the harness offers a run-everything mode"


def test_02_exactly_one_scenario_runs_per_invocation() -> None:
    assert "Mandatory = $true" in _text(), "a scenario must be named explicitly"
    switch = re.search(r"switch \(\$Scenario\) \{(.*?)^\    \}", _executable(), re.M | re.S)
    assert switch, "the dispatch is not a switch on the requested scenario"
    arms = re.findall(r"^\s+'(W\d)' \{", switch.group(1), re.M)
    assert sorted(arms) == sorted(SCENARIO_IDS), arms


def test_03_every_scenario_declares_its_prerequisites() -> None:
    """The chain is DATA, so the report can state what a run had to establish."""
    expected = {"W1": [], "W2": [], "W3": [], "W4": [], "W5": ["W4"],
                "W6": ["W4", "W5"], "W7": [], "W8": ["W4", "W5"]}
    for scenario in SCENARIO_IDS:
        block = re.search(rf"Id = '{scenario}'; Prerequisites = @\((.*?)\)", _text())
        assert block, scenario
        found = re.findall(r"'(W\d)'", block.group(1))
        assert found == expected[scenario], (scenario, found)
    # AND THE DISPATCH ACTUALLY PERFORMS THEM. A declared prerequisite that no
    # arm establishes would be a promise the report makes and the run does not.
    switch = re.search(r"switch \(\$Scenario\) \{(.*?)^\    \}", _executable(), re.M | re.S).group(1)
    for scenario, chain in expected.items():
        arm = re.search(rf"'{scenario}' \{{(.*?)\n\s+'W\d' \{{|'{scenario}' \{{(.*)",
                        switch, re.S)
        assert arm, scenario
        text = arm.group(1) or arm.group(2)
        for needed in chain:
            assert f"Invoke-P7{needed}" in text, (scenario, needed)


def test_04_the_compile_check_precedes_every_behavioural_scenario() -> None:
    """W1 must be first, and that is enforced structurally rather than asked for.

    A project that does not compile cannot produce behavioural evidence, so the
    compile is re-checked at the start of every other scenario and a failure
    throws before any fixture is built.
    """
    body = _executable()
    guard = body.index("if ($Scenario -ne 'W1')")
    dispatch = body.index("switch ($Scenario)")
    assert guard < dispatch, "the compile guard runs after the scenario dispatch"
    window = body[guard:dispatch]
    assert "Invoke-P7Compile" in window
    assert "throw" in window, "a compile failure does not stop the run"


# ===========================================================================
# B. THE AUTHORITY BOUNDARIES
# ===========================================================================
def test_05_no_gate_b_scenario_is_invoked() -> None:
    """Dot-sourcing the accepted files reuses their PRIMITIVES. Running their
    matrices would produce a Gate-B result this session has no authority to
    record."""
    for banned in ("Invoke-Phase5GateBScenarios", "Invoke-Phase6GateBScenarios",
                   "Add-Phase6Result", "Reset-Phase6ResultLedger"):
        assert banned not in _executable(), f"the harness calls {banned}"


def test_06_only_definition_only_files_are_dot_sourced() -> None:
    """`phase4_functional_test.ps1` and the timing harness both have executable
    top-level code; dot-sourcing either would RUN it."""
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([\w.]+)'\)", _executable(), re.M)
    assert sorted(sourced) == ["com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1",
                               "phase6_gate_b_scenarios.ps1"], sorted(sourced)


def test_07_the_borrowed_helpers_are_byte_identical_to_the_accepted_ones() -> None:
    """Copied because they cannot be dot-sourced - so they must be COPIES, not
    a second implementation that is free to diverge."""
    timing = TIMING.read_text(encoding="utf-8")
    borrowed = ("Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
                "Set-TableCell", "Get-Phase7SourceRevision", "Get-Phase7ModuleIdentities")
    for name in borrowed:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        assert theirs, f"{name} is no longer in the accepted timing harness"
        assert _function(name) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")


def test_08_the_harness_writes_no_production_authority() -> None:
    """It reads published cells and calls published endpoints. The only things
    it writes are accepted INPUTS: the Setup controls and a register cell."""
    body = _executable()
    for banned in ("Set-SimRawCell", "Set-SimField", "Set-SimPending",
                   "Set-CalcScalar", "VBComponents.Import", "\\.Save\\(",
                   "CommandBars"):
        assert not re.search(banned, body), f"the harness reaches {banned}"
    # The workbook is closed without saving, always.
    assert "$wb.Close($false)" in body


def test_09_no_worksheet_address_is_typed_into_the_harness() -> None:
    """Every address comes from a projection. A letter typed here is a second
    authority for a coordinate, which is the defect that put the sensitivity
    availability formula on the statistics band."""
    body = _executable()
    # THE CLAIM IS ABOUT WHAT REACHES A CELL, not about every short string: 'W1'
    # and 'A2' are a scenario id and a run label, and a rule that convicted them
    # would be a rule nobody could satisfy. So the arguments of every Range call
    # are what is inspected, and a quoted address among them is the defect.
    calls = re.findall(r"\.Range\(([^)]*)\)", body)
    assert calls, "the harness makes no Range call at all"
    literals = []
    for call in calls:
        literals += re.findall(r"'([A-Z]{1,3}\d{1,4}(?::[A-Z]{1,3}\d{1,4})?)'", call)
    assert literals == [], f"A1-style address literal(s) inside a Range call: {literals}"
    # And every Range call BUILDS its address rather than naming one: the
    # argument is a variable or an expression over variables, never a constant.
    for call in calls:
        assert "$" in call, f"a Range call takes a constant address: {call}"


def test_10_no_annual_column_letter_or_stamp_row_is_restated() -> None:
    annual = _inspection()["annual_records"]
    letters = set()
    for group in ("index_columns", "quantile_first_column", "selected_px_profile_columns"):
        for bank in annual[group].values():
            letters |= {str(value) for value in bank.values()}
    letters |= {str(value) for value in annual["stamp"]["bank_value_columns"].values()}
    body = _executable()
    for letter in sorted(letters):
        assert f"'{letter}'" not in body, f"the harness restates the column {letter}"
    for row in annual["stamp"]["rows"].values():
        assert not re.search(rf"\bstamp\w*\s*\+\s*'?{row}\b", body)


def test_11_the_handoff_state_words_come_from_the_projection() -> None:
    """A state word typed here would be a second vocabulary, free to drift from
    the contract's."""
    states = set(_inspection()["handoff"]["distribution_states"]) | \
        set(_inspection()["handoff"]["profile_states"])
    body = _executable()
    # 'CURRENT' IS ALSO A PHASE-6 SIMULATION STATUS AND A PHASE-5 CALCULATION
    # STATUS, and the harness compares against those by name because they belong
    # to other contracts. So the claim is narrowed to what it is actually about:
    # no line that reads an ANNUAL handoff accessor may compare it to a state
    # word typed here.
    for line in body.splitlines():
        if not re.search(r"PCCM_Annual\w*State", line):
            continue
        for state in sorted(states):
            assert f"'{state}'" not in line, (
                f"an annual state is spelled beside the accessor: {line.strip()}")
    assert "$P7.handoff.distribution_states" in body
    assert "$P7.handoff.inconsistent_stamp_state" in body
    # AND THE STATES ARE ONLY EVER TAKEN FROM THE PROJECTION.
    assert body.count("$P7.handoff.distribution_states[") >= 3


def test_12_the_endpoint_and_accessor_names_come_from_the_projection() -> None:
    body = _executable()
    assert "$P7.command_surface.annual_endpoint" in body
    assert "$P7.command_surface.handoff_accessors" in body
    # AND THE ENDPOINT NAME IS NEVER SPELLED, anywhere in executable code -
    # including W1's own surface map, which reads it from the projection too.
    spelled = [line for line in body.splitlines()
               if "'PCCM_RunAnnualStochastic'" in line]
    assert spelled == [], spelled


# ===========================================================================
# C. THE TWO INHERITED-BEHAVIOUR SCENARIOS
# ===========================================================================
def test_13_w2_and_w3_are_independent_in_their_own_dimension() -> None:
    """The whole justification for two scenarios instead of one 300 x 200 case.

    W2 must carry many more drivers and few years; W3 many more years and few
    drivers. If either were large in both, one of them would be the Cartesian
    case the minimum matrix exists to avoid.
    """
    w2, w3 = _case("W2"), _case("W3")
    w2_drivers = len(w2["model"]["cost_lines"]) + len(w2["model"]["risks"])
    w3_drivers = len(w3["model"]["cost_lines"]) + len(w3["model"]["risks"])
    w2_years = w2["model"]["timeline"]["duration"]
    w3_years = w3["model"]["timeline"]["duration"]
    assert w2["dimension"] == "driver_count"
    assert w3["dimension"] == "year_count"
    assert w2_drivers >= 300 and w2_years <= 5, (w2_drivers, w2_years)
    assert w3_years == 200 and w3_drivers <= 10, (w3_drivers, w3_years)
    # NEITHER IS LARGE IN BOTH.
    assert w2_drivers * w2_years < w2_drivers * w3_years
    assert w3_drivers * w3_years < w2_drivers * w3_years


def test_13b_the_emitter_itself_refuses_a_cartesian_case() -> None:
    """THE ARTEFACT IS NOT THE ONLY PLACE THE RULE CAN LIVE.

    test_13 reads the BUILT corpus, which is right - it is what the harness
    consumes - but it cannot see a change to the generator until someone
    rebuilds. So the independence is also a precondition of emitting at all, and
    that is asserted here by regenerating from the live contracts and by
    planting the collapse the rule exists to refuse.
    """
    from pccm_builder import load_calc_contract
    from pccm_builder import phase7_acceptance as emitter

    calc = load_calc_contract(SPEC / "calc_contract.yaml")
    sim = load_sim_contract(SPEC / "sim_contract.yaml")
    regenerated = emitter.build_phase7_cases(calc, sim, 200)
    built = {entry["id"]: entry for entry in regenerated["scenarios"]}
    for scenario in ("W2", "W3", "W4", "W7"):
        assert built[scenario]["model"] == _case(scenario)["model"], (
            f"{scenario}: the built corpus is not what the generator now produces")

    # AND THE GENERATOR REFUSES THE COLLAPSE. A W3 grown to W2's driver count is
    # the Cartesian case the minimum matrix exists to avoid, and it must fail at
    # emission rather than quietly become the corpus.
    original = emitter._model
    try:
        def collapsed(driver_count, duration, **rest):
            if duration == 200:
                driver_count = 300
            return original(driver_count, duration, **rest)

        emitter._model = collapsed
        with pytest.raises(ValueError, match="more drivers"):
            emitter.build_phase7_cases(calc, sim, 200)
    finally:
        emitter._model = original


def test_14_w3_sits_at_the_structural_year_maximum() -> None:
    from pccm_builder import load_structure_contract

    structure = load_structure_contract(SPEC / "structure_contract.yaml")
    assert _case("W3")["model"]["timeline"]["duration"] == \
        structure.limits.max_generated_year_columns


def test_15_the_expectation_is_the_accepted_pre_phase_7_oracle() -> None:
    provenance = _cases()["provenance"]
    assert provenance["expectation_source"] == "pccm_builder.calc_oracle.calculate"
    assert "pre-Phase-7" in provenance["expectation_authority"]
    # AND IT IS NOT A PHASE-7 RECALCULATION: the harness compares the live
    # workbook against the corpus, never against a second live calculation.
    body = _executable()
    assert "$Case.expected" in body
    assert body.count("PCCM_Calculate") >= 1
    assert "Compare-P7CalcTable" in body


def test_16_the_fixture_and_the_expectation_are_the_same_model() -> None:
    """The model is read from the artefact, not built in PowerShell. Two copies
    would be free to drift and the comparison would stop meaning anything."""
    body = _function("Invoke-P7InheritedScenario")
    assert "-Model $model" in body
    assert "$model = $Case.model" in body
    # No driver is constructed here.
    for banned in ("permanent_id =", "profile_weights =", "New-Object PSObject"):
        assert banned not in body, f"the harness builds a model of its own: {banned}"


def test_17_the_comparison_allowance_is_the_projects_own() -> None:
    from pccm_builder import load_calc_contract

    calc = load_calc_contract(SPEC / "calc_contract.yaml")
    assert _cases()["provenance"]["comparison_absolute_floor"] == \
        calc.tolerances.identity_absolute_floor
    body = _executable()
    assert "$Cases.provenance.comparison_absolute_floor" in body
    # NO NEW BROAD TOLERANCE. The only other numeric allowance in the harness is
    # the I3c/I4c shape used for the profile reconciliation, and it is built from
    # the accepted floor and coefficient.
    invented = re.findall(r"-gt\s+([0-9.]+e-?\d+)", body)
    assert invented == [], f"an invented tolerance: {invented}"
    assert "1e-6" in body and "1e-12" in body


# ===========================================================================
# D. THE BEHAVIOURAL FLOW PROPERTIES
# ===========================================================================
def test_18_w4_tests_the_refusal_before_it_creates_the_run() -> None:
    """Order is the whole point: the no-run refusal can only be observed while
    no successful simulation exists."""
    body = _function("Invoke-P7W4")
    refusal = body.index("REFUSES with no successful simulation")
    run = body.index("Invoke-P7RunSimulation")
    assert refusal < run, "the simulation is created before the refusal is tested"


def test_19_w4_uses_a_fixed_seed_and_binds_to_it() -> None:
    case = _case("W4")
    assert case["seed_mode"] == "FIXED"
    assert isinstance(case["supplied_seed"], int)
    assert case["iterations"] >= 1000
    body = _function("Invoke-P7W4")
    assert "$Case.supplied_seed" in body
    assert "the FIXED effective seed is the requested one" in body


def test_20_w6_moves_the_selector_and_never_re_simulates() -> None:
    """The mandatory Phase-7 behavioural claim, and the thing that would make it
    vacuous is a hidden re-run between the two observations."""
    body = _function("Invoke-P7W6")
    assert "Set-P7NamedText" in body, "the selector is not moved through its own cell"
    assert "Invoke-P7RunSimulation" not in body
    assert "PCCM_RunSimulation" not in body
    # It DOES re-run the annual step, which is the second half of the scenario.
    assert body.count("Invoke-P7Annual") == 1
    for claim in ("the persisted annual records are value-identical",
                  "the annual DISTRIBUTION state is still CURRENT",
                  "the annual PROFILE state is OTHER Px",
                  "NOT RELABELLED",
                  "every annual ladder value is unchanged",
                  "the selected-Px profile DID change"):
        assert claim in body, claim


def test_21_w6_would_fail_rather_than_pass_vacuously() -> None:
    """A profile identical at the two levels would satisfy every 'unchanged'
    check in the scenario, so the scenario requires it to have MOVED."""
    body = _function("Invoke-P7W6")
    assert "$profileMoved" in body
    assert "would make this scenario vacuous" in body
    case = _case("W4")
    assert case["selected_confidence_level"] != case["second_confidence_level"]


def test_22_w7_is_a_to_b_to_a_and_shrinks_on_the_reused_bank() -> None:
    body = _function("Invoke-P7W7")
    # Three runs, in order, and the third returns to the first bank.
    assert body.count("Invoke-P7RunSimulation") == 3, "W7 does not run three simulations"
    assert "the third run returned to the first bank" in body
    assert "$bankA2 -ceq $bankA1" in body
    # The shrink happens on the bank that was already populated.
    shrink = body.index("A2 prerequisite: the 4-year model was applied")
    reuse = body.index("the third run returned to the first bank")
    residue = body.index("were cleared from the reused bank")
    assert shrink < reuse < residue
    case = _case("W7")
    assert case["shrink_model"]["timeline"]["duration"] < case["model"]["timeline"]["duration"]
    assert case["model"]["timeline"]["duration"] == 20
    assert case["shrink_model"]["timeline"]["duration"] == 4


def test_23_w7_proves_isolation_in_both_directions() -> None:
    body = _function("Invoke-P7W7")
    assert "did not disturb bank " in body
    # TWO COMPARISONS, ONE PER DIRECTION: bank A after B published, and bank B
    # after A published again.
    assert body.count("Test-P7SameGrid") == 2, (
        "W7 must compare the untouched bank across BOTH publications")
    assert "did not disturb bank ' + $bankA1" in body or "$bankA1)" in body


def test_24_w8_reaches_both_states_through_accepted_input_paths() -> None:
    """No hidden-sheet corruption, and the two states are reached by genuinely
    different routes - which is what makes them two states."""
    body = _function("Invoke-P7W8")
    assert "monte_carlo_iterations.defined_name" in body, "STALE is not reached via the control"
    assert "Set-TableCell" in body, "INVALID is not reached via the register"
    for banned in ("Set-SimRawCell", "Set-SimField", "Set-Phase6CellFixture"):
        assert banned not in body, f"W8 corrupts the machine sheet with {banned}"
    # STALE leaves Phase 5 CURRENT; INVALID does not. That distinction is
    # asserted live, not assumed.
    assert "STALE was reached without disturbing Phase 5" in body
    assert "the model edit made Phase 5 non-CURRENT" in body
    assert "the annual endpoint REFUSES a " in body


def test_25_w8_restores_the_state_it_borrowed() -> None:
    body = _function("Invoke-P7W8")
    assert "$originalIterations" in body
    restore = body.index("restoring the iteration control returns the simulation to CURRENT")
    assert restore > body.index("the annual endpoint REFUSES a ")


# ===========================================================================
# E. THE INVARIANTS AND THE IDENTITY CAPTURE
# ===========================================================================
def test_26_every_annual_invocation_is_bracketed_by_the_invariant_check() -> None:
    """The annual step is not a simulation. Every scenario that invokes it must
    prove no run identity, nonce, pending marker or digest moved."""
    body = _executable()
    invocations = body.count("Invoke-P7Annual -Excel")
    checks = body.count("Add-P7InvariantChecks")
    assert checks >= 4, f"{checks} invariant brackets for {invocations} annual invocations"
    for scenario in ("Invoke-P7W4", "Invoke-P7W5", "Invoke-P7W6", "Invoke-P7W8"):
        assert "Add-P7InvariantChecks" in _function(scenario), scenario


def test_27_the_two_derived_rows_are_excluded_explicitly_and_checked_separately() -> None:
    """Excluding them silently would hide a status that changed; requiring them
    frozen would require the annual step not to check whether it may run."""
    body = _function("Get-P7RunInvariants")
    assert "$script:Phase7DerivedRows" in body
    declared = re.search(r"\$script:Phase7DerivedRows = @\((.*?)\)", _text()).group(1)
    assert sorted(re.findall(r"'(\w+)'", declared)) == \
        ["simulation_status", "status_evaluated_at"]
    assert "Add-P7StatusChecks" in _executable()
    assert "the derived simulation status is unchanged" in _function("Add-P7StatusChecks")


def test_28_the_run_records_every_required_identity() -> None:
    body = _executable()
    for required in ("git HEAD", "SHA-256", "canonicalised SHA-256",
                     "EXCEL PROCESS OWNERSHIP", "model version",
                     "sim contract version", "calc contract version"):
        assert required in _text(), required
    assert "Get-Phase7SourceRevision" in body
    assert "Get-Phase6RuntimeArtefactIdentity" in body
    assert "Get-Phase7ModuleIdentities" in body
    # AND IT REFUSES A DIRTY TREE BEFORE EXCEL IS STARTED.
    refusal = body.index("REFUSED, BEFORE EXCEL WAS STARTED")
    excel = body.index("New-Object -ComObject Excel.Application")
    assert refusal < excel


def test_29_the_phase7_projection_identity_is_not_forced_to_the_phase6_one() -> None:
    """The generated projection legitimately changed during Phase 7. The harness
    records what THIS run executes; it does not rewrite history."""
    assert "daa4d278" not in _text(), "the harness pins the Phase-6 Run-6 projection identity"
    assert "is PHASE-7 and is NOT the Phase-6 Run-6" in _text()


def test_30_the_com_lifecycle_is_the_accepted_one() -> None:
    body = _executable()
    for required in ("Get-PreExistingExcelPids", "Get-ExcelIdentity", "Wait-ExcelExit",
                     "New-ReleaseLedger", "Invoke-NamedRelease", "Invoke-EmergencyExcelCleanup"):
        assert required in body, required
    # Emergency cleanup only for a positively identified process.
    cleanup = re.search(r"Invoke-EmergencyExcelCleanup -Identity (\$\w+)", body)
    assert cleanup and cleanup.group(1) == "$excelIdentity"
    assert "Stop-Process" not in body, "the harness kills a process directly"


# ===========================================================================
# F. THE PROJECTION AND THE CORPUS
# ===========================================================================
def test_31_the_projection_is_generated_not_handwritten() -> None:
    assert not (WINDOWS / INSPECTION_FILENAME).exists()
    assert (BUILD / INSPECTION_FILENAME).exists()
    assert set(_inspection()) <= set(ALLOWED_INSPECTION_KEYS)
    assert set(_cases()) <= set(ALLOWED_CASE_KEYS)


def test_32_the_projection_matches_the_contract() -> None:
    sim = load_sim_contract(SPEC / "sim_contract.yaml")
    annual = sim.raw["sim_data"]["annual_records"]
    projected = _inspection()["annual_records"]
    assert projected["header_row"] == annual["header_row"]
    assert projected["first_record_row"] == annual["first_record_row"]
    assert projected["quantile_count"] == annual["quantile_count"]
    for group in ("index_columns", "quantile_first_column", "selected_px_profile_columns"):
        for bank, entries in annual[group].items():
            for key, letter in entries.items():
                assert projected[group][bank][key] == letter, (group, bank, key)
    for field in annual["stamp"]["fields"]:
        assert projected["stamp"]["rows"][field["key"]] == field["row"]
    assert _inspection()["handoff"] == {
        "accessors": [entry["name"] for entry in annual["handoff"]["accessors"]],
        "distribution_states": list(annual["handoff"]["distribution_states"]),
        "profile_states": list(annual["handoff"]["profile_states"]),
        "inconsistent_stamp_state": annual["handoff"]["inconsistent_stamp_state"],
    }


def test_33_the_projection_matches_the_generated_module() -> None:
    """The two projections cannot drift: if they disagreed, the harness would
    inspect a cell modSimContract does not name."""
    module = (BUILD / "vba" / "modSimContract.bas").read_text(encoding="utf-8")
    constants = dict(re.findall(r"^Public Const (\w+) As \w+ = (.*)$", module, re.M))

    def value(name: str) -> str:
        return constants[name].split("    '")[0].strip().strip('"')

    annual = _inspection()["annual_records"]
    assert int(value("SIM_ANNUAL_HEADER_ROW")) == annual["header_row"]
    assert int(value("SIM_ANNUAL_FIRST_ROW")) == annual["first_record_row"]
    assert int(value("SIM_ANNUAL_QUANTILE_COUNT")) == annual["quantile_count"]
    for bank in ("A", "B"):
        assert value(f"SIM_ANNUAL_{bank}_PROJECT_INDEX_COLUMN") == \
            annual["index_columns"][bank]["project_index"]
        assert value(f"SIM_ANNUAL_{bank}_CALENDAR_YEAR_COLUMN") == \
            annual["index_columns"][bank]["calendar_year"]
        assert value(f"SIM_ANNUAL_STAMP_COLUMN_{bank}") == \
            annual["stamp"]["bank_value_columns"][bank]
        for measure in ("NOMINAL", "PV"):
            assert value(f"SIM_ANNUAL_{bank}_{measure}_FIRST_COLUMN") == \
                annual["quantile_first_column"][bank][measure.lower()]
            assert value(f"SIM_ANNUAL_{bank}_{measure}_PROFILE_COLUMN") == \
                annual["selected_px_profile_columns"][bank][measure.lower()]
    for key, row in annual["stamp"]["rows"].items():
        assert int(value(f"SIM_ANNUAL_STAMP_ROW_{key.upper()}")) == row
    assert value("SIM_ANNUAL_PUBLISHED") == annual["stamp"]["published_marker"]


def test_34_the_command_surface_matches_production() -> None:
    surface = _inspection()["command_surface"]
    run = (PCCM_ROOT / "src" / "vba" / "modSimAnnualRun.bas").read_text(encoding="utf-8")
    store = (PCCM_ROOT / "src" / "vba" / "modSimAnnualStore.bas").read_text(encoding="utf-8")
    assert f"Public Sub {surface['annual_endpoint']}()" in run
    for accessor in surface["handoff_accessors"]:
        assert f"Public Function {accessor}(" in store, accessor


def test_35_every_scenario_model_is_a_usable_fixture() -> None:
    """A model with no Risks, a degenerate driver or weights that do not sum to
    one would exercise the wrong arm - or refuse before the scenario began."""
    for scenario in _cases()["scenarios"]:
        model = scenario["model"]
        models = [model]
        if "shrink_model" in scenario:
            models.append(scenario["shrink_model"])
        for payload in models:
            where = f"{scenario['id']}"
            assert payload["cost_lines"], where
            assert payload["risks"], where
            duration = payload["timeline"]["duration"]
            for driver in payload["cost_lines"] + payload["risks"]:
                assert driver["min_value"] < driver["most_likely"] < driver["max_value"], \
                    (where, driver["permanent_id"])
                assert len(driver["profile_weights"]) == duration, (where, driver["permanent_id"])
                assert abs(sum(driver["profile_weights"]) - 1.0) <= 1e-9, \
                    (where, driver["permanent_id"], sum(driver["profile_weights"]))
            for risk in payload["risks"]:
                assert 0.0 < risk["probability"] < 1.0, (where, risk["permanent_id"])
            # Every calendar year the span can require carries a rate.
            first = payload["timeline"]["base_year"] + 1
            last = payload["timeline"]["start_year"] + duration - 1
            for profile, rates in payload["inflation"].items():
                for year in range(first, last + 1):
                    assert str(year) in rates, (where, profile, year)


def test_36_the_corpus_carries_an_expectation_for_every_model() -> None:
    for scenario in _cases()["scenarios"]:
        expected = scenario["expected"]
        model = scenario["model"]
        assert len(expected["drivers"]) == len(model["cost_lines"]) + len(model["risks"])
        assert len(expected["calc_years"]) == model["timeline"]["duration"]
        assert len(expected["annual"]) == model["timeline"]["duration"]
        assert set(expected["totals"]) == {"a_nom", "a_pv", "b_nom", "b_pv", "c_nom",
                                           "c_pv", "d_nom", "d_pv", "e_nom", "e_pv"}
        # NOT DEGENERATE. A model whose every total is zero would compare equal
        # to almost anything.
        assert abs(expected["totals"]["e_nom"]) > 1.0, scenario["id"]
        assert abs(expected["totals"]["e_pv"]) > 1.0, scenario["id"]


def test_37_the_harness_reads_the_projection_and_the_corpus_it_needs() -> None:
    body = _executable()
    for required in (INSPECTION_FILENAME, CASES_FILENAME,
                     "phase5_gate_b_inspection.json", "phase6_gate_b_inspection.json",
                     "stage_b_manifest.json"):
        assert required in body, required
    # And it refuses to start without them.
    assert "Run the Stage-A build first" in _text()


# ===========================================================================
# G. WINDOWS POWERSHELL 5.1 COMPATIBILITY
# ===========================================================================
# THE ACCEPTANCE MACHINE RUNS WINDOWS POWERSHELL 5.1, and the first W1 attempt
# never reached Excel: `Join-Path $scriptDir '..' '..' '..'` - a child per
# positional argument - is PowerShell 6+ only, and 5.1 refuses the third
# argument outright with "A positional parameter cannot be found that accepts
# argument '..'". No Excel started, so no runtime evidence was produced and
# nothing was learned except that the harness could not run.
#
# A Linux test suite cannot execute PowerShell, so it cannot prove the script
# runs. What it CAN do is refuse the constructs that are known not to exist in
# 5.1 - which is exactly the class the failure belonged to.
#
# THE SCAN COVERS THE WHOLE SUBTREE, not just the new file. The other six are
# clean today and have run on 5.1; a control scoped to one file would let the
# next edit of any of them reintroduce the same defect.
PS51_ONLY_CONSTRUCTS = (
    ("ternary ?: (PS7+)", r"[^'\"]\s\?\s[^?]"),
    ("null-coalescing ?? or ??= (PS7+)", r"\?\?"),
    ("null-conditional ?. (PS7+)", r"\$\{?\w+\}?\?\."),
    ("pipeline chain && or || (PS7+)", r"(?<![&|])(&&|\|\|)(?![&|])"),
    ("ForEach-Object -Parallel (PS7+)", r"-Parallel\b"),
    ("Split-Path -LeafBase (PS6+)", r"-LeafBase\b"),
    ("ConvertFrom-Json -AsHashtable (PS6+)", r"-AsHashtable\b"),
    ("ConvertFrom-Json -Depth (PS6+)", r"ConvertFrom-Json[^\n]*-Depth\b"),
    ("utf8NoBOM / utf8BOM encoding (PS6+)", r"utf8(No)?BOM"),
    ("Get-Content -AsByteStream (PS6+)", r"-AsByteStream\b"),
    ("Test-Json (PS6+)", r"\bTest-Json\b"),
    ("Get-Error (PS7+)", r"\bGet-Error\b"),
    ("Start-ThreadJob (PS6+)", r"\bStart-ThreadJob\b"),
    ("$IsWindows / $IsLinux / $IsMacOS (PS6+)", r"\$Is(Windows|Linux|MacOS)\b"),
    ("Where-Object -Not (PS6+)", r"Where-Object\s+-Not\b"),
    ("Sort-Object -Stable (PS6+)", r"Sort-Object[^\n]*-Stable\b"),
    ("Get-ChildItem -FollowSymlink (PS6+)", r"-FollowSymlink\b"),
    ("Copy-Item -FromSession (PS6+)", r"-FromSession\b"),
    ("$PSStyle (PS7+)", r"\$PSStyle\b"),
    ("[System.IO.Path]::Join (.NET Core only)", r"\[System\.IO\.Path\]::Join\b"),
    ("StringSplitOptions.TrimEntries (.NET Core only)", r"TrimEntries"),
)


def _powershell_files() -> list[Path]:
    return sorted(WINDOWS.glob("*.ps1"))


def _ps_code(path: Path) -> str:
    """The script with its block comment and line comments removed.

    The .SYNOPSIS header of the acceptance harness EXPLAINS the 5.1 rule and
    names the construct it refuses, so a scan that read the prose would convict
    the file of documenting itself.
    """
    body = re.sub(r"^<#.*?#>\n", "", path.read_text(encoding="utf-8"), flags=re.S)
    return "\n".join(line for line in body.splitlines()
                      if not line.strip().startswith("#"))


def _join_path_positional_count(line: str) -> int | None:
    """How many POSITIONAL arguments a Join-Path call on this line takes.

    Tokenised with parenthesis and quote awareness on purpose: a whitespace
    split calls `Join-Path $a (Split-Path -Leaf $b)` four arguments and would
    fail every correct call in the tree.
    """
    at = line.find("Join-Path")
    if at < 0:
        return None
    depth = 0
    quote = ""
    token = ""
    tokens: list[str] = []
    for character in line[at + len("Join-Path"):]:
        if quote:
            token += character
            if character == quote:
                quote = ""
            continue
        if character in "'\"":
            quote = character
            token += character
            continue
        if character in "([{":
            depth += 1
            token += character
            continue
        if character in ")]}":
            if depth == 0:
                break
            depth -= 1
            token += character
            continue
        if character.isspace() and depth == 0:
            if token:
                tokens.append(token)
                token = ""
            continue
        token += character
    if token:
        tokens.append(token)
    return len([t for t in tokens if not t.startswith("-")])


def test_39_no_join_path_takes_more_than_one_child() -> None:
    """THE EXACT DEFECT THAT STOPPED W1, refused across the whole subtree.

    Windows PowerShell 5.1's Join-Path binds -Path and -ChildPath and nothing
    else; a second child is an unbindable positional argument and the call fails
    before anything runs.
    """
    offenders: list[str] = []
    for path in _powershell_files():
        for number, line in enumerate(_ps_code(path).splitlines(), 1):
            count = _join_path_positional_count(line)
            if count is not None and count > 2:
                offenders.append(f"{path.name}:{number}: {line.strip()[:90]}")
    assert not offenders, (
        "Join-Path with more than one child is PowerShell 6+ only and fails on "
        "the 5.1 acceptance machine:\n  " + "\n  ".join(offenders))


def test_40_the_acceptance_harness_resolves_its_roots_the_accepted_way() -> None:
    """And it resolves them the way the four harnesses that have RUN on 5.1 do.

    Refusing the broken form is not enough on its own: the replacement has to be
    a form with Windows evidence behind it, not a second guess.
    """
    code = _ps_code(HARNESS)
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code
    # THE SAME TWO LINES THE ACCEPTED HARNESSES USE, character for character.
    accepted = _ps_code(TIMING)
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in accepted
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in accepted
    # And the broken construction is gone from the file entirely.
    assert "Join-Path $scriptDir '..'" not in HARNESS.read_text(encoding="utf-8")


def test_41_no_powershell_6_or_7_only_construct_is_used() -> None:
    """The class the W1 failure belonged to, refused as a class.

    A Linux suite cannot run PowerShell, so it cannot prove the script works. It
    can refuse what is known not to exist on 5.1, which is what this does.
    """
    offenders: list[str] = []
    for path in _powershell_files():
        code = _ps_code(path)
        for label, pattern in PS51_ONLY_CONSTRUCTS:
            for number, line in enumerate(code.splitlines(), 1):
                if re.search(pattern, line):
                    offenders.append(f"{path.name}:{number}: {label}: {line.strip()[:80]}")
    assert not offenders, (
        "a construct that does not exist in Windows PowerShell 5.1:\n  " +
        "\n  ".join(offenders))


def test_42_the_harness_declares_the_shell_it_is_written_for() -> None:
    """Recorded where the next person editing it will read it."""
    text = HARNESS.read_text(encoding="utf-8")
    assert "WINDOWS POWERSHELL 5.1" in text
    assert "PowerShell 6+ only" in text


# ===========================================================================
# H. THE CUSTOM-COMMAND CLOSURE
# ===========================================================================
# THE DEFECT THIS EXISTS FOR, AND WHY THE EARLIER CONTROL MISSED IT.
#
# W1 reached Excel, opened the workbook, and died before a single check with
# "The term 'Write-RowObject' is not recognized". Nothing in the harness calls
# it: `Get-Phase5TypedTableBody` does, that function lives in
# `phase5_gate_b_scenarios.ps1` - which this harness dot-sources - and its own
# helper lives in `phase4_functional_test.ps1`, which this harness deliberately
# does not. In the accepted Gate-B runs the Phase-4 driver dot-sources the
# scenarios file, so the helper is in scope; reaching the scenarios file
# directly leaves the dependency unmet.
#
# The earlier resolution check only looked at commands THIS FILE names. The
# dependency that broke W1 was transitive, so it was invisible to it. This
# control follows the calls instead: from the harness's top level and its own
# functions, INTO the dot-sourced files, transitively, and requires every custom
# command reached along the way to be defined somewhere in that closure.
#
# It found five more of the same class after Write-RowObject - Get-TableBody,
# Get-TableRowCount, Add-BlankTableRow, Remove-TableRow, Get-IdColumnValues -
# each of which would have failed a later Windows scenario one at a time. A
# helper found by running Windows and failing is a helper found too late.
#
# THREE KINDS OF COMMAND, distinguished on purpose: PowerShell built-ins, which
# need no definition; functions this harness defines; and functions reached
# through the three files it dot-sources. Anything else is undefined.
POWERSHELL_BUILTINS = frozenset({
    "Write-Host", "Write-Output", "Write-Verbose", "Write-Warning", "Write-Error",
    "Write-Debug", "Write-Progress", "Get-Content", "Set-Content", "Add-Content",
    "Out-File", "Out-Null", "Out-String", "Join-Path", "Split-Path", "Test-Path",
    "Resolve-Path", "Convert-Path", "New-Item", "Remove-Item", "Copy-Item",
    "Move-Item", "Rename-Item", "Get-Item", "Get-ChildItem", "New-Object",
    "Get-Date", "Start-Sleep", "Measure-Command", "Measure-Object",
    "ConvertFrom-Json", "ConvertTo-Json", "ConvertFrom-Csv", "ConvertTo-Csv",
    "Import-Csv", "Export-Csv", "Where-Object", "ForEach-Object", "Select-Object",
    "Sort-Object", "Group-Object", "Compare-Object", "Add-Member", "Get-Member",
    "Select-String", "Format-List", "Format-Table", "Set-StrictMode",
    "Set-Variable", "Get-Variable", "Remove-Variable", "New-Variable",
    "Get-Command", "Get-Module", "Import-Module", "Get-Process", "Stop-Process",
    "Wait-Process", "Start-Process", "Invoke-Expression", "Invoke-Command",
    "Get-Random", "New-Guid", "Get-Location", "Set-Location", "Push-Location",
    "Pop-Location", "Get-Host", "Read-Host", "Get-FileHash", "Get-Culture",
})

_VERB_NOUN = re.compile(r"(?<![\w\-.$])([A-Z][A-Za-z]*-[A-Za-z0-9]+)(?![\w\-])")


def _ps_body(text: str) -> str:
    """Comments AND string literals removed.

    A verb-noun token inside a message - 'Gate-B', 'SHA-256', 'Stage-A' - is
    prose, not a call. A scan that read those would report dozens of phantom
    commands and bury the real ones among them.
    """
    text = re.sub(r"^<#.*?#>\n", "", text, flags=re.S)
    text = "\n".join(l for l in text.splitlines() if not l.strip().startswith("#"))
    text = re.sub(r"'[^'\n]*'", "''", text)
    text = re.sub(r'"[^"\n]*"', '""', text)
    return re.sub(r"\s#[^\n]*", "", text)


def _ps_functions(path: Path) -> dict[str, str]:
    """name -> body, by brace matching rather than by a line pattern."""
    code = _ps_body(path.read_text(encoding="utf-8"))
    out: dict[str, str] = {}
    for match in re.finditer(r"^function\s+([\w-]+)\s*\{", code, re.M):
        start = match.end() - 1
        depth = 0
        for index in range(start, len(code)):
            if code[index] == "{":
                depth += 1
            elif code[index] == "}":
                depth -= 1
                if depth == 0:
                    out[match.group(1)] = code[start:index]
                    break
    return out


def _ps_top_level(path: Path) -> str:
    code = _ps_body(path.read_text(encoding="utf-8"))
    for body in _ps_functions(path).values():
        code = code.replace(body, "")
    return code


def _dot_sourced() -> list[Path]:
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([\w.]+)'\)",
                         _executable(), re.M)
    assert sourced, "the harness dot-sources nothing"
    return [WINDOWS / name for name in sourced]


def _command_closure() -> tuple[set[str], list[tuple[str, str]]]:
    defined: dict[str, tuple[str, str]] = {}
    for path in [HARNESS] + _dot_sourced():
        for name, body in _ps_functions(path).items():
            defined.setdefault(name, (path.name, body))

    seeds = [("<harness top level>", _ps_top_level(HARNESS))]
    seeds += [(name, body) for name, (owner, body) in defined.items()
              if owner == HARNESS.name]

    seen: set[str] = set()
    missing: list[tuple[str, str]] = []
    work = list(seeds)
    while work:
        where, body = work.pop()
        for call in sorted(set(_VERB_NOUN.findall(body))):
            if call in POWERSHELL_BUILTINS:
                continue
            if call not in defined:
                missing.append((where, call))
                continue
            if call in seen:
                continue
            seen.add(call)
            work.append((call, defined[call][1]))
    return seen, missing


def test_43_every_custom_command_the_harness_reaches_is_defined() -> None:
    """THE CONTROL THAT WOULD HAVE CAUGHT THE W1 FAILURE.

    Transitive, because the dependency that broke it was: the harness never
    names Write-RowObject, and would have failed on it - and then on five more
    like it - one Windows session at a time.
    """
    reached, missing = _command_closure()
    assert not missing, (
        "custom command(s) called but defined nowhere in the harness or the "
        "files it dot-sources:\n  " +
        "\n  ".join(f"{call}  (reached from {where})"
                     for where, call in sorted(set(missing))))
    # The closure is real: a harness that reached almost nothing would pass the
    # assertion above by proving nothing.
    assert len(reached) > 80, f"only {len(reached)} custom commands reached"


def test_44_the_transitively_needed_helpers_are_present_and_verbatim() -> None:
    """The six the closure found, pinned to the accepted implementation.

    None of them is called from this file - the accepted Phase-5 fixture
    choreography calls them - so a reader has no local reason to expect them and
    a control has to say why they are here.
    """
    timing = TIMING.read_text(encoding="utf-8")
    transitive = ("Write-RowObject", "Get-TableBody", "Get-TableRowCount",
                  "Add-BlankTableRow", "Remove-TableRow", "Get-IdColumnValues")
    for name in transitive:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        assert theirs, f"{name} is no longer in the accepted timing harness"
        assert _function(name) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")
        # AND IT REALLY IS TRANSITIVE: this file defines it and never calls it.
        # That is the whole reason it is easy to leave out, and the reason a
        # reader needs telling why it is here at all.
        body = _ps_body(HARNESS.read_text(encoding="utf-8"))
        for definition in _ps_functions(HARNESS).values():
            body = body.replace(definition, "")
        # The `function <name> {` headers survive the body removal; they are
        # declarations, not calls.
        body = "\n".join(line for line in body.splitlines()
                         if not line.strip().startswith("function "))
        assert not re.search(rf"(?<![\w\-.$]){name}(?![\w\-])", body), (
            f"{name} is called from the harness after all; the transitive "
            "justification in the source no longer describes it")
    # The dot-sourced set is unchanged: adding phase4_functional_test.ps1 would
    # RUN the Phase-4 matrix, which is why these are copied instead.
    assert [p.name for p in _dot_sourced()] == [
        "com_lifecycle.ps1", "phase5_gate_b_scenarios.ps1", "phase6_gate_b_scenarios.ps1"]


def test_45_the_owned_excel_is_shut_down_on_every_path() -> None:
    """A HARNESS FAILURE MUST NOT ORPHAN EXCEL, and the W1 failure was one.

    The exception was raised inside the session try, so the finally ran and the
    owned PID was closed, quit, waited for and - if it had not exited - cleaned
    up. What the report could NOT show is that this happened, because the
    shutdown wrote a line only on the emergency path. The ledger now reports on
    every path, so a future failure leaves evidence of the process's fate rather
    than silence.
    """
    body = _executable()
    session = body[body.index("$excel = New-Object -ComObject Excel.Application"):]
    finally_at = session.index("} finally {")
    shutdown = session[finally_at:]
    # The whole shutdown is itself inside a try/finally, so a failure while
    # releasing still reaches the wait and the emergency path.
    assert shutdown.count("finally") >= 2, (
        "the shutdown has no inner finally; a release that throws would skip "
        "Wait-ExcelExit and orphan the process")
    for required in ("$wb.Close($false)", "$excel.Quit()", "Wait-ExcelExit",
                     "Invoke-EmergencyExcelCleanup"):
        assert required in shutdown, required
    # AND IT REPORTS THE OUTCOME WHATEVER HAPPENED.
    assert "EXCEL SHUTDOWN:" in shutdown
    assert shutdown.count("EXCEL SHUTDOWN:") >= 2, (
        "the shutdown reports only one outcome; the report must say what became "
        "of the owned process on the normal path too")


def test_38_the_build_emits_both_artefacts() -> None:
    build = (PCCM_ROOT / "builder" / "build_stage_a.py").read_text(encoding="utf-8")
    assert "emit_phase7_acceptance" in build
    assert "max_generated_year_columns" in build
