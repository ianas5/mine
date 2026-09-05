#!/usr/bin/env python3
"""P7-7: the MINIMAL W4 runner, proved on Linux before a Windows session.

W4 ESTABLISHES THE BASELINE W5 AND W6 WILL ANALYSE, and it proves two things in
one session because the state each needs exists exactly once: the annual
endpoint REFUSES while no successful simulation exists and moves nothing, and a
FIXED-seed simulation then succeeds and publishes a CURRENT identity.

THE TWO THINGS MOST WORTH PROVING HERE ARE ORDER AND RESTRAINT. The refusal is
only evidence if it happens BEFORE the first successful run - afterwards the
endpoint would legitimately succeed - so the order is a control, not a comment.
And W4 must STOP at a successful simulation: running the annual endpoint
successfully would consume the first-annual-run state W5 exists to observe, so
that is a control too.

WHAT IS INDEPENDENT AND WHAT IS NOT is checked as carefully as the values. The
Phase-5 calculation underneath the simulation is compared against
`pccm_builder.calc_oracle.calculate`; the request fields are exact and
contract-owned; the published minimum and maximum are selections and are
compared exactly to the iteration column. There is no independent stochastic
expectation for this model - the accepted Phase-6 Gate-B oracle owns four
analytical parity plan cases at seed 12345 - so the mean and the ladder are
recorded and not compared, and a control requires the runner to say so rather
than to manufacture agreement.

WHAT THIS FILE CANNOT PROVE: there is no PowerShell and no Excel here, so
nothing below claims the runner RAN.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "tests"))
sys.path.insert(0, str(PCCM_ROOT / "builder"))

import pytest  # noqa: E402

import test_phase7_acceptance_harness_source as accepted  # noqa: E402
from pccm_builder import (  # noqa: E402
    load_calc_contract,
    load_sim_contract,
    load_structure_contract,
)

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase7_w4_base_simulation.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
W2 = WINDOWS / "phase7_w2_many_drivers.ps1"
W3 = WINDOWS / "phase7_w3_long_years.ps1"
TIMING = WINDOWS / "phase7_timing_scenarios.ps1"
LIFECYCLE = WINDOWS / "com_lifecycle.ps1"
PHASE5 = WINDOWS / "phase5_gate_b_scenarios.ps1"
PHASE6 = WINDOWS / "phase6_gate_b_scenarios.ps1"
BUILD = PCCM_ROOT / "build"

DOT_SOURCED = (LIFECYCLE, PHASE5, PHASE6)

# The ten helpers the closure below found. Named here so the count is a claim a
# reader can check rather than a number that drifts silently.
COPIED_HELPERS = (
    "Write-RowObject", "Get-NamedValue", "Set-NamedValue", "Get-TableColumnNames",
    "Set-TableCell", "Get-TableBody", "Get-TableRowCount", "Add-BlankTableRow",
    "Remove-TableRow", "Get-IdColumnValues",
)

# NINE OF THE TEN ARE PURELY TRANSITIVE - the accepted fixture calls them and
# this runner never does. `Set-NamedValue` is the exception and is named as one:
# W4 writes the simulation request through it, so it is BOTH copied for the
# fixture's sake and called directly here.
CALLED_DIRECTLY = ("Set-NamedValue",)

# W1, W2 AND W3 ARE CLOSED. All three runners are accepted Windows evidence and
# are pinned, not maintained, by this round. The digests are LITERAL - hashing a
# file at import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"
W2_SHA256 = "558bcfa0528b38c28c6e81dd726a18fed9078098f79bcabb792304e6255d8fee"
W3_SHA256 = "97af69bf836a2d798be4937f1e81cfe7218d2cfd702c1b96aa25623e7f8c12be"

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W4 refuses to do and names the things the controls
    forbid, so a scan that read the prose would convict the file of its own
    documentation.
    """
    return accepted._ps_code(RUNNER)


def _functions() -> dict[str, str]:
    """name -> body, brace-matched over the comment-stripped source.

    String literals are KEPT: the compared column names, the field map and the
    provenance strings are all literals, and they are the subject here.
    """
    if "fns" not in _CACHE:
        code = _code()
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
        _CACHE["fns"] = out
    return _CACHE["fns"]


def _function(name: str) -> str:
    body = _functions().get(name)
    assert body is not None, f"{name} is not defined in the W4 runner"
    return body


def _own_code() -> str:
    """The runner's top level plus the functions it wrote itself.

    The ten copied helpers are accepted code from another file; a claim about
    what THIS runner does must not be made about them, and a claim about them is
    made by pinning them instead.
    """
    code = _code()
    for name in COPIED_HELPERS:
        code = code.replace(_function(name), "")
    return code


def _cases() -> dict:
    if "cases" not in _CACHE:
        _CACHE["cases"] = json.loads(
            (BUILD / "phase7_acceptance_cases.json").read_text(encoding="utf-8"))
    return _CACHE["cases"]


def _w4() -> dict:
    case = [s for s in _cases()["scenarios"] if s["id"] == "W4"]
    assert case, "the acceptance corpus carries no W4 scenario"
    return case[0]


def _year_window() -> tuple[int, int, int]:
    """(min_year, max_year, max_generated_year_columns) from the structural contract.

    THE AUTHORITY, NOT A REMEMBERED NUMBER. The W4 fixture's calendar span is
    derived from these, so the assertions about it have to be too - the round
    that produced 2026-2225 failed precisely because a literal span was carried
    around instead of being derived.
    """
    text = (PCCM_ROOT / "spec" / "structure_contract.yaml").read_text(encoding="utf-8")
    values = {}
    for key in ("min_year", "max_year", "max_generated_year_columns"):
        match = re.search(rf"^\s*{key}:\s*(\d+)\s*$", text, re.M)
        assert match, f"structure_contract.yaml declares no {key}"
        values[key] = int(match.group(1))
    return values["min_year"], values["max_year"], values["max_generated_year_columns"]


def _span() -> tuple[int, int, int]:
    """(first calendar year, last calendar year, duration) of the W4 fixture."""
    timeline = _w4()["model"]["timeline"]
    first = int(timeline["start_year"])
    duration = int(timeline["duration"])
    return first, first + duration - 1, duration


def _inspection() -> dict:
    if "inspection" not in _CACHE:
        _CACHE["inspection"] = json.loads(
            (BUILD / "phase5_gate_b_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["inspection"]


# ===========================================================================
# A. SCOPE: THIS IS W4, AND W1 IS CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 1400, (
        f"{len(lines)} lines; W4 is one scenario in two parts and most of its "
        "behaviour is meant to come from the accepted Phase-5 fixture, the "
        "accepted Phase-6 state readers and the projection")
    # Most of the file is either the accepted copied helpers or W4's own small
    # comparison; neither should have grown into a second harness.
    assert len(_own_code().splitlines()) < 1000


def test_02_it_dot_sources_only_definition_only_files() -> None:
    """The three that define and run nothing. `phase4_functional_test.ps1` is
    NOT among them: dot-sourcing it would run the entire Phase-4 matrix, which
    is exactly why its helpers are copied instead."""
    sourced = re.findall(r"^\. \(Join-Path \$scriptDir '([\w.]+)'\)", _code(), re.M)
    assert sourced == [p.name for p in DOT_SOURCED], sourced
    for path in DOT_SOURCED:
        # Definition-only means the top level holds declarations and variables,
        # never a scenario invocation. The function HEADERS survive the body
        # removal - they are declarations, not calls - so they are dropped
        # before the top level is read for calls.
        top = "\n".join(line for line in accepted._ps_top_level(path).splitlines()
                        if not line.strip().startswith("function "))
        for entry in ("Invoke-Phase5GateBScenarios", "Invoke-Phase6GateBScenarios"):
            assert not re.search(rf"(?<![\w\-.$]){entry}(?![\w\-])", top), (
                f"{path.name} invokes {entry} at top level; dot-sourcing it "
                "would RUN a Gate-B scenario")


def test_03_the_stage_b_bootstrap_is_a_child_script_not_a_dot_source() -> None:
    code = _code()
    assert "& $bootstrap -BuildDir $tempRoot -Force" in code
    assert ". (Join-Path $scriptDir 'build_stage_b.ps1')" not in code


def test_04_no_sensitivity_no_later_scenario_and_no_second_stochastic_oracle() -> None:
    """W4 runs the annual endpoint ONCE, expecting a refusal, and one simulation.

    Everything else stochastic belongs to another scenario, and a second
    stochastic oracle invented here would be exactly the "compare a run against
    another run and call it independence" the authorisation refuses.
    """
    code = _own_code()
    for forbidden in ("PCCM_RunSensitivity", "Invoke-Phase5GateBScenarios",
                      "Invoke-Phase6GateBScenarios", "phase7_acceptance_scenarios",
                      "phase7_w1_smoke", "phase7_w2_many_drivers", "phase7_w3_long_years"):
        assert forbidden not in code, f"{forbidden} is outside W4"
    for scenario in ("'W2'", "'W3'", "'W5'", "'W6'", "'W7'", "'W8'"):
        assert scenario not in code, scenario
    assert "'W4'" in code
    # ONE simulation, ONE annual attempt.
    assert code.count("Invoke-Phase6Simulation") == 1, (
        "W4 runs the simulation more than once; a same-seed replay compared "
        "against itself is not independent evidence and is not authorised here")
    assert code.count("Invoke-W4Annual") == 2, (
        "the annual endpoint is invoked from exactly one call site")


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """One lightweight read-only compile trigger is authorised because Excel
    must execute the current project. The surface matrix is W1's."""
    code = _own_code()
    for forbidden in ("VBProject", "VBComponents", "CodeModule", "ProcOfLine"):
        assert forbidden not in code, f"{forbidden} is W1's evidence, not W4's"
    assert "$excel.Run('PCCM_CalculationStatus')" in code
    assert "the current VBAProject compiles in real Excel" in code


def test_06_the_accepted_w1_and_w2_runners_are_untouched() -> None:
    """BOTH ARE CLOSED on accepted Windows evidence. This round does not modify
    or rerun either, and an edit to one must fail here rather than pass."""
    assert W1.exists() and W2.exists()
    assert hashlib.sha256(W1.read_bytes()).hexdigest() == W1_SHA256, (
        "the accepted W1 runner has been modified; it is closed evidence")
    assert hashlib.sha256(W2.read_bytes()).hexdigest() == W2_SHA256, (
        "the accepted W2 runner has been modified; it is closed evidence")
    code = _code()
    assert "phase7_w1_smoke" not in code and "phase7_w2_many_drivers" not in code, (
        "W4 reaches into a closed runner at run time")


# THE PARTS W4 SHARES WITH W2, PINNED SO THE MANDATED COPY CANNOT DRIFT.
# Reusing W2's proven execution architecture was the instruction; two runners
# that started identical and quietly diverged would be the cost of it, so the
# shared functions are compared after normalising the scenario name.
SHARED_WITH_W2 = ("Write-W4Line", "Add-W4Check", "Format-W4Value", "Invoke-W4Release",
                  "Get-W4SourceRevision", "Compare-W4Cell", "Compare-W4Table")


def test_06b_the_shape_shared_with_w2_is_identical_to_w2s() -> None:
    import re as _re
    ours = _text()
    theirs = W2.read_text(encoding="utf-8")
    for name in SHARED_WITH_W2:
        mine = _re.search(rf"^function {name} \{{(.*?)^\}}", ours, _re.M | _re.S)
        assert mine, f"{name} is not defined in the W4 runner"
        w2name = name.replace("W4", "W2")
        yours = _re.search(rf"^function {w2name} \{{(.*?)^\}}", theirs, _re.M | _re.S)
        assert yours, f"{w2name} is no longer in the accepted W2 runner"
        assert mine.group(1).replace("W4", "W2") == yours.group(1), (
            f"{name} has drifted from the accepted {w2name}; the two runners are "
            "meant to share one execution architecture, not two")


# ===========================================================================
# B. THE CALL CLOSURE - THE DEFECT THAT COST W1 A WINDOWS SESSION
# ===========================================================================

def _closure() -> tuple[set[str], list[tuple[str, str]]]:
    defined: dict[str, str] = {}
    for path in (RUNNER,) + DOT_SOURCED:
        for name, body in accepted._ps_functions(path).items():
            defined.setdefault(name, body)
    seen: set[str] = set()
    missing: list[tuple[str, str]] = []
    work: list[tuple[str, str]] = [("<runner top level>", accepted._ps_top_level(RUNNER))]
    work += list(accepted._ps_functions(RUNNER).items())
    while work:
        where, body = work.pop()
        for call in sorted(set(accepted._VERB_NOUN.findall(body))):
            if call in accepted.POWERSHELL_BUILTINS:
                continue
            if call not in defined:
                missing.append((where, call))
                continue
            if call in seen:
                continue
            seen.add(call)
            work.append((call, defined[call]))
    return seen, missing


def test_07_every_custom_command_the_runner_reaches_is_defined() -> None:
    """TRANSITIVELY, because the dependency that broke W1 was transitive: the
    harness never named `Write-RowObject`, and would have failed on it - and
    then on nine more like it - one Windows session at a time."""
    reached, missing = _closure()
    assert not missing, (
        "custom command(s) called but defined neither in the runner nor in the "
        "files it dot-sources:\n  " +
        "\n  ".join(f"{call}  (reached from {where})" for where, call in sorted(set(missing))))
    assert len(reached) > 40, f"only {len(reached)} custom commands reached"


def test_08_the_copied_helpers_are_verbatim_and_genuinely_transitive() -> None:
    """Pinned to the accepted implementation, and justified: nothing in this
    runner calls any of them. The accepted Phase-5 fixture does, and its own
    file does not define them."""
    timing = TIMING.read_text(encoding="utf-8")
    raw = _text()
    own = _own_code()
    for name in COPIED_HELPERS:
        theirs = re.search(rf"^function {name} \{{(.*?)^\}}", timing, re.M | re.S)
        assert theirs, f"{name} is no longer in the accepted timing harness"
        # RAW AGAINST RAW, comments included: a copied helper whose commentary
        # was edited is no longer the accepted implementation, and the whole
        # point of copying rather than paraphrasing is that it stays identical.
        ours = re.search(rf"^function {name} \{{(.*?)^\}}", raw, re.M | re.S)
        assert ours, f"{name} is not defined in the W4 runner"
        assert ours.group(1) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")
        body = own
        for line in body.splitlines():
            if line.strip().startswith("function "):
                continue
        called = re.search(rf"(?<![\w\-.$]){name}(?![\w\-])",
                           "\n".join(l for l in own.splitlines()
                                     if not l.strip().startswith("function ")))
        if name in CALLED_DIRECTLY:
            assert called, (
                f"{name} is listed as called directly but nothing calls it; "
                "the justification in this file no longer describes the runner")
        else:
            assert not called, (
                f"{name} is called from the runner after all; the transitive "
                "justification in the source no longer describes it")


# ===========================================================================
# C. WINDOWS POWERSHELL 5.1
# ===========================================================================

def test_09_the_runner_declares_and_obeys_the_shell_it_targets() -> None:
    text = _text()
    assert "WINDOWS POWERSHELL 5.1" in text
    code = _code()
    assert "$pccmRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)" in code
    assert "$repoRoot = Split-Path -Parent $pccmRoot" in code
    for number, line in enumerate(code.splitlines(), 1):
        count = accepted._join_path_positional_count(line)
        assert count is None or count <= 2, f"line {number}: {line.strip()}"


def test_10_no_powershell_6_or_7_only_construct_is_used() -> None:
    offenders: list[str] = []
    code = _code()
    for label, pattern in accepted.PS51_ONLY_CONSTRUCTS:
        for number, line in enumerate(code.splitlines(), 1):
            if re.search(pattern, line):
                offenders.append(f"{number}: {label}: {line.strip()[:80]}")
    assert not offenders, "\n  ".join(offenders)


# ===========================================================================
# D. THE FIXTURE AND THE ORACLE
# ===========================================================================

def test_11_the_fixture_is_the_authorised_behavioural_request() -> None:
    """SMALL AND DETERMINISTIC ON PURPOSE. A behavioural baseline that had grown
    into a performance case would be a different scenario, and a seed or an
    iteration count that had drifted would break the identity W5 binds to."""
    case = _w4()
    model = case["model"]
    drivers = len(model["cost_lines"]) + len(model["risks"])
    assert case["dimension"] == "behavioural"
    assert drivers == 5, drivers
    assert model["timeline"]["duration"] == 4
    assert case["iterations"] == 1000
    assert case["seed_mode"] == "FIXED"
    assert case["supplied_seed"] == 20260905
    assert case["selected_confidence_level"] == "P80"
    # AND THE RUNNER REQUIRES ALL OF IT before Excel does any work.
    code = _own_code()
    assert "($driverCount -eq 5) -and ($yearCount -eq 4) -and ($iterations -eq 1000)" in code
    assert "(([string]$case.seed_mode) -ceq 'FIXED')" in code
    assert "($suppliedSeed -eq 20260905)" in code
    assert "the W4 fixture is the authorised size and request" in code


def test_12_the_expectation_is_the_independent_phase5_oracle() -> None:
    """NOT ANOTHER VBA CALCULATION. The corpus expectations come from the same
    independent Phase-5 implementation the accepted phase5_cases.json corpus is
    built from, which Phase 7 did not touch."""
    provenance = _cases()["provenance"]
    assert provenance["expectation_source"] == "pccm_builder.calc_oracle.calculate"
    assert "pre-Phase-7 authority" in provenance["expectation_authority"]
    # And the runner checks it at run time rather than trusting the file name.
    code = _own_code()
    assert "'pccm_builder.calc_oracle.calculate'" in code
    assert "the expectations come from the independent Phase-5 oracle" in code


def test_13_the_model_and_the_expectation_are_the_same_artefact() -> None:
    """The runner reads BOTH from the corpus. A fixture built here and an
    expectation read from there would compare two different models."""
    code = _own_code()
    assert "$model = $case.model" in code
    assert "-Model $model" in code
    assert "$case.expected" in code
    assert code.count("phase7_acceptance_cases.json") >= 1
    # There is no second source of model data anywhere in the runner.
    assert "cost_lines = @(" not in code
    assert "New-Object PSObject" not in code


def test_14_every_supplied_driver_has_exactly_one_expected_record() -> None:
    case = _w4()
    model, expected = case["model"], case["expected"]
    supplied = [line["permanent_id"] for line in model["cost_lines"]]
    supplied += [risk["permanent_id"] for risk in model["risks"]]
    published = [record["permanent_id"] for record in expected["drivers"]]
    assert published == supplied, "the oracle's driver order is not the supply order"
    assert len(set(published)) == len(published), "duplicate ids in the corpus"
    kinds = {record["driver_kind"] for record in expected["drivers"]}
    assert kinds == {"Cost Line", "Risk"}


def test_15_the_oracle_carries_every_projected_column_of_every_table() -> None:
    """This is what makes "all the other already-contracted fields" a fact
    rather than an intention: the runner compares every projected column, so a
    column the oracle did not carry would be a failure it could not evaluate."""
    tables = _inspection()["calc"]["tables"]
    expected = _w4()["expected"]
    pairs = (("calc_drivers", expected["drivers"], None),
             ("calc_years", expected["calc_years"], None),
             ("calc_annual", expected["annual"], None),
             ("calc_inflation_factors", expected["inflation_factors"],
              {"inflation_profile": "profile",
               "cumulative_inflation_factor": "cumulative_factor"}))
    for key, rows, field_map in pairs:
        columns = tables[key]["columns"]
        assert rows, key
        for column in columns:
            field = (field_map or {}).get(column, column)
            assert field in rows[0], f"{key}: the oracle carries no {field} for column {column}"
    # The driver table really is the wide one: 21 contracted columns, all compared.
    assert len(tables["calc_drivers"]["columns"]) == 21


def test_16_the_corpus_axis_is_complete_for_the_behavioural_model() -> None:
    """The deterministic base W4 samples around has to be whole before any
    stochastic observation on top of it means anything."""
    case = _w4()
    years = case["model"]["timeline"]["duration"]
    expected = case["expected"]
    assert len(expected["calc_years"]) == years
    assert len(expected["annual"]) == years
    assert len(expected["discount_factors"]) == years
    profiles = len(case["model"]["inflation"])
    assert len(expected["inflation_factors"]) == profiles * years
    assert {len(record["weights"]) for record in expected["drivers"]} == {years}
    indices = [row["project_index"] for row in expected["calc_years"]]
    assert indices == list(range(1, years + 1))


def test_17_the_allowance_is_the_accepted_phase5_floor_and_nothing_else() -> None:
    """No tolerance is invented here. The corpus carries the project's own
    identity absolute floor, straight from the calc contract."""
    contract = load_calc_contract(PCCM_ROOT / "spec" / "calc_contract.yaml")
    floor = float(contract.tolerances.identity_absolute_floor)
    assert float(_cases()["provenance"]["comparison_absolute_floor"]) == floor
    assert floor == 1e-6
    code = _own_code()
    assert "$allowance = [double]$cases.provenance.comparison_absolute_floor" in code
    # AND THE RUNNER NAMES NO TOLERANCE OF ITS OWN.
    literals = re.findall(r"-Allowance\s+([^\s`]+)", code)
    assert literals, "nothing is compared with an allowance at all"
    for literal in literals:
        # `$Allowance` is the parameter Compare-W4Table forwards; `$allowance`
        # is the value read from the corpus. Anything else is a new tolerance.
        assert literal in ("$allowance", "$Allowance"), (
            f"a second tolerance appears: {literal}")
    assert not re.search(r"\b\d*\.?\d+e-\d+\b", code), (
        "a numeric tolerance literal appears in the runner")


# ===========================================================================
# E. THE COMPARISON
# ===========================================================================

def test_18_one_rule_chosen_by_the_oracle_not_by_a_list_of_column_names() -> None:
    """A hand-written column list is how "all the other contracted fields"
    quietly stops being true. The rule is keyed on what the ORACLE holds, so a
    column the contract adds is compared the moment the oracle carries it."""
    cell = _function("Compare-W4Cell")
    assert "$null -eq $Expect" in cell, "a missing oracle value has no rule"
    assert "$Expect -is [string]" in cell
    assert "$Got -isnot [double]" in cell, (
        "a value published as text would be coerced instead of failing")
    table = _function("Compare-W4Table")
    assert "$columns = @($Inspection.calc.tables.$TableKey.columns)" in table, (
        "the compared columns are not read from the projection")
    assert "foreach ($column in $columns)" in table
    # No per-table column list is passed in anywhere.
    code = _own_code()
    for banned in ("-Numeric @(", "-Exact @("):
        assert banned not in code, f"{banned} reintroduces a hand-written column list"


def test_19_a_fabricated_zero_and_a_stringified_number_both_fail() -> None:
    """The two failures that matter most, and the two a lenient comparison
    hides: a published value where the model has no risk at all, and a number
    Excel wrote as text."""
    cell = _function("Compare-W4Cell")
    assert "where the oracle has none" in cell
    assert "is not a number" in cell
    # The reader is the accepted TYPED one, which is what makes the second
    # possible at all.
    code = _own_code()
    assert "Get-CalcTableRows" in code
    typed = re.search(r"^function Get-CalcTableRows \{(.*?)^\}",
                      PHASE5.read_text(encoding="utf-8"), re.M | re.S)
    assert typed and "Get-Phase5TypedTableBody" in typed.group(1), (
        "the accepted reader is no longer the typed one")


def test_20_the_inflation_field_map_is_pinned_to_both_sides() -> None:
    """The corpus names two of these fields differently from the projected
    columns. The map is the only place that is true, so it is checked against
    both the projection and the corpus rather than assumed."""
    code = _own_code()
    match = re.search(r"\$inflationMap = @\{([^}]*)\}", code)
    assert match, "the field map is gone"
    pairs = dict(re.findall(r"'([\w]+)' = '([\w]+)'", match.group(1)))
    assert pairs == {"inflation_profile": "profile",
                     "cumulative_inflation_factor": "cumulative_factor"}
    columns = _inspection()["calc"]["tables"]["calc_inflation_factors"]["columns"]
    row = _w4()["expected"]["inflation_factors"][0]
    for column, field in pairs.items():
        assert column in columns, column
        assert field in row, field


# ===========================================================================
# F. THE DRIVER-COUNT EVIDENCE
# ===========================================================================

def test_21_the_deterministic_base_is_compared_on_every_surface() -> None:
    """THE SAME COMPARISON W2 AND W3 MAKE, on a five-driver, four-year model.

    The simulation samples around this calculation, so a wrong base would make
    every stochastic observation below meaningless.
    """
    code = _own_code()
    for key in ("'calc_drivers'", "'calc_years'", "'calc_inflation_factors'", "'calc_annual'"):
        assert key in code, key
    assert "publishes one row per expected row" in code
    assert "the ten published totals match the accepted Phase-5 oracle" in code
    assert "the applied timeline is the one the oracle computed against" in code
    assert "Compare-W4Table -Live $live" in code


def test_22_the_deterministic_anchor_inside_the_simulation_block_is_compared() -> None:
    """THE ONE NUMBER IN THE SIMULATION BLOCK A PRE-PHASE-7 AUTHORITY OWNS.

    `deterministic_base_a` is published beside the summary ladder, and the
    oracle's A totals are its expectation - which ties the simulation block to
    something other than itself.
    """
    code = _own_code()
    assert "-RowKey 'deterministic_base_a'" in code
    assert "'a_nom'" in code and "'a_pv'" in code
    assert "matches the independent Phase-5 oracle" in code
    projection = _inspection()
    assert "deterministic_base_a" in json.loads(
        (BUILD / "phase6_gate_b_inspection.json").read_text(encoding="utf-8")
    )["sim_data"]["summary_statistics"]["rows"], (
        "the projection no longer publishes a deterministic base beside the "
        "summary; the anchor this control relies on is gone")
    assert projection is not None


def test_24_the_calculation_runs_through_its_own_accepted_path() -> None:
    """Through production's own operation and production's own status, never by
    the runner writing a cell or judging currency for itself."""
    code = _own_code()
    assert "Invoke-Phase5ProductionOperation -Excel $excel" in code
    assert "-Operation 'PCCM_Calculate'" in code
    assert "the deterministic calculation is still CURRENT after the refusal" in code
    assert "the simulation reports CURRENT" in code
    assert "Set-Phase5Fixture -Excel $excel" in code
    assert "Save-Phase5LockedFxSeed" in code
    # THE RUNNER WRITES NO MODEL DATA ITSELF. Every register row, grid cell and
    # timeline comes from the accepted fixture driving production's own Add
    # endpoints. What W4 does write is the simulation REQUEST - three projected
    # defined names - and it writes them through the projection, never by
    # address.
    outside = _own_code().replace(_function("Set-W4NamedText"), "")
    for banned in ("$rng.Value2 =", ".ListObjects", "$ws.Cells", "Set-SimRawCell",
                   "Set-SimField", "Set-TableCell -"):
        assert banned not in outside, banned
    code = _own_code()
    written = re.findall(r"-DefinedName \(\[string\]([^)]+)\)", code)
    assert sorted(written) == sorted([
        "$simInspection.controls.monte_carlo_iterations.defined_name",
        "$simInspection.controls.random_seed.defined_name",
        "$inspection.inputs.selected_confidence_level.defined_name",
    ]), written


# ===========================================================================
# G. COM LIFECYCLE - THE DISCIPLINE W1 CLOSED ON
# ===========================================================================

def test_25_every_acquisition_is_counted_and_released_into_one_ledger() -> None:
    code = _own_code()
    # RELEASE-TRANSIENT IS ALLOWED IN EXACTLY ONE PLACE, and it is named.
    # `Set-W4NamedText` writes one defined name: its three objects are acquired
    # and released inside a single statement, in the same convention every
    # accepted reader uses, and they are covered by the accepted-reader
    # transient gate. Everything the runner HOLDS goes through its own ledger.
    outside = code.replace(_function("Set-W4NamedText"), "")
    assert "Release-Transient" not in outside, (
        "a release the runner holds still goes through the silent-on-success "
        "helper; only Set-W4NamedText may use it")
    assert _function("Set-W4NamedText").count("Release-Transient") == 3
    assert "Invoke-NamedRelease" not in code, (
        "Invoke-NamedRelease keeps the release count to itself")
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-W4Release $rel" in code and label in code, label
    ledger_at = code.index("$rel = New-ReleaseLedger")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert ledger_at < excel_at
    assert code.count("$comAcquired = $comAcquired + 1") == 3, (
        "the three session objects are each counted where they are acquired")
    # AND THE ITERATION RANGE IS COUNTED TOO. It is the only other COM object
    # this runner opens itself, and an uncounted acquisition would break the
    # balance check without anything saying which one.
    assert "$comAcquired = $comAcquired + [int]$iterationBlock.Acquired" in code
    block = _function("Get-W4IterationBlock")
    assert block.count("Invoke-W4Release $Ledger") == 3
    assert block.count("$acquired = $acquired + 1") == 3
    helper = _function("Invoke-W4Release")
    assert "Release-ComObjectSafe" in helper
    assert "[int]$rec.Count -ne 0" in helper
    assert "$script:W4Residual.Add" in helper


def test_26_the_lifecycle_verdict_checks_are_all_present() -> None:
    code = _own_code()
    for claim in ("'the owned Excel process exited naturally' $naturalExit",
                  "'no emergency cleanup was required' (-not $emergencyRequired)",
                  "'every COM object this runner acquired was released'",
                  "'every COM release succeeded'",
                  "'every COM release left 0 outstanding references'",
                  "'every transient release inside the accepted readers succeeded'"):
        assert claim in code, claim
    # THE ACCEPTED READERS OWN THEIR OWN TRANSIENTS, and that gate is reported
    # rather than silently assumed: extending the residual accounting into them
    # would mean editing accepted files.
    assert "Get-TransientFailures" in code


def test_27_the_shutdown_is_the_accepted_path_and_emergency_is_never_a_pass() -> None:
    code = _own_code()
    shutdown = code[code.index("} catch {\n    $fatal = (Format-Err $_)"):]
    for required in ("$wb.Close($false)", "$excel.Quit()", "Wait-ExcelExit",
                     "Invoke-EmergencyExcelCleanup", "EXCEL SHUTDOWN"):
        assert required in shutdown, required
    assert shutdown.count("finally") >= 1
    assert shutdown.count("EXCEL SHUTDOWN:") >= 2
    clear_at = shutdown.index("$Error.Clear()")
    collect_at = shutdown.index("[System.GC]::Collect()")
    wait_at = shutdown.index("Wait-ExcelExit")
    assert clear_at < collect_at < wait_at
    match = re.search(r"Wait-ExcelExit -Identity \$excelIdentity -TimeoutSeconds (\d+)", code)
    assert match and int(match.group(1)) >= 60
    assert "$emergencyRequired = $true" in code
    assert "if ($ok) { exit 0 } else { exit 1 }" in code


def test_28_no_hidden_reference_pattern_and_nothing_is_saved() -> None:
    """The two patterns W1 refused structurally, held refused here, and the
    workbook is a disposable copy that is never written back."""
    own = _own_code()
    chains = re.findall(r"\$(?:wb|excel|workbooks)\.\w+\.\w+", own)
    assert not chains, f"a chained COM property access acquires an unnamed intermediate: {chains}"
    assert "$tempRoot = Join-Path ([System.IO.Path]::GetTempPath())" in own
    assert "$workbooks.Open($stageBPath)" in own
    for forbidden in (".Save()", ".SaveAs(", ".SaveCopyAs(", "Stop-Process"):
        assert forbidden not in own, forbidden
    assert "$wb.Close($false)" in own


def test_29_the_tree_is_proved_clean_before_excel_is_started() -> None:
    code = _own_code()
    revision = _function("Get-W4SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    body = _function("Write-W4Line")
    assert "Set-Content -LiteralPath $script:W4Path" in body
    code = _own_code()
    assert "'PREREQUISITE'" in code
    assert "$results.Count -gt 0" in code, (
        "a run that recorded no RESULT at all would otherwise pass")




# ===========================================================================
# H. ORDER AND RESTRAINT - WHAT MAKES W4 W4
# ===========================================================================
# The refusal is only evidence if it happens BEFORE the first successful run,
# and W4 is only a baseline if it STOPS at one. Both are controls, not comments.

def _at(needle: str) -> int:
    code = _own_code()
    assert needle in code, needle
    return code.index(needle)


def test_31_the_refusal_is_attempted_before_any_successful_simulation() -> None:
    """AFTERWARDS THE ENDPOINT WOULD LEGITIMATELY SUCCEED, so an annual attempt
    made after the run would prove the opposite of what W4 claims."""
    annual_at = _at("$announcement = Invoke-W4Annual -Excel $excel -P7 $p7")
    simulation_at = _at("$result = Invoke-Phase6Simulation -Excel $excel")
    assert annual_at < simulation_at, (
        "the annual attempt is made after the simulation; a refusal there would "
        "be a different fact and would not be the no-run refusal")
    # AND THE STATE IT NEEDS IS PROVED, not assumed: nothing has been published.
    empty_at = _at("'no simulation has been published yet'")
    assert empty_at < annual_at, (
        "the runner does not establish that no simulation exists before asking "
        "the endpoint to refuse for that reason")
    assert "[string]::IsNullOrEmpty($statusBefore)" in _own_code()


def test_32_w4_never_runs_the_annual_endpoint_successfully() -> None:
    """W4 ENDS AT A SUCCESSFUL SIMULATION. Producing an annual answer here would
    consume the first-annual-run state W5 exists to observe."""
    code = _own_code()
    assert code.count("Invoke-W4Annual") == 2, (
        "the annual endpoint is invoked from more than one call site")
    annual_at = _at("$announcement = Invoke-W4Annual -Excel $excel -P7 $p7")
    simulation_at = _at("$result = Invoke-Phase6Simulation -Excel $excel")
    assert code.find("Invoke-W4Annual", simulation_at) == -1, (
        "the annual endpoint is invoked again after the simulation succeeds; "
        "that is W5's subject and W4 must not consume it")
    # The runner says so where the next reader will look.
    assert "W4 STOPS AT A SUCCESSFUL SIMULATION" in code
    assert annual_at > 0


def test_33_the_annual_attempt_is_required_to_be_refused() -> None:
    """A refusal is the RESULT here, so an attempt that succeeded must fail the
    scenario rather than be reported as an interesting outcome."""
    code = _own_code()
    assert "$announcement -like 'FAIL|*'" in code
    assert "the annual endpoint REFUSES with no successful simulation" in code
    # And the answer comes from production, not from the runner's judgement.
    annual = _function("Invoke-W4Annual")
    assert "PCCM_AutomationResult" in annual
    assert "$P7.command_surface.annual_endpoint" in annual, (
        "the annual endpoint is named in the runner rather than read from the "
        "projection")


# ===========================================================================
# I. THE REFUSAL MOVED NOTHING
# ===========================================================================

def test_34_the_frozen_identity_set_is_the_whole_published_state() -> None:
    """EVERY published run-identity field in BOTH banks, the AUTO nonce and its
    durable pending marker - minus the two rows that are DERIVED and must be
    allowed to move."""
    invariants = _function("Get-W4RunInvariants")
    assert "Get-Phase6State" in invariants
    assert "$state['shared']" in invariants
    assert "foreach ($bank in @($Inspection.publication.bank_labels))" in invariants, (
        "the bank blocks are not frozen, so a refusal could write one")
    assert "$out.Add('pending_auto_nonce'" in invariants
    code = _own_code()
    assert "$script:W4DerivedRows = @('simulation_status', 'status_evaluated_at')" in code
    # THE TWO EXCLUSIONS ARE CHECKED SEPARATELY, on the terms that matter.
    assert "the refused attempt left the derived simulation status word unchanged" in code
    # The frozen set really covers the fields the authorisation names.
    projected = json.loads(
        (BUILD / "phase6_gate_b_inspection.json").read_text(encoding="utf-8"))
    rows = projected["sim_data"]["run_identity"]["rows"]
    for field in ("run_id", "result_digest", "request_fingerprint", "next_auto_nonce",
                  "last_run_id", "last_attempt_result", "consumed_auto_nonce"):
        assert field in rows, field
        assert field not in ("simulation_status", "status_evaluated_at")


def test_35_the_refusal_published_nothing_in_either_bank() -> None:
    """A STAMP AND A RECORD ROW ARE DIFFERENT FACTS. A run could in principle
    write one and not the other, so both are read, in both banks."""
    code = _own_code()
    assert "foreach ($bank in @($simInspection.publication.bank_labels))" in code
    assert "wrote no annual stamp in bank" in code
    assert "published no annual record in bank" in code
    stamp = _function("Get-W4AnnualStamp")
    record = _function("Get-W4AnnualFirstRecord")
    assert "$P7.annual_records.stamp" in stamp
    assert "$records.index_columns.$Bank" in record
    assert "selected_px_profile_columns" in record
    # NOT ONE COLUMN LETTER OR ROW NUMBER IS TYPED INTO THE RUNNER.
    for body in (stamp, record):
        assert not re.search(r"'[A-Z]{1,2}\d{1,3}'", body), (
            "an address literal appears where the projection should supply it")


def test_36_the_handoff_still_answers_the_unrun_state_after_the_refusal() -> None:
    code = _own_code()
    assert "the annual distribution state is still NOT PRODUCED" in code
    assert "the annual profile state is still NOT PRODUCED" in code
    assert "no annual profile Px is reported" in code
    assert "no fake year count is reported" in code
    assert "$p7.handoff.distribution_states[0]" in code
    assert "$p7.handoff.profile_states[0]" in code
    assert "'NOT PRODUCED'" not in code, (
        "the unrun state is typed into the runner; it is the contract's to say")
    # AND THE REFUSAL DID NOT DAMAGE THE DETERMINISTIC CALCULATION.
    assert "the deterministic calculation is still CURRENT after the refusal" in code


# ===========================================================================
# J. THE FIXED-SEED BASELINE
# ===========================================================================

def test_37_every_request_field_is_checked_exactly() -> None:
    """EXACT, AND ONLY FOR A REAL DOUBLE. A seed Excel published as text is a
    different fact from a seed it published as a number, and the accepted
    comparison refuses to coerce the first into the second."""
    code = _own_code()
    for claim in ("the effective seed is exactly the FIXED seed requested",
                  "the supplied seed is republished unchanged",
                  "the run records the FIXED seed mode",
                  "the accepted iteration count is the requested one"):
        assert claim in code, claim
    # EACH FIELD BY NAME. A count alone would survive one of them being
    # loosened into a mere "is not blank".
    for expression in (
            "Test-SimExactDouble -Actual $block['effective_seed'] -Expected $suppliedSeed",
            "Test-SimExactDouble -Actual $block['supplied_seed'] -Expected $suppliedSeed",
            "Test-SimExactDouble -Actual $block['iterations_run'] -Expected ([double]$iterations)",
            "Test-SimExactText -Actual $block['seed_mode'] -Expected ([string]$case.seed_mode)"):
        assert expression in code, expression
    # The expected values come from the corpus, never from a literal here.
    assert "-Expected $suppliedSeed" in code
    assert "-Expected ([double]$iterations)" in code
    assert "$suppliedSeed = [double]$case.supplied_seed" in code
    assert "$iterations = [int]$case.iterations" in code


def test_38_the_identity_w5_will_bind_to_is_captured() -> None:
    code = _own_code()
    for field in ("'run_id'", "'request_fingerprint'", "'result_digest'",
                  "'last_successful_stamp'", "'applied_timeline'"):
        assert field in code, field
    assert "the run published a " in code
    assert "the shared last run id is the run just published" in code
    assert "the attempt is recorded as a success" in code
    assert "a publication bank is active" in code
    assert "@($simInspection.publication.bank_labels) -contains $bank" in code
    # THE WHOLE STATE IS WRITTEN OUT, so W5 has something to bind against.
    assert "THE BASELINE IDENTITY W5 WILL BIND TO" in code
    assert "Format-Phase6State -State $state" in code


def test_39_the_fixed_seed_nonce_contract_is_proved_not_assumed() -> None:
    """THE AUTO NONCE LIFECYCLE GOVERNS AUTO RUNS. A FIXED request allocates
    nothing, so a nonce movement here would be a defect - not, as it would be
    under AUTO, the expected outcome."""
    code = _own_code()
    assert "the FIXED run consumed no AUTO nonce" in code
    assert "the FIXED run did not advance the AUTO nonce" in code
    assert "the FIXED run left the pending AUTO nonce marker alone" in code
    assert "$nonceBefore = $stateAfter['shared']['next_auto_nonce']" in code
    assert "$pendingBefore = $stateAfter['pending_auto_nonce']" in code
    # AND THE CONTRACT REALLY SAYS THE NONCE IS AUTO's.
    contract = (PCCM_ROOT / "spec" / "sim_contract.yaml").read_text(encoding="utf-8")
    assert re.search(r"^\s*auto:\s*$", contract, re.M)
    assert "nonce_lifecycle:" in contract
    # Production writes a blank consumed nonce when none was consumed, which is
    # what makes the blank check a true statement rather than a hopeful one.
    report = (PCCM_ROOT / "src" / "vba" / "modSimReport.bas").read_text(encoding="utf-8")
    assert "If package.NonceConsumed Then" in report
    assert "vbNullString" in report


def test_40_the_iteration_column_is_checked_where_no_tolerance_is_needed() -> None:
    """A MINIMUM AND A MAXIMUM ARE SELECTIONS, not accumulations: they are
    values that appear in the column, so they are compared exactly. The mean and
    the ladder are accumulations no authority owns for this model, and are
    recorded rather than compared."""
    code = _own_code()
    assert "the iteration column carries " in code
    assert "the published iteration indices are 1..N with no hole" in code
    assert "every published iteration total is a number in both measures" in code
    assert "the published nominal minimum is the minimum of the published column" in code
    assert "the published nominal maximum is the maximum of the published column" in code
    assert "Test-SimExactDouble -Actual $publishedMin" in code
    assert "Test-SimExactDouble -Actual $publishedMax" in code
    # RECORDED, NOT COMPARED - and the runner says which is which.
    assert "RECORDED and not compared" in code
    assert "@('mean', 'sample_standard_deviation', 'minimum', 'maximum')" in code, (
        "the summary ladder is no longer recorded, so a reader has no way to "
        "see what the run produced")
    assert "-RowKey $rowKey" in code
    # No tolerance is invented for the accumulations.
    assert not re.search(r"\b\d*\.?\d+e-\d+\b", code)


def test_41_the_runner_states_what_no_oracle_owns() -> None:
    """THE HONESTY CONTROL. The strongest thing W4 can say about its stochastic
    output is what an authority owns; the report has to say that plainly rather
    than let a reader assume the ladder was validated."""
    text = _text()
    assert "THERE IS NO INDEPENDENT STOCHASTIC EXPECTATION FOR THIS MODEL" in text
    assert "seed 12345" in text, (
        "the report does not say why the accepted Phase-6 oracle cannot speak "
        "for this run")
    # AND THAT CLAIM IS TRUE OF THE ARTEFACT.
    gate_b = json.loads((BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    assert gate_b["supplied_seed"] == 12345
    assert gate_b["supplied_seed"] != _w4()["supplied_seed"]
    assert gate_b["case_count"] == len(gate_b["parity_cases"])
    for case in gate_b["parity_cases"]:
        assert "plan_case_id" in case, (
            "the Phase-6 corpus is no longer a per-plan-case parity corpus; if "
            "it now owns a whole-workbook simulation, W4 should use it")


def test_42_the_projected_vocabulary_is_used_instead_of_typed_words() -> None:
    """FIXED, CURRENT and SUCCESS are the contract's words, and the runner
    checks them against the projection that owns them."""
    code = _own_code()
    assert "@($gateBCases.vocabulary.seed_modes) -contains" in code
    assert "@($gateBCases.vocabulary.sim_states) -contains" in code
    assert "@($gateBCases.vocabulary.attempt_results) -contains" in code
    assert "@($gateBCases.vocabulary.quantile_labels) -contains" in code
    gate_b = json.loads((BUILD / "phase6_gate_b_cases.json").read_text(encoding="utf-8"))
    assert "FIXED" in gate_b["vocabulary"]["seed_modes"]
    assert "CURRENT" in gate_b["vocabulary"]["sim_states"]
    assert "SUCCESS" in gate_b["vocabulary"]["attempt_results"]
    assert _w4()["selected_confidence_level"] in gate_b["vocabulary"]["quantile_labels"]
