#!/usr/bin/env python3
"""P7-7: the MINIMAL W2 runner, proved on Linux before a Windows session.

W2 IS AN INHERITED PHASE-5 REGRESSION SCENARIO. Phase 7 gave `DriverFactors`
two dynamic arrays per driver; W2 asks whether ~300 instances of that UDT still
cross the accepted Phase-5 calculation bridge and publish the answer the
INDEPENDENT Phase-5 oracle computed for the same model. The array-LENGTH
dimension belongs to W3 and is deliberately not combined with this one.

WHAT THIS FILE CAN AND CANNOT PROVE. There is no PowerShell and no Excel here,
so nothing below claims the runner RAN. What it proves is what a reader would
otherwise take on trust: that the fixture really is 300 drivers over 5 years and
really came from the independent oracle; that every projected column of the
compared tables is actually compared, and by a rule the contract drives rather
than a list of names typed into the runner; that the allowance is the project's
own accepted floor and no new tolerance was invented; that the call closure is
complete, which is the defect that cost W1 a Windows session; and that the COM
lifecycle discipline W1 closed on is carried over intact.
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

import test_phase7_acceptance_harness_source as accepted  # noqa: E402
from pccm_builder import load_calc_contract  # noqa: E402

WINDOWS = PCCM_ROOT / "bootstrap" / "windows"
RUNNER = WINDOWS / "phase7_w2_many_drivers.ps1"
W1 = WINDOWS / "phase7_w1_smoke.ps1"
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

# W1 IS CLOSED. Its runner is accepted Windows evidence and is pinned, not
# maintained, by this round. The digest is the LITERAL one accepted at 164113f:
# hashing the file at import time would be a control that can never fail.
W1_SHA256 = "bf41a2193dbf2ea5fbddbe4dad3c6e94649f21262dcf9800ae71eda6e274b59f"

_CACHE: dict = {}


def _text() -> str:
    return RUNNER.read_text(encoding="utf-8")


def _code() -> str:
    """The runner with its block comment and line comments removed.

    Its header explains what W2 refuses to do and names the things the controls
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
    assert body is not None, f"{name} is not defined in the W2 runner"
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


def _w2() -> dict:
    case = [s for s in _cases()["scenarios"] if s["id"] == "W2"]
    assert case, "the acceptance corpus carries no W2 scenario"
    return case[0]


def _inspection() -> dict:
    if "inspection" not in _CACHE:
        _CACHE["inspection"] = json.loads(
            (BUILD / "phase5_gate_b_inspection.json").read_text(encoding="utf-8"))
    return _CACHE["inspection"]


# ===========================================================================
# A. SCOPE: THIS IS W2, AND W1 IS CLOSED
# ===========================================================================

def test_01_the_runner_exists_and_is_dedicated() -> None:
    assert RUNNER.exists()
    lines = _text().splitlines()
    assert len(lines) < 1100, (
        f"{len(lines)} lines; W2 is one scenario and most of its behaviour is "
        "meant to come from the accepted Phase-5 fixture and readers")
    # Most of the file is either the accepted copied helpers or W2's own small
    # comparison; neither should have grown into a second harness.
    assert len(_own_code().splitlines()) < 900


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


def test_04_no_phase7_annual_no_sensitivity_no_later_scenario() -> None:
    """W2 is an INHERITED Phase-5 regression. Anything stochastic in here would
    be a scenario nobody authorised it to run."""
    code = _own_code()
    for forbidden in ("PCCM_RunAnnualStochastic", "PCCM_RunSensitivity",
                      "PCCM_RunSimulation", "PCCM_AnnualDistributionState",
                      "PCCM_AnnualProfileState", "PCCM_AnnualProfilePx",
                      "PCCM_AnnualYearCount", "phase7_acceptance_inspection",
                      "Invoke-Phase5GateBScenarios", "Invoke-Phase6GateBScenarios",
                      "phase7_acceptance_scenarios"):
        assert forbidden not in code, f"{forbidden} is outside W2"
    for scenario in ("'W3'", "'W4'", "'W5'", "'W6'", "'W7'", "'W8'"):
        assert scenario not in code, scenario
    # The only scenario it names is its own.
    assert "'W2'" in code


def test_05_it_does_not_repeat_the_w1_surface_matrix() -> None:
    """One lightweight read-only compile trigger is authorised because Excel
    must execute the current project. The surface matrix is W1's."""
    code = _own_code()
    for forbidden in ("VBProject", "VBComponents", "CodeModule", "ProcOfLine"):
        assert forbidden not in code, f"{forbidden} is W1's evidence, not W2's"
    assert "$excel.Run('PCCM_CalculationStatus')" in code
    assert "the current VBAProject compiles in real Excel" in code


def test_06_the_accepted_w1_runner_is_untouched() -> None:
    """W1 IS CLOSED on accepted Windows evidence. This round does not modify or
    rerun it, and an edit to it must fail here rather than pass unnoticed."""
    assert W1.exists(), "the accepted W1 runner is missing"
    assert hashlib.sha256(W1.read_bytes()).hexdigest() == W1_SHA256
    assert "phase7_w1_smoke" not in _code(), "W2 reaches into the W1 runner"


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
        assert ours, f"{name} is not defined in the W2 runner"
        assert ours.group(1) == theirs.group(1), (
            f"{name} has drifted from the accepted implementation")
        body = own
        for line in body.splitlines():
            if line.strip().startswith("function "):
                continue
        assert not re.search(rf"(?<![\w\-.$]){name}(?![\w\-])",
                             "\n".join(l for l in own.splitlines()
                                       if not l.strip().startswith("function "))), (
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

def test_11_the_fixture_is_the_authorised_driver_count_dimension() -> None:
    case = _w2()
    model = case["model"]
    drivers = len(model["cost_lines"]) + len(model["risks"])
    assert case["dimension"] == "driver_count"
    assert 250 <= drivers <= 350, drivers
    assert drivers == 300, drivers
    assert len(model["cost_lines"]) == 180 and len(model["risks"]) == 120
    # FIVE YEARS, NOT 200. The length dimension is W3's and must not be combined.
    assert model["timeline"]["duration"] == 5
    other = [s for s in _cases()["scenarios"] if s["id"] == "W3"][0]
    assert other["model"]["timeline"]["duration"] > model["timeline"]["duration"]
    assert len(other["model"]["cost_lines"]) + len(other["model"]["risks"]) < drivers


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
    case = _w2()
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
    expected = _w2()["expected"]
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


def test_16_the_five_year_axis_is_five_years_everywhere_in_the_corpus() -> None:
    """THE NO-TRUNCATION PROPERTY, stated on the authority side first."""
    case = _w2()
    years = case["model"]["timeline"]["duration"]
    expected = case["expected"]
    assert len(expected["calc_years"]) == years
    assert len(expected["annual"]) == years
    assert len(expected["discount_factors"]) == years
    profiles = len(case["model"]["inflation"])
    assert len(expected["inflation_factors"]) == profiles * years
    # Every driver's weight array is the year array's length: a truncated or
    # shared array inside any one of the 300 UDT instances moves that driver's
    # Knom and Kpv, which the runner compares one driver at a time.
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
        # `$Allowance` is the parameter Compare-W2Table forwards; `$allowance`
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
    cell = _function("Compare-W2Cell")
    assert "$null -eq $Expect" in cell, "a missing oracle value has no rule"
    assert "$Expect -is [string]" in cell
    assert "$Got -isnot [double]" in cell, (
        "a value published as text would be coerced instead of failing")
    table = _function("Compare-W2Table")
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
    cell = _function("Compare-W2Cell")
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
    row = _w2()["expected"]["inflation_factors"][0]
    for column, field in pairs.items():
        assert column in columns, column
        assert field in row, field


# ===========================================================================
# F. THE DRIVER-COUNT EVIDENCE
# ===========================================================================

def test_21_all_three_hundred_drivers_are_proved_present_and_in_order() -> None:
    code = _own_code()
    assert "calc_drivers publishes one row per supplied driver" in code
    assert "the published drivers are in supply order, one row per driver" in code
    assert "no driver is published twice and none is missing" in code
    assert "Sort-Object -Unique" in code


def test_22_misalignment_is_checked_independently_of_order() -> None:
    """THE CONTROL FOR A SHARED OR REUSED DriverFactors INSTANCE.

    Comparing row for row proves order and values together, so it cannot say
    which of the two went wrong. Matching each published row to the oracle
    record its OWN id names separates them: this failing while the positional
    comparison passes means the numbers are wrong, and the reverse means the
    numbers are right and the order is not.
    """
    code = _own_code()
    assert "$byId = @{}" in code
    assert "$byId.ContainsKey($id)" in code
    assert "every driver carries the factors of the driver its own id names" in code
    by_id_at = code.index("every driver carries the factors")
    positional_at = code.index("Compare-W2Table -Live $liveDrivers")
    assert by_id_at < positional_at, (
        "the order-independent check must be reported before the positional "
        "one, so a reader sees which of the two failed first")


def test_23_the_year_arrays_are_proved_directly_as_well_as_through_the_factors() -> None:
    code = _own_code()
    assert "calc_years publishes exactly" in code
    assert "calc_annual publishes exactly" in code
    assert "calc_inflation_factors publishes one row per profile per project year" in code
    assert "the applied timeline is the one the oracle computed against" in code
    assert "last_successful_applied_timeline" in code


def test_24_the_calculation_runs_through_its_own_accepted_path() -> None:
    """Through production's own operation and production's own status, never by
    the runner writing a cell or judging currency for itself."""
    code = _own_code()
    assert "Invoke-Phase5ProductionOperation -Excel $excel" in code
    assert "-Operation 'PCCM_Calculate'" in code
    assert "the calculation reports CURRENT" in code
    assert "Set-Phase5Fixture -Excel $excel" in code
    assert "Save-Phase5LockedFxSeed" in code
    # THE RUNNER WRITES NOTHING ITSELF.
    for banned in ("$rng.Value2 =", ".ListObjects", "$ws.Cells"):
        assert banned not in _own_code(), banned


# ===========================================================================
# G. COM LIFECYCLE - THE DISCIPLINE W1 CLOSED ON
# ===========================================================================

def test_25_every_acquisition_is_counted_and_released_into_one_ledger() -> None:
    code = _own_code()
    assert "Release-Transient" not in code, (
        "a release still goes through the silent-on-success helper")
    assert "Invoke-NamedRelease" not in code, (
        "Invoke-NamedRelease keeps the release count to itself")
    for label in ("'Workbook'", "'Workbooks'", "'Excel.Application'"):
        assert f"Invoke-W2Release $rel" in code and label in code, label
    ledger_at = code.index("$rel = New-ReleaseLedger")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert ledger_at < excel_at
    assert code.count("$comAcquired = $comAcquired + 1") == 3
    helper = _function("Invoke-W2Release")
    assert "Release-ComObjectSafe" in helper
    assert "[int]$rec.Count -ne 0" in helper
    assert "$script:W2Residual.Add" in helper


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
    revision = _function("Get-W2SourceRevision")
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in revision
    assert "rev-parse HEAD" in revision
    refusal_at = code.index("REFUSED, BEFORE EXCEL WAS STARTED.")
    excel_at = code.index("New-Object -ComObject Excel.Application")
    assert refusal_at < excel_at
    assert "exit 1" in code[refusal_at:excel_at]


def test_30_the_report_is_written_through_and_separates_prerequisites() -> None:
    body = _function("Write-W2Line")
    assert "Set-Content -LiteralPath $script:W2Path" in body
    code = _own_code()
    assert "'PREREQUISITE'" in code
    assert "$results.Count -gt 0" in code, (
        "a run that recorded no RESULT at all would otherwise pass")
