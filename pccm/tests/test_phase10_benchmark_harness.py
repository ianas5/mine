#!/usr/bin/env python3
"""P10-4A: the performance benchmark harness, before anyone runs it.

`bootstrap/windows/phase10_benchmark.ps1` cannot be executed on Linux, and this
batch is forbidden from executing it on Windows. So everything that can be
proved without Excel is proved here, BEFORE the operator is asked to spend an
afternoon of real machine time on it - because the failure mode of a benchmark
is not a crash. It is a number that looks fine and measured the wrong thing.

WHAT COULD GO WRONG WITH A BENCHMARK, AND WHAT IS ASSERTED ABOUT EACH

  it measures the wrong work    Setup, Excel startup, the workbook open and the
                                fixture build dwarf several of the operations
                                being timed. The clock must surround one call
                                and nothing else, and the setup figures must be
                                reported in their own field.

  it measures a refusal         A command that refuses is the fastest command in
                                the report. Every timed execution carries its own
                                evidence, and a sample that fails a gate is not a
                                fast sample - it is not a sample.

  it measures less work than    A stochastic run that quietly executed ten
  it says                       thousand iterations when a hundred thousand were
                                asked for is fast and useless. The count the
                                workbook actually executed is read back and
                                compared.

  it measures something else    A benchmark-only code path would measure the
                                entirely                      harness. The
                                runner calls the real endpoints by name and
                                implements no calculation of its own.

  it is not comparable          Two runs are comparable only if they did the same
                                work on a machine whose record exists. The seed
                                is fixed, the scenarios are deterministic, and
                                every environment field the plan names is present.

  it passes or fails for a      This is the FIRST baseline. There is nothing to
  reason nobody can defend      compare against, so there is no pass mark - and
                                the ratios that apply to LATER runs are recorded
                                rather than applied to this one.

THE PLAN IS THE AUTHORITY. Sizes, split, years, iteration counts, forbidden
combinations, the timing method and the policy are emitted by
`builder/pccm_builder/benchmark.py` into `build/phase10_benchmark_plan.json`.
The runner reads that file. These controls check the plan exactly and check that
the runner declares no matrix of its own.

Runs standalone or under pytest.
"""
from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
REPO_ROOT = PCCM_ROOT.parent
BOOTSTRAP = PCCM_ROOT / "bootstrap" / "windows"
SPEC = PCCM_ROOT / "spec"
SRC_VBA = PCCM_ROOT / "src" / "vba"
DOCS = PCCM_ROOT / "docs"
BUILD = PCCM_ROOT / "build"
BUILDER = PCCM_ROOT / "builder"
sys.path.insert(0, str(BUILDER))

import pytest  # noqa: E402
import yaml  # noqa: E402

from pccm_builder.benchmark import (  # noqa: E402
    BENCHMARK_SCHEMA_VERSION,
    build_benchmark_plan,
    cost_line_count,
)

RUNNER = BOOTSTRAP / "phase10_benchmark.ps1"
TIMING = BOOTSTRAP / "phase7_timing_scenarios.ps1"
LIFECYCLE = BOOTSTRAP / "com_lifecycle.ps1"
PLAN_FILE = BUILD / "phase10_benchmark_plan.json"
RECORD = DOCS / "phase10_step1_contract.md"

# The accepted tree this batch is measured against. Step 3 closed at e3f83e2;
# a benchmark harness may not change one line of what it benchmarks.
ACCEPTED = "e3f83e2"

# The dimensions the Phase-10 contract settled. Restated here ONLY so a control
# can compare three independent statements of them: this list, the plan, and the
# sentence in the contract record.
CONTRACT_DIMENSIONS = {"Small": (20, 10), "Medium": (100, 25), "Large": (300, 40)}


_MEMO: dict = {}


def _plan() -> dict:
    if "plan" not in _MEMO:
        _MEMO["plan"] = build_benchmark_plan()
    return _MEMO["plan"]


def _runner() -> str:
    if "runner" not in _MEMO:
        _MEMO["runner"] = RUNNER.read_text(encoding="utf-8")
    return _MEMO["runner"]


def _code() -> str:
    """The runner with its comment lines removed.

    A BAN MUST NOT BE SATISFIED BY THE PARAGRAPH EXPLAINING IT, and a
    requirement must not be satisfied by a comment promising it. Every control
    that asserts what the script DOES reads this.
    """
    if "code" not in _MEMO:
        text = re.sub(r"<#.*?#>", "", _runner(), flags=re.S)
        _MEMO["code"] = "\n".join(
            line for line in text.splitlines()
            if not line.strip().startswith("#"))
    return _MEMO["code"]


def _scenario(key: str) -> dict:
    return next(entry for entry in _plan()["scenarios"] if entry["id"] == key)


def _operation(key: str) -> dict:
    return next(entry for entry in _plan()["operations"] if entry["key"] == key)


def _git(*args: str) -> str:
    if args not in _MEMO:
        _MEMO[args] = subprocess.run(["git", *args], cwd=REPO_ROOT, check=True,
                                     stdout=subprocess.PIPE, text=True).stdout
    return _MEMO[args]


needs_build = pytest.mark.skipif(
    not PLAN_FILE.is_file(), reason="Stage A has not been built into pccm/build")


# ===========================================================================
# A. THE MATRIX
# ===========================================================================
def test_01_the_three_model_sizes_are_the_settled_ones() -> None:
    """REQUIRED CONTROL 1. Twenty over ten, a hundred over twenty-five, three
    hundred over forty - and the sentence in the contract record is read rather
    than trusted, so the plan cannot quietly redefine what Small means."""
    sizes = {entry["title"]: (entry["drivers"], entry["years"])
             for entry in _plan()["scenarios"]}
    assert sizes == CONTRACT_DIMENSIONS, sizes

    settled = re.sub(r"\s+", " ", RECORD.read_text(encoding="utf-8"))
    assert ("**Models:** Small = 20 drivers / 10 years · Medium = 100 / 25 · "
            "Large = 300 / 40.") in settled, "the record no longer settles these sizes"


def test_02_the_scenario_ids_are_stable_and_named() -> None:
    assert [entry["id"] for entry in _plan()["scenarios"]] == [
        "PERF-SMALL", "PERF-MEDIUM", "PERF-LARGE"]


def test_03_the_driver_split_is_the_accepted_one_not_a_new_one() -> None:
    """THE SPLIT IS NOT INVENTED HERE. `phase7_timing_scenarios.ps1` settled it
    for the accepted Phase-7 measurement - 60% Cost Lines, the remainder Risks -
    and the same rule is used, so the two harnesses build comparably shaped
    models and the historical numbers stay readable beside the new ones."""
    accepted = TIMING.read_text(encoding="utf-8")
    assert "20 -> 12/8, 100 -> 60/40," in accepted
    assert "300 -> 180/120" in accepted
    assert "[Math]::Ceiling([double]$DriverCount * 0.6)" in accepted

    assert _plan()["cost_line_share"] == 0.6
    for key, cost, risk in (("PERF-SMALL", 12, 8), ("PERF-MEDIUM", 60, 40),
                            ("PERF-LARGE", 180, 120)):
        entry = _scenario(key)
        assert (entry["cost_lines"], entry["risks"]) == (cost, risk), entry
        assert entry["cost_lines"] + entry["risks"] == entry["drivers"]
        assert cost_line_count(entry["drivers"]) == cost


def test_04_the_iteration_matrix_is_exactly_the_permitted_one() -> None:
    """REQUIRED CONTROL 2."""
    matrix = {entry["id"]: entry["iterations"] for entry in _plan()["scenarios"]}
    assert matrix == {
        "PERF-SMALL": [10_000, 50_000, 100_000],
        "PERF-MEDIUM": [10_000, 50_000, 100_000],
        "PERF-LARGE": [10_000, 50_000],
    }, matrix


def test_05_large_at_one_hundred_thousand_is_structurally_unreachable() -> None:
    """REQUIRED CONTROL 3, and it is asserted three ways because "we did not run
    it" and "it cannot be run" are different claims.

      * it is absent from the expanded matrix;
      * it is RECORDED as forbidden, with a reason, so nobody reads it as
        missing evidence;
      * and the runner refuses it on the command line rather than silently
        measuring nothing.
    """
    for run in _plan()["runs"]:
        assert not (run["scenario"] == "PERF-LARGE" and run["iterations"] == 100_000), run
    assert 100_000 not in _scenario("PERF-LARGE")["iterations"]

    forbidden = _plan()["forbidden"]
    assert len(forbidden) == 1, forbidden
    rule = forbidden[0]
    assert rule["scenario"] == "PERF-LARGE"
    assert rule["iterations"] == 100_000
    assert set(rule["operations"]) == {"simulation", "sensitivity", "annual"}
    assert "caps Large at 50,000" in rule["reason"]
    assert "not required / impractical" in rule["reason"]

    code = _code()
    assert "$declaredIterations -notcontains [int]$_" in code, (
        "the runner does not check an asked-for iteration count against the plan")
    assert "REFUSED: " in code


def test_06_calculate_is_measured_once_per_size_not_once_per_iteration_count() -> None:
    """REQUIRED CONTROL 4. Repeating an iteration-independent operation at 10k,
    50k and 100k would be three measurements of identical work, and would make
    the matrix look richer than the evidence is."""
    for key in ("calculate", "recalculation"):
        assert _operation(key)["iteration_dependent"] is False
        for scenario in _plan()["scenarios"]:
            runs = [run for run in _plan()["runs"]
                    if run["scenario"] == scenario["id"] and run["operation"] == key]
            assert len(runs) == 1, (key, scenario["id"], runs)
            assert runs[0]["iterations"] is None


def test_07_every_stochastic_operation_is_measured_at_every_declared_count() -> None:
    for scenario in _plan()["scenarios"]:
        for key in ("simulation", "sensitivity", "annual"):
            counts = sorted(run["iterations"] for run in _plan()["runs"]
                            if run["scenario"] == scenario["id"] and run["operation"] == key)
            assert counts == sorted(scenario["iterations"]), (key, scenario["id"], counts)


def test_08_the_expanded_matrix_is_the_size_it_should_be() -> None:
    expected = 0
    for scenario in _plan()["scenarios"]:
        expected += 2 + 3 * len(scenario["iterations"])
    assert len(_plan()["runs"]) == expected == 30, len(_plan()["runs"])


# ===========================================================================
# B. THE TIMING METHOD
# ===========================================================================
def test_10_one_cold_run_and_exactly_three_warm_runs_are_required() -> None:
    """REQUIRED CONTROL 5."""
    timing = _plan()["timing"]
    assert timing["cold_runs"] == 1
    assert timing["warm_runs"] == 3
    for run in _plan()["runs"]:
        assert run["cold_runs"] == 1 and run["warm_runs"] == 3, run
    code = _code()
    assert "$total = [int]$run.cold_runs + [int]$run.warm_runs" in code, (
        "the runner does not take its run counts from the plan")
    assert "$phase = $(if ($index -lt [int]$run.cold_runs) { 'cold' } else { 'warm' })" in code


def test_11_the_warm_median_is_the_comparison_statistic() -> None:
    """REQUIRED CONTROL 6, and the mean is refused by name. Averaging three runs
    would average away exactly the outlier a scheduler introduces, and that
    outlier is information about the machine."""
    timing = _plan()["timing"]
    assert timing["comparison_statistic"] == "median_of_warm_runs"
    assert "The mean is NOT used" in timing["aggregation"]

    code = _code()
    assert "function Get-BenchmarkMedian" in code
    assert "Sort-Object" in code
    assert "Measure-Object" not in code, "the runner is averaging something"
    assert "-Average" not in code
    assert "warm_median_ms" in code


def test_12_the_median_needs_three_valid_warm_samples() -> None:
    """A median of two good runs and a refusal is not a measurement of
    anything, so it is not computed at all."""
    assert _plan()["correctness_gates"]["median_requires"] == "three valid warm samples"
    code = _code()
    assert "if ($validWarm.Count -eq [int]$run.warm_runs) {" in code
    assert "NOT COMPUTED" in _code()


def test_13_cold_and_warm_are_defined_and_neither_is_a_workbook_open() -> None:
    """AND COLD IS NOT FAKED. The one thing "cold" is most easily made to mean
    is the cost of starting Excel, which would be a measurement of Excel."""
    timing = _plan()["timing"]
    assert "first execution of the operation" in timing["cold_definition"]
    assert "immediately after the cold one" in timing["warm_definition"]
    for excluded in ("Excel startup", "workbook open", "Stage-B bootstrap",
                     "scenario fixture construction"):
        assert excluded in timing["excluded_from_elapsed"], excluded


def test_14_the_clock_is_monotonic_and_high_resolution() -> None:
    assert _plan()["timing"]["timing_source"] == "System.Diagnostics.Stopwatch"
    assert "[System.Diagnostics.Stopwatch]::StartNew()" in _code()
    assert "Get-Date" not in _code().split("function Invoke-BenchmarkExecution")[1][:900], (
        "a wall clock is being used to time an operation")


def test_15_nothing_but_the_call_is_inside_the_measured_interval() -> None:
    """REQUIRED CONTROL 7, at the only place it can be proved: the body between
    StartNew and Stop.

    The automation envelope is opened before the clock starts, and the
    announcement and every evidence read happen after it stops.
    """
    body = _code().split("function Invoke-BenchmarkExecution")[1]
    envelope = body[:body.index("$evidence = ")]
    body = envelope
    measured = re.findall(r"StartNew\(\)\s*(.*?)\s*\$watch\.Stop\(\)", body, re.S)
    assert len(measured) == 2, measured
    for interval in measured:
        statements = [line.strip() for line in interval.splitlines() if line.strip()]
        assert len(statements) == 1, statements
        assert ("$Excel.CalculateFull()" in statements[0]
                or "$Excel.Run([string]$Operation.endpoint)" in statements[0]), statements
    # AND THE AUTOMATION ENVELOPE IS OPENED BEFORE THE CLOCK, so answering a
    # confirmation prompt is never inside a measurement.
    assert "PCCM_AutomationBegin" in envelope.split("StartNew()")[1]
    assert envelope.index("PCCM_AutomationBegin") < envelope.index(
        "$Excel.Run([string]$Operation.endpoint)")


def test_16_setup_is_timed_separately_and_reported_as_setup() -> None:
    """REQUIRED CONTROL 7's other half. Setup is not hidden - it is measured,
    reported in its own field, and labelled where a reader will see it."""
    code = _code()
    for key in ("stage_b_bootstrap_ms", "excel_startup_ms", "workbook_open_ms",
                "scenario_fixture_ms"):
        assert key in code, key
    assert "$report.Add('setup_ms', $setupTimings)" in code
    assert "SETUP, part of no measurement" in _code(), (
        "the runner does not label setup where a reader of the log will see it")
    assert "None of it is inside any operation elapsed time" in _code()


# ===========================================================================
# C. THE ENVIRONMENT RECORD
# ===========================================================================
def test_20_the_environment_inventory_names_windows_excel_and_its_bitness() -> None:
    """REQUIRED CONTROL 8."""
    fields = _plan()["environment_fields"]
    for field in ("windows_edition", "windows_version", "windows_build",
                  "excel_version", "excel_build", "excel_bitness",
                  "excel_calculation_mode", "other_workbooks_open"):
        assert field in fields, field


def test_21_the_cpu_and_the_memory_are_captured() -> None:
    """REQUIRED CONTROL 9."""
    fields = _plan()["environment_fields"]
    for field in ("cpu_model", "logical_processors", "installed_ram_gb"):
        assert field in fields, field
    code = _code()
    assert "Win32_Processor" in code
    assert "Win32_ComputerSystem" in code
    assert "Win32_OperatingSystem" in code


def test_22_the_workbook_path_and_what_kind_of_place_it_is_are_captured() -> None:
    """REQUIRED CONTROL 10, and the repository is recorded SEPARATELY.

    The timed workbook is a disposable build made under -WorkDir; the repository
    it was built from is somewhere else, and on this user's machine that
    somewhere is OneDrive. Recording only the workbook path would let a synced
    repository be read as a local unsynchronised one, which is exactly the
    silence the contract asked to avoid.
    """
    fields = _plan()["environment_fields"]
    for field in ("workbook_path", "workbook_location_type", "workbook_location_note",
                  "repository_path", "repository_location_type", "onedrive_roots"):
        assert field in fields, field
    assert set(_plan()["location_types"]) == {"local", "onedrive", "synced", "unknown"}

    code = _code()
    assert "function Get-BenchmarkLocationType" in code
    assert "OneDriveCommercial" in code and "OneDriveConsumer" in code
    assert "'onedrive'" in code and "'synced'" in code
    assert "-RepositoryPath $repoRoot" in code
    # AND NO CLAIM IS MADE ABOUT WHAT ANY OF IT DOES. The harness records where
    # the file is; it does not theorise about sync overhead.
    assert "slower" not in _runner().lower()
    assert "no claim is made" in _runner().lower()


def test_23_the_commit_and_the_release_metadata_are_captured() -> None:
    """REQUIRED CONTROL 11."""
    fields = _plan()["environment_fields"]
    for field in ("git_branch", "git_commit", "git_worktree_clean",
                  "model_version", "builder_version", "build_phase",
                  "harness_version", "schema_version"):
        assert field in fields, field
    code = _code()
    assert "rev-parse --abbrev-ref HEAD" in code
    assert "$Manifest.model_version" in code
    assert "$Manifest.builder_version" in code
    assert "$Manifest.build_phase" in code


def test_24_every_declared_environment_field_is_actually_populated() -> None:
    """A FIELD THE PLAN NAMES AND THE RUNNER NEVER SETS would be a promise in
    one file and silence in the other."""
    code = _code()
    for field in _plan()["environment_fields"]:
        assert f"$record.Add('{field}'" in code, f"the runner never records {field}"


def test_25_an_unavailable_field_is_recorded_rather_than_omitted() -> None:
    """SO A MISSING FIELD ALWAYS MEANS THE HARNESS DID NOT ASK, never that the
    machine did not answer."""
    code = _code()
    assert "function Get-BenchmarkUnavailable" in code
    assert "return 'unavailable'" in code
    assert code.count("Get-BenchmarkUnavailable") >= 6


def test_26_the_runner_prints_the_environment_in_the_plans_own_order() -> None:
    assert "foreach ($field in @($plan.environment_fields)) {" in _code()


# ===========================================================================
# D. THE SCENARIO IS DETERMINISTIC
# ===========================================================================
def test_30_nothing_in_the_scenario_is_random() -> None:
    """REQUIRED CONTROL 12. The same scenario id and the same plan produce
    byte-identical inputs on every machine and every run - which is what makes
    two benchmark runs a comparison rather than an anecdote."""
    code = _code()
    for banned in ("Get-Random", "[System.Random]", "New-Guid", "[guid]::NewGuid"):
        assert banned not in code, f"the scenario builder uses {banned}"
    assert "function Get-BenchmarkSeed" in code
    assert "return 20260101" in code
    assert "'FIXED'" in code
    # AND THE SEED IS SET BEFORE ANY OPERATION RUNS.
    assert "controls.random_seed.defined_name" in code


def test_31_the_model_is_built_from_the_plans_dimensions_not_from_literals() -> None:
    """THE RUNNER DECLARES NO MATRIX OF ITS OWN. Every dimension it builds with
    is read from the scenario the plan handed it."""
    code = _code()
    assert "$costCount = [int]$ScenarioSpec.cost_lines" in code
    assert "$riskCount = [int]$ScenarioSpec.risks" in code
    assert "$years = [int]$ScenarioSpec.years" in code
    for literal in ("20", "100", "300"):
        assert f"DriverCount = {literal}" not in code
    assert "Get-Phase7TimingScenarios" not in code


def test_32_the_profile_weights_sum_to_exactly_one() -> None:
    """THE ONE ARITHMETIC THE HARNESS DOES, and it is done so that the accepted
    profiling rule is satisfied exactly rather than to six decimal places. The
    Python restatement below is the same construction; a weights vector that
    summed to 0.9999999999999999 would be refused by production, correctly, and
    the harness must never provoke that."""
    for years in (10, 25, 40):
        for offset in (1, 2, 7, 13):
            weights = _weights(years, offset)
            assert len(weights) == years
            assert sum(weights) == 1.0, (years, offset, sum(weights))
            assert all(weight > 0 for weight in weights), weights
            assert len(set(weights)) > 1, "a flat profile is not representative spend"
    code = _code()
    assert "$weights += [double](1.0 - $running)" in code, (
        "the last year does not absorb the residual, so the profile will not sum to 1")


def _weights(years: int, offset: int) -> list[float]:
    """The runner's construction, restated. Kept beside the control that uses
    it so the arithmetic being claimed is visible."""
    raw = [float(1 + ((year + offset) % 5)) for year in range(1, years + 1)]
    total = sum(raw)
    weights = [round(value / total, 6) for value in raw[:-1]]
    return weights + [1.0 - sum(weights)]


def test_33_no_driver_is_degenerate() -> None:
    """A ZERO-VARIANCE DRIVER IS NOT WHAT IS BEING TIMED. Min < Most Likely <
    Max holds for every driver, so the ranked table is not mostly refusals and
    the sensitivity measurement is of the populated path."""
    code = _code()
    assert "min_value         = $base" in code
    assert "most_likely       = [double]($base * 1.35)" in code
    assert "max_value         = [double]($base * 2.10)" in code
    # AND NO RISK IS CERTAIN OR IMPOSSIBLE.
    assert "((($Index - 1) % 7) + 1) / 10.0" in code


def test_34_the_scenario_exercises_every_surface_the_matrix_names() -> None:
    """FX, INFLATION, PROFILING AND BOTH REGISTERS. A model with one currency,
    no inflation profile and flat weights would be a cheap benchmark of
    something the user does not have."""
    code = _code()
    # THE FX TABLE ITSELF CARRIES TWO DISTINCT CURRENCIES. Asserting that the
    # strings appear somewhere would pass on a model whose FX table listed the
    # reporting currency twice - the conversion path would then never run.
    fx = code.split("        fx            = @(")[1]
    fx = fx[:fx.index("        )")]
    currencies = re.findall(r"currency = '(\w+)'", fx)
    assert sorted(currencies) == ["SAR", "USD"], currencies
    assert "rate = 3.75" in fx, "the foreign currency converts at parity"
    assert "'Standard'" in code and "'Escalated'" in code
    assert "profile_weights" in code
    assert "cost_lines    = $costLines" in code and "risks         = $risks" in code
    assert "discount_rate = 0.05" in code


# ===========================================================================
# E. CORRECTNESS GATES
# ===========================================================================
def test_40_a_refused_operation_is_not_a_performance_sample() -> None:
    """REQUIRED CONTROL 13. A fast failure is not a fast operation."""
    gates = _plan()["correctness_gates"]
    assert "is NOT a performance sample" in gates["invalid_sample_policy"]
    assert "no median is computed from it" in gates["invalid_sample_policy"]

    code = _code()
    assert "function Test-BenchmarkSample" in code
    assert "if ($result -notlike 'OK|*') {" in code
    assert "$validWarm = @($warm | Where-Object { $_.Valid })" in code
    assert "INVALID" in _code()


def test_41_the_iteration_count_is_read_back_and_compared() -> None:
    """REQUIRED CONTROL 14, and it is compared as a NUMBER. A run that executed
    ten thousand iterations when a hundred thousand were asked for would be the
    fastest and most useless measurement in the report."""
    code = _code()
    assert "([int]$published) -ne ([int]$RequestedIterations)" in code
    assert "the workbook executed " in code
    assert "did not report how many iterations it executed" in code
    assert "the published iteration count equals the requested iteration count" in \
        _plan()["correctness_gates"]["iteration_dependent"]


def test_42_every_operation_declares_the_evidence_it_must_produce() -> None:
    for operation in _plan()["operations"]:
        assert operation["evidence"], operation["key"]
        if operation["kind"] == "command":
            assert "automation_result" in operation["evidence"], operation["key"]
    assert "published_iterations" in _operation("simulation")["evidence"]
    assert "annual_year_count" in _operation("annual")["evidence"]
    assert "sensitivity_record_count" in _operation("sensitivity")["evidence"]


def test_43_the_refusal_text_is_kept_rather_than_summarised_away() -> None:
    code = _code()
    assert "$sample.Add('problems', @($_.Problems))" in code
    assert "$sample.Add('evidence', $_.Evidence)" in code
    assert "$report.Add('abandoned', $abandoned)" in code


def test_44_the_annual_state_is_required_to_be_current() -> None:
    """AN ANNUAL RUN THAT PUBLISHED NOTHING would still return quickly."""
    code = _code()
    assert "if ($distribution -ne 'CURRENT') {" in code


def test_45_the_dimensions_are_read_back_out_of_the_workbook() -> None:
    """WHAT IS TIMED IS WHAT THE WORKBOOK HOLDS, not what the model intended. A
    fixture that silently added 299 drivers would otherwise be reported as 300."""
    code = _code()
    assert "$costIds = @(Get-IdColumnValues -Workbook $wb -Info $costRegister)" in code
    assert "([int]$actual['drivers']) -ne ([int]$scenarioSpec.drivers)" in code


# ===========================================================================
# F. IT MEASURES PRODUCTION, AND NOTHING ELSE
# ===========================================================================
def test_50_every_command_operation_is_a_declared_production_entry_point() -> None:
    """REQUIRED CONTROL 15. The endpoints are checked against the structure
    contract's own list, so a benchmark of a procedure that is not a user
    command cannot be declared."""
    declared = set(yaml.safe_load(
        (SPEC / "structure_contract.yaml").read_text(encoding="utf-8"))["vba"]["entry_points"])
    commands = [entry for entry in _plan()["operations"] if entry["kind"] == "command"]
    assert [entry["endpoint"] for entry in commands] == [
        "PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
        "PCCM_RunAnnualStochastic"]
    for entry in commands:
        assert entry["endpoint"] in declared, entry["endpoint"]
    # AND THE ONE OPERATION THAT IS NOT A COMMAND SAYS SO.
    recalculation = _operation("recalculation")
    assert recalculation["kind"] == "recalculation"
    assert recalculation["endpoint"] is None
    assert recalculation["mechanism"] == "Application.CalculateFull"


def test_51_the_runner_invokes_those_endpoints_by_name_from_the_plan() -> None:
    code = _code()
    assert "$Excel.Run([string]$Operation.endpoint)" in code
    for endpoint in ("PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
                     "PCCM_RunAnnualStochastic"):
        # The endpoint names live in the PLAN. The only one the runner spells is
        # Calculate, which it uses to re-establish the basis between iteration
        # counts - setup, outside every clock.
        if endpoint != "PCCM_Calculate":
            assert endpoint not in code, f"the runner hard-codes {endpoint}"
    assert "-Operation 'PCCM_Calculate'" in code


def test_52_the_harness_implements_no_business_algorithm_of_its_own() -> None:
    """REQUIRED CONTROL 16. Nothing here re-implements sampling, ranking,
    percentiles, correlation or the annual replay. If it did, the benchmark
    would be measuring the harness."""
    code = _code()
    for banned in ("Spearman", "spearman", "rank correlation", "percentile",
                   "Percentile", "Triangular sample", "Beta-PERT sample",
                   "MonteCarlo", "iterate", "[Math]::Sqrt", "cumulative"):
        assert banned not in code, f"the harness implements {banned}"
    # THE ONE ARITHMETIC IT DOES is the median of three numbers and the weights
    # that must sum to one; both are named and neither is business logic.
    maths = re.findall(r"\[Math\]::(\w+)", code)
    assert set(maths) <= {"Floor", "Round"}, maths


def test_53_no_timing_code_was_added_to_production_vba() -> None:
    for path in sorted(SRC_VBA.glob("*.bas")):
        source = path.read_text(encoding="utf-8")
        for banned in ("Timer", "GetTickCount", "QueryPerformanceCounter", "Stopwatch"):
            assert banned not in source, f"{path.name} carries {banned}"


def test_54_no_production_source_changed_since_the_accepted_tree() -> None:
    """REQUIRED CONTROL 17. This batch adds a way to measure the release. It
    does not change one line of what it measures."""
    changed = [line for line in _git("diff", "--name-only", ACCEPTED, "--",
                                     "pccm/src", "pccm/spec").splitlines() if line.strip()]
    assert changed == [], f"production source changed: {changed}"


def test_55_no_builder_owner_changed_except_the_new_plan_and_its_wiring() -> None:
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/builder").splitlines()
               if line.strip()}
    assert changed <= {"benchmark.py", "__init__.py", "build_stage_a.py"}, sorted(changed)


def test_56_no_accepted_windows_harness_changed() -> None:
    """THE ACCEPTED HARNESSES ARE LEFT ALONE. The benchmark carries its own copy
    of the COM primitives precisely so none of them had to be edited."""
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/bootstrap").splitlines()
               if line.strip()}
    assert changed <= {"phase10_benchmark.ps1"}, sorted(changed)


def test_57_the_copied_primitives_are_verbatim() -> None:
    """AND THE COPY IS PROVED TO BE ONE. A drifted copy would be a second, worse
    implementation of the same COM access, and the drift would show up as a
    benchmark that behaved differently from every accepted scenario."""
    accepted = TIMING.read_text(encoding="utf-8")
    runner = _runner()
    start = accepted.index("function Write-RowObject")
    end = accepted.index("# Phase-7 sensitivity block") if False else accepted.index(
        "$script:Phase7SensitivityGeometry")
    block = accepted[start:end]
    functions = re.findall(r"^function ([\w-]+) \{", block, re.M)
    assert len(functions) == 10, functions
    for name in functions:
        body = re.search(rf"^function {re.escape(name)} \{{.*?^}}", block, re.S | re.M)
        assert body, name
        assert body.group(0) in runner, f"{name} drifted from the accepted copy"


# ===========================================================================
# G. THE ARTIFACT, AND THE POLICY IT CARRIES
# ===========================================================================
def test_60_the_artifact_schema_is_deterministic_and_versioned() -> None:
    """REQUIRED CONTROL 18. Ordered keys, a declared schema version, and a
    runner that refuses a plan it does not understand."""
    assert _plan()["schema_version"] == BENCHMARK_SCHEMA_VERSION == 1
    code = _code()
    assert "$script:BenchmarkSupportedSchema = 1" in code
    assert "-ne $script:BenchmarkSupportedSchema" in code
    assert "OrderedDictionary" in code, "the artifact is built from unordered hashtables"
    assert "ConvertTo-Json -Depth 12" in code
    for key in ("schema_version", "baseline_id", "environment", "scenario",
                "scenario_actual", "setup_ms", "timing", "correctness_gates",
                "regression_policy", "historical_context", "forbidden", "results",
                "shutdown"):
        assert f"$report.Add('{key}'" in code, key


def test_61_both_a_machine_readable_and_a_human_readable_artifact_are_emitted() -> None:
    code = _code()
    assert "'.json'" in code and "'.md'" in code
    assert "Set-Content -LiteralPath $jsonPath" in code
    assert "Set-Content -LiteralPath $mdPath" in code
    assert "| Operation | Iterations | Cold | Warm 1 | Warm 2 | Warm 3 |" in _code()


def test_62_no_benchmark_number_is_written_into_the_workbook() -> None:
    """A CELL HOLDING A TIMING WOULD BE A STATE AUTHORITY NOBODY ASKED FOR, and
    the next reader would have no way to tell it from a model output."""
    code = _code()
    assert ".Save(" not in code
    assert "$wb.Close($false)" in code
    assert "SaveAs" not in code
    assert "$shutdownRecord.Add('workbook_saved', $false)" in code
    # The only writes into the workbook are the fixture's, through the accepted
    # helpers, and the two declared simulation controls.
    assert "Set-NamedValue -Workbook $wb" in code
    assert code.count("Set-NamedValue -Workbook $wb") == 2


def test_63_the_first_run_carries_no_absolute_pass_mark() -> None:
    """REQUIRED CONTROL 19. There is nothing to compare against yet, so a
    threshold invented now would make the first run pass or fail for a reason
    nobody could defend."""
    policy = _plan()["regression_policy"]
    assert policy["baseline_first"] is True
    assert policy["absolute_threshold_seconds"] is None
    assert "No absolute pass/fail threshold is contracted" in policy["absolute_threshold_note"]
    assert "does not judge" in _plan()["purpose"]

    code = _code()
    assert "This run RECORDS a baseline. It does not judge one" in code
    # AND NO SECOND-COUNT ANYWHERE IS A BAR TO CLEAR.
    for banned in ("BudgetSeconds", "MaxSeconds", "-gt $threshold", "TooSlow",
                   "PASS", "FAIL|", "-le $limit"):
        assert banned not in code, f"the runner carries a pass mark: {banned}"


def test_64_the_future_regression_ratios_are_recorded_exactly() -> None:
    """REQUIRED CONTROL 20."""
    policy = _plan()["regression_policy"]
    assert policy["investigate_ratio"] == 1.5
    assert policy["blocking_ratio"] == 2.0
    rules = " | ".join(policy["rules"])
    assert "at or below 1.5x the accepted baseline: no automatic regression finding" in rules
    assert "above 1.5x: investigate" in rules
    assert "above 2.0x: blocking regression unless explained and explicitly accepted" in rules
    assert "warm median against warm median" in policy["applies_to"]

    settled = re.sub(r"\s+", " ", RECORD.read_text(encoding="utf-8"))
    assert "**1.5×** a comparable warm baseline → investigate" in settled
    assert "**2×** a comparable warm baseline → **blocks delivery**" in settled


def test_65_the_phase_seven_timings_are_context_and_never_a_threshold() -> None:
    """REQUIRED CONTROL 21."""
    context = _plan()["historical_context"]
    assert "NOT a threshold" in context["status"]
    assert "NOT a baseline" in context["status"]
    assert context["source"] == "docs/phase7_closure.md, the operator's report"
    assert len(context["not_a_threshold_because"]) == 3
    observed = {(entry["drivers"], entry["iterations"]): entry["seconds"]
                for entry in context["observations"]}
    assert observed == {(20, 10_000): "6.4-7.6", (100, 10_000): "31.4",
                        (300, 10_000): "105.4"}, observed
    # THEY ARE NEVER COMPARED AGAINST. The runner prints them and does no
    # arithmetic with them.
    code = _code()
    assert "historical_context" in code
    assert "6.4" not in code and "105.4" not in code, (
        "the runner carries a historical number as a literal")


def test_66_the_baseline_identifier_is_stable() -> None:
    assert _plan()["baseline_id"] == "PCCM-P10-BASELINE-1"
    assert "$report.Add('baseline_id', [string]$plan.baseline_id)" in _code()


# ===========================================================================
# H. EXCEL LIFECYCLE
# ===========================================================================
def test_70_the_accepted_com_shutdown_path_is_used_whole() -> None:
    """REQUIRED CONTROL 22. Leaf before parent, a named release for every
    object, a wait for a natural exit, and emergency cleanup ONLY for a process
    whose identity is still positively verified."""
    code = _code()
    for step in ("New-ReleaseLedger", "Invoke-NamedRelease $rel $wb        'Workbook'",
                 "Invoke-NamedRelease $rel $workbooks 'Workbooks'",
                 "Invoke-NamedRelease $rel $excel 'Application'",
                 "$excel.Quit()", "Wait-ExcelExit -Identity $excelIdentity",
                 "Invoke-EmergencyExcelCleanup -Identity $excelIdentity",
                 "Format-ReleaseLedger $rel", "Get-TransientFailures"):
        assert step in code, step
    assert "} finally {" in code, "shutdown is not in a finally block"
    assert "[System.GC]::Collect()" in code


def test_71_the_instance_is_owned_and_no_foreign_process_is_touched() -> None:
    code = _code()
    assert "$preExisting = @(Get-PreExistingExcelPids)" in code
    assert "Get-ExcelIdentity -ExcelApp $excel -PreExistingPids $preExisting" in code
    assert "Stop-Process" not in code, "the runner kills a process directly"
    assert "taskkill" not in code.lower()
    # The one cleanup path is the accepted one, which verifies identity itself.
    assert "function Test-IsOurExcelProcess" in LIFECYCLE.read_text(encoding="utf-8")


def test_72_the_shutdown_outcome_reaches_the_artifact() -> None:
    """SO A RUN THAT LEFT A ZOMBIE SAYS SO IN THE EVIDENCE, rather than only in
    a console line nobody kept."""
    code = _code()
    for key in ("workbook_closed", "quit_called", "natural_exit",
                "emergency_required", "failed_releases", "transient_release_failures"):
        assert f"$shutdownRecord.Add('{key}'" in code, key
    assert "$report.Add('shutdown', $shutdownRecord)" in code


def test_73_the_real_build_output_is_never_opened_or_mutated() -> None:
    code = _code()
    assert "$tempRoot = Join-Path $WorkDir" in code
    assert "$stageBPath = Join-Path $tempRoot" in code
    assert "Copy-Item -LiteralPath (Join-Path $BuildDir" in code


def test_74_the_run_refuses_an_unattributable_workbook() -> None:
    """A MEASUREMENT NOBODY CAN ATTRIBUTE TO A REVISION IS NOT EVIDENCE, and the
    refusal happens before Excel is started rather than after an hour of work."""
    code = _code()
    assert "function Get-BenchmarkSourceRevision" in code
    # AND IT IS ACTUALLY CALLED. A refusal path that is defined and never
    # reached is a comment with a function keyword in front of it.
    assert "$revision = Get-BenchmarkSourceRevision -RepoRoot $repoRoot" in code
    assert "if (@($revision.Dirty).Count -gt 0) {" in code
    assert "'pccm/src', 'pccm/spec', 'pccm/builder'" in code
    assert "REFUSED, BEFORE EXCEL WAS STARTED." in _code()
    assert "exit 1" in code


def test_75_one_scenario_per_invocation() -> None:
    """THESE RUNS ARE LONG, and one command at a time is the point: the operator
    can stop after any of them and still hold complete evidence for the ones
    that finished."""
    code = _code()
    assert "[ValidateSet('PERF-SMALL', 'PERF-MEDIUM', 'PERF-LARGE')]" in code
    assert "[Parameter(Mandatory = $true)]" in code
    assert "[string]$Scenario," in code


# ===========================================================================
# I. THE PLAN IS BUILT, AND THE RUNNER READS IT
# ===========================================================================
@needs_build
def test_80_stage_a_emits_the_plan() -> None:
    emitted = json.loads(PLAN_FILE.read_text(encoding="utf-8"))
    assert emitted == _plan(), "the emitted plan is not what the builder builds"


def test_81_the_runner_reads_the_plan_and_declares_no_matrix() -> None:
    code = _code()
    assert "phase10_benchmark_plan.json" in code
    assert "$plan          = Get-Content -LiteralPath $planPath       -Raw | ConvertFrom-Json" in code
    assert "foreach ($run in @($plan.runs))" in code
    assert "foreach ($candidate in @($plan.scenarios))" in code
    # NO ITERATION COUNT AND NO DRIVER COUNT IS SPELLED IN THE RUNNER.
    for literal in ("10000", "50000", "100000", "10,000", "100,000"):
        assert literal not in code, f"the runner spells the iteration count {literal}"


def test_82_the_build_emits_it_beside_the_other_projections() -> None:
    driver = (BUILDER / "build_stage_a.py").read_text(encoding="utf-8")
    assert "emit_benchmark_plan(out_path.parent / \"phase10_benchmark_plan.json\")" in driver
    assert "It measures nothing" in driver


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
