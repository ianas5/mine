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
    RELEASE_IDENTITY_AUTHORITIES,
    build_benchmark_plan,
    cost_line_count,
)
from pccm_builder.spec_loader import load_spec  # noqa: E402
from pccm_builder.workbook_builder import BUILDER_VERSION  # noqa: E402

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


def _spec():
    if "spec" not in _MEMO:
        _MEMO["spec"] = load_spec(SPEC / "workbook.yaml")
    return _MEMO["spec"]


def _plan() -> dict:
    if "plan" not in _MEMO:
        _MEMO["plan"] = build_benchmark_plan(_spec())
    return _MEMO["plan"]


def _manifest_text() -> str:
    if "manifest_text" not in _MEMO:
        _MEMO["manifest_text"] = (SPEC / "workbook.yaml").read_text(encoding="utf-8")
    return _MEMO["manifest_text"]


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
    # AND A PATH THE HARNESS COULD NOT CLASSIFY IS 'unknown', NEVER 'local'.
    # Falling through to 'local' after a failed drive lookup would state the one
    # thing the contract asked never to be assumed: a synced repository silently
    # recorded as an unsynchronised path.
    classifier = code.split("function Get-BenchmarkLocationType")[1]
    classifier = classifier[:classifier.index("\nfunction ")]
    failure_branch = classifier.rsplit("} catch {", 1)[1]
    failure_branch = failure_branch[:failure_branch.index("\n    }")]
    assert "return 'unknown'" in failure_branch, failure_branch
    assert "return 'local'" not in failure_branch, failure_branch
    assert "'unknown'" in _plan()["location_types"] or "unknown" in _plan()["location_types"]
    # AND NO CLAIM IS MADE ABOUT WHAT ANY OF IT DOES. The harness records where
    # the file is; it does not theorise about sync overhead.
    assert "slower" not in _runner().lower()
    assert "no claim is made" in _runner().lower()


def test_23_the_commit_and_the_release_metadata_are_captured() -> None:
    """REQUIRED CONTROL 11.

    THIS CONTROL WAS PART OF THE W2 DEFECT. It used to assert
    `"$Manifest.builder_version" in code` - it quoted the implementation's own
    expression back at itself, so it could only ever prove the runner was
    self-consistent. Nothing checked that expression against the SHAPE OF THE
    ARTIFACT it reads, and the manifest has never carried a builder version.

    What it asserts now is the authority path, and `test_110` checks every
    property the runner dereferences against the real emitted JSON.
    """
    fields = _plan()["environment_fields"]
    for field in ("git_branch", "git_commit", "git_worktree_clean",
                  "model_version", "builder_version", "build_phase",
                  "harness_version", "schema_version"):
        assert field in fields, field
    code = _code()
    assert "rev-parse --abbrev-ref HEAD" in code
    assert "$Manifest.builder_version" not in code, "the W2 defect is back"
    assert "$Manifest.build_phase" not in code, "the W2 defect's twin is back"
    for field in ("model_version", "builder_version", "build_phase"):
        assert (f"-InputObject $ReleaseIdentity -Name '{field}'") in code, field


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


def _production_changed_since(commit: str) -> list[str]:
    """Production paths that changed since `commit`, EXCLUDING the declared
    P10-RP structural window - and only after proving each of those reverses to
    that same commit's bytes exactly.

    P10-RP. Windows disproved the accepted Phase-10 assumption that
    UserInterfaceOnly:=True permits ListObject structural mutation on a protected
    sheet: PCCM_ApplyTimeline was invoked and refused with Error 1004. Six
    production modules gained the structural window under their own
    authorisation, LATER than the harness batches these controls belong to.
    Deleting the controls would lose the claim they exist to make; naming the six
    and proving the reversal keeps it, and keeps it in the stronger form - the
    tree these harness batches measured is still exactly recoverable.
    """
    import sys as _sys
    _sys.path.insert(0, str(PCCM_ROOT / "tests"))
    from vba_structural_window import (DECLARED_STRUCTURAL_WINDOW_CHANGES,
                                       strip_structural_window)

    declared = {f"pccm/src/vba/{name}" for name in DECLARED_STRUCTURAL_WINDOW_CHANGES}
    changed = [line for line in _git("diff", "--name-only", commit, "--",
                                     "pccm/src", "pccm/spec").splitlines() if line.strip()]
    for path in sorted(set(changed) & declared):
        name = Path(path).name
        # LINE ENDINGS NORMALISED ON BOTH SIDES. modCalcReport.bas is CRLF on
        # disk and _git returns text, so one side arrives translated. This
        # control is about CONTENT; the line-ending convention has its own
        # control (test_phase9_model_check test_46_2) and keeps it.
        current = (PCCM_ROOT / "src" / "vba" / name).read_bytes().decode("utf-8")
        current = strip_structural_window(name, current).replace("\r\n", "\n")
        accepted = _git("show", f"{commit}:{path}").replace("\r\n", "\n")
        assert current == accepted, (
            f"{path} moved outside the declared P10-RP structural window")
    return [path for path in changed if path not in declared]


def test_54_no_production_source_changed_since_the_accepted_tree() -> None:
    """REQUIRED CONTROL 17. This batch adds a way to measure the release. It
    does not change one line of what it measures."""
    changed = _production_changed_since(ACCEPTED)
    assert changed == [], f"production source changed: {changed}"


def test_55_no_builder_owner_changed_except_the_new_plan_and_its_wiring() -> None:
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/builder").splitlines()
               if line.strip()}
    assert changed <= {"benchmark.py", "__init__.py", "build_stage_a.py"}, sorted(changed)


# W3/W4. DECLARED, NOT EXEMPTED. Each name here was added because a batch had a
# reason a reviewer accepted, and the reason is written beside it. A file that is
# not on this list still cannot change.
CHANGED_BY_DECLARATION = {
    # W3. The protection probe is a NEW file added by a later batch to ask Excel
    # one question. It is not an accepted harness and it changes none.
    "phase10_protection_probe.ps1",
    # The benchmark harness this suite is about.
    "phase10_benchmark.ps1",
    # W4. Stage-B verification was failing BEFORE the probe could run: Excel
    # refused an incoming COM read with RPC_E_CALL_REJECTED on the first sheet of
    # the reopened workbook, twice in a row. com_lifecycle.ps1 gained the bounded
    # read-boundary retry that settles it, and build_stage_b.ps1's verification
    # block now reads through that helper. Neither is a scenario harness.
    "com_lifecycle.ps1",
    "build_stage_b.ps1",
}

# The scenario harnesses this control exists to protect. Named, so the control
# says what it defends rather than only what it forbids.
FROZEN_HARNESSES = (
    "phase4_functional_test.ps1",
    "phase5_gate_b_scenarios.ps1",
    "phase6_gate_b_scenarios.ps1",
    "phase7_acceptance_scenarios.ps1",
    "phase7_timing_scenarios.ps1",
    "phase7_w1_smoke.ps1",
    "phase7_w2_many_drivers.ps1",
    "phase7_w3_long_years.ps1",
    "phase7_w4_base_simulation.ps1",
    "phase7_w5_annual_success.ps1",
    "phase7_w6_selector_move.ps1",
    "phase7_w7_bank_cycle.ps1",
    "phase7_w8_refusal.ps1",
    "phase8_p1_results_surface.ps1",
    "phase8_p2_dashboard_surface.ps1",
    "phase8_p3_chart_surface.ps1",
    "phase8_pz_zero_variance.ps1",
    "phase9_p1_model_check.ps1",
)


def test_56_no_accepted_windows_harness_changed() -> None:
    """THE ACCEPTED SCENARIO HARNESSES ARE LEFT ALONE. The benchmark carries its
    own copy of the COM primitives precisely so none of them had to be edited."""
    changed = {Path(line).name for line in
               _git("diff", "--name-only", ACCEPTED, "--", "pccm/bootstrap").splitlines()
               if line.strip()}
    assert changed <= CHANGED_BY_DECLARATION, sorted(changed - CHANGED_BY_DECLARATION)


def test_56a_every_frozen_scenario_harness_is_byte_identical() -> None:
    """SAID POSITIVELY, NOT BY SUBTRACTION.

    test_56 widened twice by declaration, and a set that only ever grows stops
    defending anything the day someone adds one more name to it. This states the
    thing that must remain true directly: every scenario harness is unchanged
    since the accepted commit. Widening the declaration above cannot weaken this.
    """
    for name in FROZEN_HARNESSES:
        assert (BOOTSTRAP / name).is_file(), f"{name} is named as frozen but is not on disk"
        diff = _git("diff", "--name-only", ACCEPTED, "--", f"pccm/bootstrap/windows/{name}")
        assert not diff.strip(), f"{name} changed since {ACCEPTED}"


def test_56b_the_declaration_and_the_freeze_do_not_overlap() -> None:
    """A NAME CANNOT BE ON BOTH LISTS. If a scenario harness were quietly added to
    the declaration, test_56a would still fail -- but only after it had already
    been edited. This refuses the declaration itself."""
    overlap = CHANGED_BY_DECLARATION & set(FROZEN_HARNESSES)
    assert not overlap, sorted(overlap)
    # And the freeze must cover every scenario harness actually on disk, so a new
    # one cannot be added and silently left unprotected.
    on_disk = {p.name for p in BOOTSTRAP.glob("*.ps1")}
    unaccounted = on_disk - set(FROZEN_HARNESSES) - CHANGED_BY_DECLARATION
    assert not unaccounted, sorted(unaccounted)


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
    assert ('emit_benchmark_plan(out_path.parent / "phase10_benchmark_plan.json", spec)'
            in driver), "the plan is emitted without the spec that carries the release"
    assert "It measures nothing" in driver


# ===========================================================================
# J. WINDOWS RUN 1 - THE DEFECT THAT ABORTED IT
# ===========================================================================
# WHAT HAPPENED. On PowerShell 5.1, on a machine with a consumer OneDrive and no
# work account, `Get-Item Env:OneDriveCommercial -ErrorAction SilentlyContinue`
# emitted NOTHING. The parenthesised pipeline was therefore $null, and under
# `Set-StrictMode -Version 2.0` reading `.Value` off $null is a terminating
# PropertyNotFoundException (FullyQualifiedErrorId PropertyNotFoundStrict), not
# a quiet $null. The run died in the environment capture, before a single
# operation was timed.
#
# WHY THE EXISTING CONTROLS DID NOT CATCH IT. Every one of them reads the
# runner as TEXT. Text controls can prove what a script says; they cannot prove
# what a cmdlet returns on a machine this repository never sees. So the controls
# below assert the PROPERTY THEY CAN: that no property is read off a value that
# a pipeline might not have produced, anywhere in the harness.

# The shapes a PowerShell 5.1 pipeline can hand back where the harness reads a
# property: nothing at all, one object, or several. Only the first is fatal, and
# it is the only one that never occurs on the developer's machine.
_PIPELINE_SHAPES = ("no output at all (an absent Env: entry, an unmatched CIM query)",
                    "exactly one object",
                    "a collection")


def test_90_no_property_is_read_off_an_unguarded_pipeline_result() -> None:
    """THE W1 DEFECT, AS A PROPERTY RATHER THAN AS A LINE NUMBER.

    `(<pipeline>).Property` is the shape that killed Windows Run 1: it reads a
    member off whatever the pipeline produced, and "nothing" is one of the
    things a pipeline produces. Under StrictMode 2.0 that is a terminating
    error, so the harness may not contain the shape at all.
    """
    # NAMED, NOT COUNTED. `Get-Date` with no arguments cannot emit nothing - it
    # is not a LOOKUP, it is a value - so a method call on it is not the defect
    # class. Every other command member-read has to be justified here or the
    # control refuses it, which is what makes a new one a decision rather than
    # an oversight.
    allowed = ("Get-Date",)
    offenders = [entry for entry in _command_member_reads(_code())
                 if not entry.startswith(allowed)]
    assert offenders == [], (
        "a member is read directly off a command's output, which is $null when "
        f"the command emits nothing: {offenders}")
    assert "-ErrorAction SilentlyContinue).Value" not in _code(), (
        "the Windows Run 1 defect is back verbatim")
    # AND THE DETECTOR IS NOT VACUOUS. The exact statement that aborted Windows
    # Run 1 is fed to it here, so a control that stopped detecting anything
    # fails rather than passing over a clean file.
    w1 = ("        $value = [string](Get-Item -LiteralPath ('Env:' + $name) "
          "-ErrorAction SilentlyContinue).Value")
    assert _command_member_reads(w1), "the detector no longer detects the W1 defect"
    assert _command_member_reads("@(Get-Item -LiteralPath 'Env:X').Count") == [], (
        "the detector flags the normalised form the fix introduced")


# Verb-Noun is how PowerShell spells a command, and a command is the only thing
# in an expression that can legitimately produce NOTHING. A member read off a
# string, a cast or a method result cannot be $null-shaped; a member read off
# `(Get-Something ...)` can, and that is the whole defect class.
_COMMAND = re.compile(r"^\s*[&]?\s*(?:Get|Set|New|Select|Where|ForEach|Measure|Sort|"
                      r"Import|Export|Test|Start|Stop|Wait|Copy|Remove|Join|Split|"
                      r"Convert|ConvertTo|ConvertFrom|Out|Write|Add|Invoke)-")


def _command_member_reads(code: str) -> list[str]:
    """Every `(<command ...>).<member>` in the source, by the member read."""
    found: list[str] = []
    for match in re.finditer(r"\)\s*\.(\w+)", code):
        # Walk back to the matching open parenthesis so a nested expression is
        # attributed to the expression it actually closes.
        depth, index = 0, match.start()
        while index >= 0:
            if code[index] == ")":
                depth += 1
            elif code[index] == "(":
                depth -= 1
                if depth == 0:
                    break
            index -= 1
        if index < 0:
            continue
        # `@(<command>).<member>` IS THE FIX, NOT THE DEFECT. The array
        # subexpression turns "emitted nothing" into an empty collection before
        # anything is read off it, which is exactly the normalisation this round
        # introduced. Only a BARE `(<command>).<member>` is the W1 shape.
        if index > 0 and code[index - 1] == "@":
            continue
        inner = code[index + 1:match.start()]
        if _COMMAND.match(inner):
            found.append(f"{inner.strip()[:48]}....{match.group(1)}")
    return found


def test_91_the_environment_provider_is_normalised_before_it_is_read() -> None:
    """NORMALISE, VALIDATE, THEN USE. `@()` turns "emitted nothing" into an
    empty collection, and the loop body then runs zero times instead of reading
    a member off $null."""
    code = _code()
    assert ("foreach ($item in @(Get-Item -LiteralPath ('Env:' + $name) "
            "-ErrorAction SilentlyContinue)) {") in code
    assert "function Get-BenchmarkProperty" in code
    assert "$property = $InputObject.PSObject.Properties[$Name]" in code
    assert "if ($null -eq $property) { return $null }" in code
    # AND THE CIM QUERIES ARE NORMALISED THE SAME WAY, for the same reason: a
    # query that matches nothing is an empty pipeline, not a $null object.
    for cls in ("Win32_OperatingSystem", "Win32_Processor", "Win32_ComputerSystem",
                "Win32_LogicalDisk"):
        assert f"@(Get-CimInstance -ClassName {cls}" in code, cls


def test_92_the_access_path_is_powershell_5_1_compatible() -> None:
    """THE TARGET IS WINDOWS POWERSHELL 5.1, and the syntax that would have made
    this a one-liner does not exist there."""
    code = _code()
    for later_only in ("??", "?.", "-AsHashtable", "ForEach-Object -Parallel",
                       "$PSStyle", "-AsByteStream", "System.Text.Json",
                       "using namespace", "&&", "||"):
        assert later_only not in code, f"{later_only} is not Windows PowerShell 5.1"
    # `PSObject.Properties[...]` and `@(...)` are both 5.1, and so is every
    # cmdlet the harness calls.
    assert "PSObject.Properties[" in code
    assert "Get-CimInstance" in code


def test_93_strict_mode_stays_on() -> None:
    """THE FIX IS NOT TO STOP CHECKING. StrictMode is what turned a silent $null
    into a loud failure, and turning it off would have hidden the defect rather
    than removed it - the environment record would have carried empty strings
    and nobody would have known the machine was never asked."""
    code = _code()
    assert "Set-StrictMode -Version 2.0" in code
    assert "Set-StrictMode -Off" not in code
    assert code.count("Set-StrictMode") == 1
    assert "$ErrorActionPreference = 'Stop'" in code


def test_94_a_missing_property_refuses_rather_than_defaulting() -> None:
    """NO FAKE DEFAULT. An absent optional fact is recorded as unavailable - a
    truthful statement about the machine - and a required one throws with the
    shape named."""
    code = _code()
    assert "function Get-BenchmarkRequiredProperty" in code
    assert "It is not defaulted." in code
    assert "function Format-BenchmarkFact" in code
    assert "if ($null -eq $Value) { return (Get-BenchmarkUnavailable) }" in code
    # AND NOTHING IS SWALLOWED WHOLESALE. Every catch in the harness either
    # records a named field as unavailable or is the outer handler that reports
    # the failure; none of them continues as if nothing happened.
    assert "On Error Resume Next" not in code
    assert "-ErrorAction Ignore" not in code
    assert "catch { }" not in code, "a bare swallowing catch is back"


def test_95_the_failure_diagnostics_name_the_stage() -> None:
    """REQUIRED CONTROL 8 OF THIS ROUND. Windows Run 1 reached the outer handler
    and said only what the exception said. It now says where it was."""
    code = _code()
    assert "$script:BenchmarkCursor" in code
    assert "function Set-BenchmarkStage" in code
    assert "function Set-BenchmarkOperationContext" in code
    assert "function New-BenchmarkFailureRecord" in code
    for field in ("stage", "action", "scenario", "operation", "iterations",
                  "execution_phase", "exception_type", "message",
                  "script_line_number", "script_line", "command"):
        assert f"$record.Add('{field}'" in code, field
    assert "$report.Add('failure', $failure)" in code
    assert "Format-BenchmarkFailure $failure" in code


def test_96_every_setup_stage_sets_the_cursor() -> None:
    """SO NO STAGE CAN FAIL ANONYMOUSLY. The stages between the bootstrap and
    the first timed call are exactly where Windows Run 1 died."""
    code = _code()
    stages = re.findall(r"Set-BenchmarkStage -Stage '(\w+)'", code)
    assert set(stages) == {"preflight", "setup", "scenario", "measurement"}, stages
    assert stages.count("setup") >= 4, stages
    for action in ("capturing the environment inventory",
                   "opening the benchmark workbook",
                   "starting an owned Excel instance",
                   "running the Stage-B bootstrap"):
        assert action in code, action


def test_97_no_diagnostic_work_happens_inside_a_measured_interval() -> None:
    """THE CURSOR COSTS NOTHING A TIMING CAN SEE. It is set before the call and
    read only by the failure path; `test_15` proves the interval still holds one
    statement, and this proves the diagnostics are not that statement."""
    body = _code().split("function Invoke-BenchmarkExecution")[1]
    body = body[:body.index("$evidence = ")]
    for banned in ("Set-BenchmarkStage", "Set-BenchmarkOperationContext",
                   "New-BenchmarkFailureRecord", "Write-BenchmarkLine"):
        assert banned not in body, f"{banned} runs inside a timed execution"


def test_98_a_setup_failure_cannot_become_a_valid_sample() -> None:
    """REQUIRED CONTROL 5 OF THIS ROUND. Windows Run 1 produced no timed
    operation at all; a run like it must not be able to look like evidence."""
    code = _code()
    assert "$runComplete = ([bool](([string]::IsNullOrWhiteSpace($abandoned)) -and" in code
    assert "@($completed).Count -eq $plannedCount" in code
    assert "$baselineEstablished = ([bool]($runComplete -and (-not $scoped)))" in code
    assert "$report.Add('baseline_established', $baselineEstablished)" in code
    assert "ABORTED BEFORE A COMPLETE BASELINE" in code
    assert "NOT a partial warm median" in code
    # AND THE PROCESS SAYS SO TOO.
    assert "if (-not $runComplete) {" in code
    assert "exit 1" in code


def test_99_a_completed_run_is_only_a_baseline_when_it_was_not_narrowed() -> None:
    """A -Operations calculate RUN IS COMPLETE FOR WHAT IT ASKED and is not the
    scenario's baseline. The two are different claims and the artifact makes
    both."""
    code = _code()
    assert "SCOPED RUN COMPLETE" in code
    assert "NOT the scenario baseline" in code
    # `@($null).Count` is 1, so an unsupplied parameter must be null-checked
    # first or every full run would be reported as scoped.
    assert "($null -ne $Iterations) -and (@($Iterations).Count -gt 0)" in code


def test_100_the_bootstrap_time_is_never_a_user_operation_figure() -> None:
    """REQUIRED CONTROL 7 OF THIS ROUND. Windows Run 1's only number was a
    68.4 s Stage-B bootstrap, and it is setup under every reading."""
    code = _code()
    assert "$setupTimings.Add('stage_b_bootstrap_ms'" in code
    assert "a Stage-B bootstrap " in code and "never a baseline" in code
    assert "stage_b_bootstrap_ms" not in code.split("$row.Add('cold_ms'")[1][:400]
    assert "Stage-B bootstrap" in _plan()["timing"]["excluded_from_elapsed"]


def test_101_the_excel_bitness_is_excels_and_not_the_hosts() -> None:
    """CORRECTED IN THE SAME ROUND, AND NO EVIDENCE IS INVALIDATED because
    Windows Run 1 never reached the line. `Is64BitProcess` is the POWERSHELL
    host's bitness; a 64-bit host automating a 32-bit Excel would have been
    recorded as 64-bit Excel, in the one field a reader uses to say which build
    was exercised."""
    code = _code()
    assert "Is64BitProcess" not in code, "the host's bitness is back"
    assert "function Get-BenchmarkExcelImage" in code
    assert "Get-Process -Id $processId" in code
    assert "derived from the Excel image path" in code
    assert "excel_executable_path" in _plan()["environment_fields"]
    assert "$record.Add('excel_executable_path', $excelPath)" in code


# ===========================================================================
# K. WINDOWS RUN 2 - THE AUTHORITY THAT WAS NEVER THERE
# ===========================================================================
# WHAT HAPPENED. The runner asked `stage_b_manifest.json` for `builder_version`.
# That manifest is a projection of the MODEL side of the specification and has
# never carried a builder version - it cannot, because P10-3 settled that the
# model version and the builder version are INDEPENDENT authorities and the
# manifest answers for only one of them. Under StrictMode 2.0 reading a property
# that is not there is a terminating PropertyNotFoundException.
#
# THE FIX IS A PROJECTION, NOT A PATCH:
#
#     BUILDER_VERSION (workbook_builder.py)
#       -> imported by builder/pccm_builder/benchmark.py
#       -> release_identity in build/phase10_benchmark_plan.json
#       -> the PowerShell runner
#       -> environment.builder_version
#
# and the two values that come from the manifest's own side travel the same way,
# each labelled with the file that answers for it.

_MANIFEST_FILE = BUILD / "stage_b_manifest.json"


def test_110_every_property_the_runner_reads_exists_on_the_artifact() -> None:
    """THE CONTROL THAT WOULD HAVE CAUGHT W2, AND CATCHES THE WHOLE CLASS.

    Text controls can prove what a script says. This one takes every property
    the runner dereferences on a loaded JSON artifact and looks it up in the
    ARTIFACT THAT STAGE A ACTUALLY EMITS - so a property that does not exist
    fails here rather than sixty-eight seconds into a Windows run.
    """
    documents = {
        "$plan": PLAN_FILE,
        "$manifest": _MANIFEST_FILE,
        "$simInspection": BUILD / "phase6_gate_b_inspection.json",
        "$inspection": BUILD / "phase5_gate_b_inspection.json",
    }
    missing: list[str] = []
    for variable, path in documents.items():
        if not path.is_file():
            pytest.skip(f"{path.name} has not been built")
        document = json.loads(path.read_text(encoding="utf-8"))
        for expression in sorted(set(re.findall(
                rf"{re.escape(variable)}((?:\.[A-Za-z_][A-Za-z0-9_]*)+)", _code()))):
            node = document
            for part in expression.lstrip(".").split("."):
                if isinstance(node, dict) and part in node:
                    node = node[part]
                    continue
                missing.append(f"{variable}{expression}")
                break
    assert missing == [], (
        "the runner reads properties that the emitted artifacts do not carry, which "
        f"is a terminating error under StrictMode: {missing}")

    # AND THE SCAN IS NOT VACUOUS. The first draft of this control carried a
    # regex that matched nothing at all and passed on every file it was given -
    # which is the same failure as the one it was written to catch, one level
    # up. So it proves it found real reads, and proves it would have found the
    # W2 defect.
    reads = re.findall(r"\$plan((?:\.[A-Za-z_][A-Za-z0-9_]*)+)", _code())
    assert len(reads) >= 20, f"the property scan found almost nothing: {reads}"
    for expected in (".schema_version", ".release_identity", ".timing.cold_runs",
                     ".regression_policy.rules"):
        assert expected in reads, expected
    manifest = json.loads(_MANIFEST_FILE.read_text(encoding="utf-8"))
    assert "builder_version" not in manifest, (
        "the scan would no longer catch the W2 defect, because the manifest now "
        "carries the property that was missing")


def test_111_the_manifest_is_not_asked_for_a_builder_version() -> None:
    """REQUIRED CONTROL 1 OF THIS ROUND, and it is checked against the emitted
    artifact rather than against a memory of what it contains."""
    if not _MANIFEST_FILE.is_file():
        pytest.skip("the Stage-B manifest has not been built")
    manifest = json.loads(_MANIFEST_FILE.read_text(encoding="utf-8"))
    assert "builder_version" not in manifest, (
        "the Stage-B manifest grew a builder version; the model side of the "
        "specification does not answer for the build tooling")
    assert "build_phase" not in manifest
    assert manifest["model_version"] == _spec().model["model_version"]
    assert "$Manifest.builder_version" not in _code()
    assert "$Manifest.build_phase" not in _code()


def test_112_the_builder_version_reaches_powershell_by_projection() -> None:
    """REQUIRED CONTROLS 2 AND 3. The runner consumes a generated artifact; it
    parses no Python and reads no source text."""
    identity = _plan()["release_identity"]
    assert identity["builder_version"] == BUILDER_VERSION
    assert identity["authorities"]["builder_version"] == (
        "builder/pccm_builder/workbook_builder.py: BUILDER_VERSION")

    benchmark = (BUILDER / "pccm_builder" / "benchmark.py").read_text(encoding="utf-8")
    assert "from .workbook_builder import BUILDER_VERSION" in benchmark, (
        "the plan no longer imports the builder version from its owner")
    assert '"builder_version": str(BUILDER_VERSION)' in benchmark

    code = _code()
    assert "workbook_builder.py" not in code, "the runner is reading Python source"
    assert "Select-String" not in code and "Get-Content" in code
    assert "$plan.release_identity" in code


def test_113_no_version_literal_is_restated_anywhere_downstream() -> None:
    """REQUIRED CONTROL 4. A second literal is a second authority the day one of
    them moves."""
    code = _code()
    for literal in ("1.0.0", "'1.0'", '"1.0"'):
        assert literal not in code, f"the runner restates a version: {literal}"
    benchmark = (BUILDER / "pccm_builder" / "benchmark.py").read_text(encoding="utf-8")
    assert 'BUILDER_VERSION = ' not in benchmark, (
        "the benchmark plan declares a builder version of its own")
    # THE HARNESS HAS A VERSION OF ITS OWN, and that is a fourth independent
    # thing rather than a restatement: it says which runner produced a result.
    # Every OTHER dotted literal in the module would be a copy of somebody
    # else's authority, so there are none.
    literals = [line.strip() for line in benchmark.splitlines()
                if re.search(r'=\s*"\d+\.\d+\.\d+"', line)]
    assert literals == ['HARNESS_VERSION = "1.0.0"'], literals
    assert '"builder_version": str(BUILDER_VERSION)' in benchmark


def test_114_the_three_release_values_stay_independent() -> None:
    """REQUIRED CONTROLS 5 AND 6. They all read 1.0.0 for this release, which is
    a coincidence of this release. Nothing derives one from another."""
    identity = _plan()["release_identity"]
    assert identity["model_version"] == _spec().model["model_version"]
    assert identity["builder_version"] == BUILDER_VERSION
    assert identity["build_phase"] == _spec().model["build_phase"]
    assert identity["authorities"] == RELEASE_IDENTITY_AUTHORITIES
    assert len(set(identity["authorities"].values())) == 3, identity["authorities"]

    benchmark = (BUILDER / "pccm_builder" / "benchmark.py").read_text(encoding="utf-8")
    for derived in ('"builder_version": str(spec', '"builder_version": values["model_version"]',
                    'builder_version = model_version'):
        assert derived not in benchmark, f"the builder version is derived: {derived}"

    # AND THE MANIFEST IS NOT GIVEN A SECOND BUILDER-VERSION AUTHORITY.
    declarations = [line for line in _manifest_text().splitlines()
                    if ("builder_version" in line or "BUILDER_VERSION" in line)
                    and not line.lstrip().startswith("#")]
    assert declarations == [], declarations


def test_115_a_plan_without_a_release_identity_is_refused_before_timing() -> None:
    """REQUIRED CONTROL 7, and the refusal happens BEFORE the bootstrap - W2
    spent sixty-eight seconds building a workbook it could not attribute."""
    code = _code()
    assert "$releaseIdentity = $null" in code
    assert "if ($null -eq $releaseIdentity) {" in code
    # THE GUARD, NOT ONLY ITS MESSAGE. A refusal sentence inside a branch that
    # can never be entered is a comment with a Write-Host in front of it.
    assert "if ($releaseProblems.Count -gt 0) {" in code
    assert "$releaseProblems += ('  ' + $field + ' is missing from the plan projection" in code
    assert "REFUSED, BEFORE ANYTHING WAS BUILT OR MEASURED." in code
    assert "is missing from the plan projection; its " in code
    assert "authority is " in code
    # THE CHECK IS AHEAD OF THE BOOTSTRAP IN THE FILE, so it cannot be reached
    # after the expensive part has already run.
    assert code.index("REFUSED, BEFORE ANYTHING WAS BUILT OR MEASURED.") <         code.index("$bootstrapWatch = [System.Diagnostics.Stopwatch]::StartNew()")


def test_116_a_missing_release_value_is_never_substituted() -> None:
    """REQUIRED CONTROL 7's other half. An optional environment fact may be
    'unavailable'; a release identity may not - not an empty string, not
    'unknown', not a default."""
    code = _code()
    for field in ("model_version", "builder_version", "build_phase"):
        assert ("Get-BenchmarkRequiredProperty `\n        -InputObject $ReleaseIdentity "
                f"-Name '{field}'") in code, field
    # NONE OF THE THREE GOES THROUGH THE OPTIONAL PATH.
    optional = code.split("$record.Add('model_version'")[1]
    optional = optional[:optional.index("$record.Add('harness_version'")]
    assert "Format-BenchmarkFact" not in optional, (
        "a release value can be recorded as unavailable")
    assert "Get-BenchmarkUnavailable" not in optional
    # AND THE BUILDER'S OWN PROJECTION REFUSES A BLANK BEFORE IT IS EMITTED.
    benchmark = (BUILDER / "pccm_builder" / "benchmark.py").read_text(encoding="utf-8")
    assert "a benchmark result may not carry an unidentified release" in benchmark


def test_117_the_environment_still_records_the_three_separately() -> None:
    """REQUIRED CONTROL 9. One projection, three values, three owners - never
    one value repeated three times."""
    for field in ("model_version", "builder_version", "build_phase"):
        assert field in _plan()["environment_fields"], field
    code = _code()
    assert code.count("$record.Add('model_version'") == 1
    assert code.count("$record.Add('builder_version'") == 1
    assert code.count("$record.Add('build_phase'") == 1
    assert "$report.Add('release_identity', $releaseIdentity)" in code, (
        "the artifact does not carry the projection it was built from")


def test_118_no_production_source_changed_since_the_w1_correction() -> None:
    """REQUIRED CONTROL 11 OF THIS ROUND. This is a harness batch."""
    changed = _production_changed_since("6e87fda")
    assert changed == [], f"production source changed: {changed}"


def test_119_the_matrix_and_the_endpoints_did_not_move() -> None:
    """REQUIRED CONTROL 10. A metadata correction may not change what is
    measured, at what sizes, or how many times."""
    accepted = json.loads(_git("show", "6e87fda:pccm/build/phase10_benchmark_plan.json")
                          ) if False else None
    plan = _plan()
    assert [(entry["id"], entry["drivers"], entry["years"], tuple(entry["iterations"]))
            for entry in plan["scenarios"]] == [
        ("PERF-SMALL", 20, 10, (10_000, 50_000, 100_000)),
        ("PERF-MEDIUM", 100, 25, (10_000, 50_000, 100_000)),
        ("PERF-LARGE", 300, 40, (10_000, 50_000))]
    assert len(plan["runs"]) == 30
    assert plan["timing"]["cold_runs"] == 1 and plan["timing"]["warm_runs"] == 3
    assert [entry["endpoint"] for entry in plan["operations"] if entry["kind"] == "command"] == [
        "PCCM_Calculate", "PCCM_RunSimulation", "PCCM_RunSensitivity",
        "PCCM_RunAnnualStochastic"]


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
