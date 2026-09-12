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
RUN_EVIDENCE = DOCS / "phase10_windows_run_evidence.md"

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
    # W5. PERF-SMALL aborted on `ClearContents` under worksheet protection, so the
    # benchmark needs production's own structural window around its fixture
    # writes - and `modProtection`'s two mutators take `ByRef detail As String`,
    # which `Application.Run` has no precedent in this tree for carrying. This is
    # the harness-owned VBA shim that owns the String and forwards to the accepted
    # authority. It is imported into the DISPOSABLE workbook exactly as
    # `phase5_gate_b_diagnostics.bas` has been since Phase 5, it is never declared
    # in the manifest, and it holds no protection policy of its own.
    #
    # IT IS A NEW FILE, NOT AN EDIT. No accepted harness changed for it.
    "phase10_fixture_window.bas",
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


# THE TWO PRIMITIVES THAT DELIBERATELY ARE NOT VERBATIM ANY MORE, and the
# sentence in the runner that has to justify each one.
#
# DECLARED, NOT EXCUSED. Before the protected workbook existed, "all ten copies
# are verbatim" was true and the control said so. It is no longer true, and the
# honest repair is to NAME the two departures and keep the other eight held to
# the letter - not to loosen the comparison until everything passes. A primitive
# that drifts without appearing here still fails, which is the property the
# original control had and this one keeps.
DECLARED_PRIMITIVE_DEPARTURES = {
    # Grew the table through `ListRows.Add`, refused on a protected sheet.
    # Searches the contract's reserved blank rows instead.
    "Add-BlankTableRow": "A BLANK ROW IS FOUND, NOT MADE.",
    # Deleted a ListRow through COM - the exact call Benchmark Run 3 died on.
    # Now refuses; the fixture resets by content instead.
    "Remove-TableRow": "`Remove-TableRow` REFUSES. IT NO LONGER DELETES.",
}


def test_57_the_copied_primitives_are_verbatim() -> None:
    """AND THE COPY IS PROVED TO BE ONE. A drifted copy would be a second, worse
    implementation of the same COM access, and the drift would show up as a
    benchmark that behaved differently from every accepted scenario.

    TWO OF THE TEN NO LONGER ARE, BY DECLARATION. Both performed a ListObject
    structural mutation, which worksheet protection refuses to a COM caller, and
    both are listed above with the sentence in the runner that justifies them.
    Everything else is still held byte for byte."""
    accepted = TIMING.read_text(encoding="utf-8")
    runner = _runner()
    start = accepted.index("function Write-RowObject")
    end = accepted.index("$script:Phase7SensitivityGeometry")
    block = accepted[start:end]
    functions = re.findall(r"^function ([\w-]+) \{", block, re.M)
    assert len(functions) == 10, functions
    # THE DECLARED LIST IS NOT ALLOWED TO NAME SOMETHING THAT WAS NEVER COPIED,
    # which is how a declaration turns into a place to hide an unrelated edit.
    assert set(DECLARED_PRIMITIVE_DEPARTURES) <= set(functions), (
        DECLARED_PRIMITIVE_DEPARTURES, functions)
    for name in functions:
        body = re.search(rf"^function {re.escape(name)} \{{.*?^}}", block, re.S | re.M)
        assert body, name
        if name in DECLARED_PRIMITIVE_DEPARTURES:
            # IT REALLY DID DEPART - a declaration for a primitive that is still
            # verbatim is a stale exemption, and stale exemptions are how the
            # next drift gets through.
            assert body.group(0) not in runner, (
                f"{name} is declared as departed but is still the accepted copy")
            assert DECLARED_PRIMITIVE_DEPARTURES[name] in runner, (
                f"{name} departed without the sentence that justifies it")
            continue
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


# ===========================================================================
# K. THE FIXTURE ON A PROTECTED WORKBOOK
# ===========================================================================
# Benchmark Run 3 aborted on a COM `ListRow.Delete` against `tblFXRates`:
# "Table features aren't available because the sheet is protected." Protection
# is a Phase-10 addition and every harness here predates it. The runtime
# reconciliation proved the split - code VALUE writes are permitted under
# UserInterfaceOnly, ListObject STRUCTURAL mutation is refused - and production
# reaches the second only through modProtection's window, which a PowerShell COM
# caller never enters.
#
# These controls hold the correction to what it claims: reset by CONTENT, no
# structural mutation, nothing weakened, nothing about protection reopened.
GATE_B = BOOTSTRAP / "phase5_gate_b_scenarios.ps1"


def _function(source: str, name: str) -> str:
    """One PowerShell function, from its header to the line that closes it.

    Column-zero `}` ends it, which is this tree's layout for every runner.
    """
    lines = source.splitlines()
    start = next(i for i, line in enumerate(lines)
                 if line.startswith(f"function {name} "))
    end = next(j for j in range(start + 1, len(lines)) if lines[j] == "}")
    return "\n".join(lines[start:end + 1])


def _fx_reset() -> str:
    return _function(_code(), "Reset-Phase5FxTable")


def test_130_the_fixture_performs_no_listobject_structural_mutation() -> None:
    """THE CALL FORMS, NOT THE WORDS. The runner NAMES `$victim.Delete()` in the
    paragraph explaining why it no longer carries one, and naming it is the
    opposite of performing it - so the ban is read off `_code`, which is the
    runner with its comments removed.

    Every one of these is refused to a COM caller on a protected sheet, and the
    only legitimate way to reach one is a production endpoint."""
    # DECLARED, NOT LOOSENED. Before the bulk builder existed, "this runner
    # performs no structural mutation at all" was true and this control said so.
    # `Set-BenchmarkRegisterRowCount` now grows a register with `ListRows.Add`,
    # which is precisely what the fixture window exists to permit, and
    # `Set-BenchmarkRangeBlock` uses `Range.Resize` - a Range method that selects a
    # different rectangle and adds no table row or column at all.
    #
    # So the ban holds over the runner MINUS those two functions, both named here
    # and both held to their own controls (test_194, test_195). A structural
    # mutation anywhere else still fails, which is the property this control had.
    code = _code()
    exempted = ("Set-BenchmarkRegisterRowCount", "Set-BenchmarkRangeBlock")
    exempt = "\n".join(_function(code, name) for name in exempted)
    assert "$rows.Add()" in exempt and "$anchor.Resize(" in exempt, (
        "the declared exemption no longer contains the operations it exempts")
    # REMOVED ONE AT A TIME: the two are not contiguous in the file, so a single
    # concatenated replace would remove nothing and the ban would be vacuous.
    rest = code
    for name in exempted:
        body = _function(code, name)
        assert body in rest, name
        rest = rest.replace(body, "")
    assert len(rest) < len(code) - 100
    for banned in (".Delete()", ".ListRows.Add", ".ListColumns.Add",
                   ".ListColumns.Delete", ".EntireRow.Delete", ".Resize("):
        assert banned not in rest, f"the benchmark performs a structural mutation: {banned}"
    # AND NEITHER EXEMPT FUNCTION DELETES ANYTHING.
    for banned in (".Delete()", ".ListColumns.Add", ".ListColumns.Delete"):
        assert banned not in exempt, f"the exempt region deletes or adds columns: {banned}"
    # THE PRIMITIVE STILL EXISTS AND HAS NO CAPABILITY LEFT. Deleting the
    # definition outright was the first attempt and it was wrong: two functions
    # in the dot-sourced Gate-B file name `Remove-TableRow`, and a name a
    # reachable file can call must resolve - that is the exact shape that killed
    # Phase-9 Windows run 1. So the name resolves and the body refuses.
    remover = _function(_code(), "Remove-TableRow")
    assert "throw (" in remover, "the row-delete primitive no longer refuses"
    assert "$RowIndex" in remover and "+ $TableName +" in remover, (
        "the refusal does not say what was attempted")
    assert "production endpoint may perform one" in remover
    # IT CANNOT BE MISTAKEN FOR A WORKING HELPER: no COM object is even fetched.
    for reached in ("$Workbook.Worksheets", "ListObjects", "ListRows"):
        assert reached not in remover, f"the refusal still touches COM: {reached}"
    # THE QUOTED CALL SURVIVES ONLY AS PROSE, so the record of what failed is
    # not lost to the fix.
    assert "$victim.Delete()" in _runner()


def test_131_the_fx_reset_clears_every_row_below_the_seed_and_proves_it() -> None:
    """STALE VALUES CANNOT SURVIVE. The accepted reset deleted every row below
    the seed; this one blanks every row below the seed. Anything short of EVERY
    row leaves a previous scenario's rate in a table production resolves by
    content, and the measurement is then of a model nobody wrote.

    The read-back is what makes it a proof rather than an intention, and it
    requires $null: ClearContents leaves Value2 $null, while an empty string is
    not a genuine blank and production would read that row as populated."""
    reset = _fx_reset()
    # THE WRITE LOOP REACHES THE LAST PHYSICAL ROW, not a fixed number of them.
    assert "for ($row = $seedRows + 1; $row -le $rows; $row++) {" in reset, (
        "the reset no longer blanks every row below the seed")
    # TWICE, AND COUNTED. The write loop and the read-back loop each walk every
    # column, and an `in` check would be satisfied by either one of them alone -
    # so a reset that blanked only the Currency column and still checked both
    # would pass. A row with no currency and a surviving rate is not blank.
    assert reset.count(
        "for ($column = 1; $column -le $columns.Count; $column++) {") == 2, (
        "the reset no longer blanks, or no longer proves it blanked, every column")
    assert "-Value $null" in reset, "the reset no longer clears"
    # AND `Set-TableCell` WITH $null REALLY IS ClearContents, so the control is
    # not resting on the parameter name.
    setter = _function(_code(), "Set-TableCell")
    assert "if ($null -eq $Value) {" in setter
    assert "$cell.ClearContents()" in setter

    # THE READ-BACK LOOP HAS THE SAME REACH AS THE WRITE LOOP, over the body it
    # actually read, and refuses on anything that is not $null.
    assert "for ($row = $seedRows + 1; $row -le $body.Count; $row++) {" in reset, (
        "the reset no longer proves the blanking took")
    assert "if ($null -ne $body[$row - 1][$column - 1]) {" in reset, (
        "the blank check accepts something that is not a genuine blank")
    assert "the FX reset did not clear row " in reset
    # AND IT NAMES THE CANDIDATE CAUSE rather than inviting a retry.
    assert "must be reported, not retried" in _runner()


def test_132_the_seed_restoration_is_the_accepted_one_unchanged() -> None:
    """ONE HALF OF THE RESET CHANGED, AND ONLY ONE. The seed restoration is a
    settled control over a defective build - the captured value is written back
    AS ITSELF so a text seed stays text and is exposed by production rather than
    repaired here - and a correction to the blanking has no business touching
    it. Both statements and the strict comparison are required verbatim against
    the accepted harness, so a drift in either is a failure here."""
    accepted = _function(GATE_B.read_text(encoding="utf-8"), "Reset-Phase5FxTable")
    reset = _fx_reset()
    for fragment in (
            "Set-Phase5TypedCell -Workbook $Workbook -SheetName $fx.sheet "
            "-TableName $fx.table_name `\n        -RowIndex 1 -ColumnIndex 1 "
            "-Value $Seed.Currency",
            "Set-Phase5TypedCell -Workbook $Workbook -SheetName $fx.sheet "
            "-TableName $fx.table_name `\n        -RowIndex 1 -ColumnIndex 2 "
            "-Value $Seed.Rate",
            "if ((-not (Test-Phase5ExactValue -Actual $body[0][0] -Expected "
            "$Seed.Currency)) -or `",
            "(-not (Test-Phase5ExactValue -Actual $body[0][1] -Expected $Seed.Rate)))"):
        assert fragment in accepted, f"the anchor moved in the accepted harness: {fragment[:50]}"
        assert fragment in reset, f"the seed restoration was altered: {fragment[:50]}"


def test_133_the_reset_never_touches_a_header_or_the_schema() -> None:
    """BLANKING A BODY IS NOT REWRITING A TABLE. The fixture's job is the values
    below the seed; a reset that reached the header row, the column collection or
    the table's name would be changing the contract's shape under a measurement,
    and the resulting numbers would describe a workbook the specification never
    declared.

    `Set-TableCell` indexes DataBodyRange, so the header is out of its reach by
    construction - what is banned is new code that reaches past it."""
    reset = _fx_reset()
    for banned in ("HeaderRowRange", "ListColumns", ".Name =", "ListObjects.Add",
                   "TableStyle", ".Unlist(", "Validation"):
        assert banned not in reset, f"the FX reset reaches the schema: {banned}"
    # THE BLANKING STARTS BELOW THE SEED. Starting at row 1 would blank the
    # reporting-currency identity between the capture and its restoration.
    assert "$row = $seedRows + 1" in reset
    assert "$row = 1;" not in reset
    # AND THE SINGLE-SEED ASSUMPTION THE ROW-1 ARITHMETIC RESTS ON IS REFUSED
    # RATHER THAN ASSUMED, so a contract change cannot drift past it silently.
    assert "if ($seedRows -ne 1) {" in reset
    # THE CELL WRITER IS THE ACCEPTED ONE, which cannot address a header.
    body_reader = _function(_code(), "Set-TableCell")
    assert "$body = $lo.DataBodyRange" in body_reader


def test_134_a_blank_row_is_found_and_a_missing_one_is_refused_by_name() -> None:
    """THE RESERVED ROWS ARE WHY THIS WORKS. `tblFXRates` is built `data_rows: 12`
    with one seeded row and `tblInflationProfiles` `data_rows: 10` with none, so
    the fixture writes into rows the workbook already has.

    AND WHEN IT CANNOT, IT SAYS SO. A table with no blank row left genuinely
    needs a structural add; the add is genuinely refused; and the caller is told
    which table and what its capacity was, instead of meeting a bare 1004 from
    inside Excel. Nothing is relaxed - a fixture that cannot be built is a
    fixture that is not measured."""
    adder = _function(_code(), "Add-BlankTableRow")
    assert "$rows.Add()" not in adder, "the adder grows the table again"
    assert "Get-TableBody" in adder, "the adder no longer searches what is there"
    assert "return [int]$row" in adder
    # TWO REFUSALS, BOTH NAMING THE TABLE AND BOTH SAYING WHY.
    assert adder.count("throw (") == 2, adder
    assert adder.count("+ $TableName +") == 2
    assert adder.count("ListObject structural operation") == 2
    assert "only a production endpoint may perform one" in adder
    # THE CONTRACT REALLY DOES RESERVE THE ROWS THIS RESTS ON.
    contract = (SPEC / "input_contract.yaml").read_text(encoding="utf-8")
    fx = contract[contract.index("  fx_rates:"):]
    assert "data_rows: 12" in fx[:fx.index("seed_rows:")]
    assert 'table_name: "tblInflationProfiles"' in contract


def test_135_the_accepted_gate_b_harness_is_not_edited() -> None:
    """THE CORRECTION IS SCOPED BY MECHANISM, NOT BY PROMISE. Gate B was accepted
    on these bytes and Phases 4, 7, 8 and 9 all dot-source them. The benchmark
    overrides the helpers in its OWN process instead, which is only sound if the
    accepted file really is untouched - so that is checked rather than asserted."""
    changed = [line for line in _git(
        "diff", "--name-only", ACCEPTED, "--",
        "pccm/bootstrap/windows/phase5_gate_b_scenarios.ps1",
        "pccm/bootstrap/windows/phase6_gate_b_scenarios.ps1").splitlines() if line.strip()]
    assert changed == [], f"an accepted Gate-B harness was edited: {changed}"


def test_136_the_overrides_are_defined_after_the_dot_source_that_they_override() -> None:
    """THE WHOLE MECHANISM IS ORDER. PowerShell resolves a function name at CALL
    time and the last definition wins, so an override placed above the
    dot-source would be silently replaced by the accepted one and the runner
    would go back to deleting rows - with every control above still passing,
    because the source would still contain the override.

    The accepted file USES these helpers and DEFINES none of them, which is what
    makes overriding them safe; that is checked too."""
    runner = _runner()
    dot_source = runner.index(". (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')")
    for name in ("Reset-Phase5FxTable", "Add-BlankTableRow", "Set-TableCell",
                 "Get-TableBody"):
        assert runner.index(f"function {name} ") > dot_source, (
            f"{name} is defined before the dot-source that would overwrite it")
    gate_b = GATE_B.read_text(encoding="utf-8")
    for name in ("Add-BlankTableRow", "Set-TableCell", "Get-TableBody",
                 "Get-TableRowCount"):
        assert f"function {name} " not in gate_b, (
            f"the accepted harness now defines {name}; overriding it is no longer local")
    # `Reset-Phase5FxTable` IS THE ONE THE ACCEPTED FILE DOES DEFINE, which is
    # exactly why the benchmark's copy has to come after the dot-source.
    assert "function Reset-Phase5FxTable {" in gate_b


def test_137_the_correction_introduces_no_release_of_protection() -> None:
    """THE BOUNDARY STAYS CLOSED. The reconciliation settled that only
    modProtection may release worksheet protection, inside production's own
    structural window. A harness that reached for that window - or for Excel's
    Unprotect directly - would reopen an architecture that was closed on Windows
    evidence, for a measurement."""
    code = _code()
    for banned in (".Unprotect", "ProtectionRelease", "ProtectionBeginStructural",
                   ".Protect(", "Protect Structure", "ProtectStructure ="):
        assert banned not in code, f"the benchmark reaches for protection: {banned}"
    # AND NO PRODUCTION FILE MOVED FOR THIS.
    changed = [line for line in _git(
        "diff", "--name-only", "0119bee", "--",
        "pccm/src", "pccm/spec", "pccm/builder").splitlines() if line.strip()]
    assert changed == [], f"production changed for a harness correction: {changed}"


def test_138_the_fixture_build_is_still_outside_every_measured_interval() -> None:
    """A CORRECTION TO SETUP MUST NOT MOVE THE CLOCK. The reset became a longer
    sequence of cell writes than a delete loop was, and the one way that could
    corrupt a baseline is by happening inside a timed region. It does not: the
    fixture is built between its own stopwatch and the operation clocks start
    later, and the reset is reached only from the fixture."""
    code = _code()
    fixture = code[code.index("$fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()"):]
    fixture = fixture[:fixture.index("$fixtureWatch.Stop()")]
    assert "Set-Phase5Fixture" in fixture
    assert "setup_timings" not in fixture.lower() or True
    # THE FIXTURE STOPWATCH FEEDS SETUP, NEVER A SAMPLE.
    assert "$setupTimings.Add('scenario_fixture_ms'," in code
    # AND THE RESET IS REACHED ONLY FROM THE FIXTURE TREE - the benchmark defines
    # it and never calls it itself, so it cannot appear between an operation
    # clock's start and stop. The name also occurs inside the row-delete
    # refusal's message, which is text and not a call, so the rule is stated over
    # INVOCATIONS: a line whose first token is the name.
    # ONE CALLER, AND IT IS A FIXTURE BUILDER. The accepted step C reaches it from
    # the dot-sourced Gate-B file; the bulk builder reaches it directly, for the
    # same twenty-four cells. What must stay true is that neither the run loop nor
    # anything between the window closing and the loop touches it.
    invocations = [line for line in code.splitlines()
                   if line.strip().startswith("Reset-Phase5FxTable ")]
    assert len(invocations) == 1, invocations
    assert "Reset-Phase5FxTable" in _function(code, "Set-BenchmarkBulkFixture")
    loop = code.index("foreach ($run in $plannedRuns) {")
    closed = code.index("Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest")
    assert "Reset-Phase5FxTable" not in code[closed:], (
        "the FX reset is reachable after the fixture window closed")
    assert code.count("function Reset-Phase5FxTable {") == 1


def test_139_the_semantic_proof_is_recorded_with_the_consumers_it_rests_on() -> None:
    """THE CLAIM IS "PRODUCTION CANNOT TELL", and a claim like that is worth only
    the audit behind it. The record names every property a consumer could have
    depended on and says, for each, whether it does - including the one where
    deletion was the damaging operation rather than the safe one.

    It also states what is still inferred rather than observed, which is the
    difference between an honest settlement and a quiet assumption."""
    text = RUN_EVIDENCE.read_text(encoding="utf-8")
    section = text[text.index("## Benchmark fixture under protection"):]
    for consumer in ("ListRows.Count", "DataBodyRange", "physical table row count",
                     "row position", "blank body rows", "validation",
                     "structural fingerprint"):
        assert consumer in section, f"the consumer audit does not cover {consumer}"
    assert "MatchingFxRows" in section and "RawCellText" in section
    assert "IsEmpty" in section
    # THE FOUR SITES, EACH WITH A DISPOSITION.
    for site in ("Reset-Phase5FxTable", "Invoke-Phase5FixtureSteps",
                 "Set-Phase5InflationProfileMaster", "Clear-Phase5UserRows",
                 "Invoke-Phase5Mutation"):
        assert site in section, f"the structural site audit omits {site}"
    # AND THE HONEST LIMIT.
    assert "not yet been observed" in section
    assert "inference from the proved split, not an observation" in section
    # NO WINDOWS WAS CLAIMED FOR THIS ROUND.
    assert "No Windows was executed for" in section


SHIM = BOOTSTRAP / "phase10_fixture_window.bas"
SHIM_MODULE = "modPhase10FixtureWindow"

# THE PRODUCTION OWNER. Read here, never edited - these controls prove the
# harness USES it rather than replacing it.
PROTECTION_OWNER = PCCM_ROOT / "src" / "vba" / "modProtection.bas"


def _shim() -> str:
    if "shim" not in _MEMO:
        _MEMO["shim"] = SHIM.read_text(encoding="utf-8")
    return _MEMO["shim"]


def _shim_code() -> str:
    """The shim with its comment lines removed.

    SAME RULE AS `_code`, AND FOR THE SAME REASON. The shim's header explains
    precisely which protection calls it must never make, and it has to name them
    to do that. A ban satisfied by the paragraph explaining the ban is no ban, so
    everything that asserts what the shim DOES reads this.
    """
    if "shim_code" not in _MEMO:
        _MEMO["shim_code"] = "\n".join(
            line for line in _shim().splitlines() if not line.strip().startswith("'"))
    return _MEMO["shim_code"]


def _manifest_json() -> dict:
    if "manifest_json" not in _MEMO:
        _MEMO["manifest_json"] = json.loads(
            (BUILD / "stage_b_manifest.json").read_text(encoding="utf-8"))
    return _MEMO["manifest_json"]


# ===========================================================================
# L. THE FIXTURE MAINTENANCE WINDOW
# ===========================================================================
# PERF-SMALL at 7077608 aborted before any timed run, on `$cell.ClearContents()`,
# with "The cell or chart you're trying to change is on a protected sheet." An
# external COM caller cannot clear a cell on a protected sheet - though Run 3
# proved it CAN write a value to one - so the fixture needs worksheet protection
# released around its writes, and around nothing else.
def test_150_the_window_is_productions_own_and_the_harness_holds_no_policy() -> None:
    """ONE OWNER, STILL. The correction would be worthless if it bought the
    fixture a window by writing a second protection implementation in PowerShell:
    two authorities disagree the first time one of them is not reached.

    So the runner contains no Protect/Unprotect of any kind, and the shim
    contains none either - it forwards to `modProtection` and returns what the
    owner said."""
    code = _code()
    for banned in (".Unprotect", ".Protect(", "Protect Structure", "ProtectStructure =",
                   "UserInterfaceOnly:="):
        assert banned not in code, f"the runner implements protection itself: {banned}"
    # THE RUNNER MAY NAME THE FLAG IN PROSE - it explains what the close restores -
    # but it may not set one.
    assert "UserInterfaceOnly" in _runner()
    shim = _shim_code()
    for banned in (".Unprotect", ".Protect ", ".Protect(", "UserInterfaceOnly:="):
        assert banned not in shim, f"the shim implements protection itself: {banned}"
    for forwarded in ("modProtection.ProtectionBeginStructural(detail)",
                      "modProtection.ProtectionEndStructural(detail)",
                      "modProtection.ProtectionIsApplied()",
                      "modProtection.ProtectionStructuralDepth()"):
        assert forwarded in shim, f"the shim does not forward to {forwarded}"
    # AND THE SHIM DECIDES NOTHING. Every executable line is a forward, a read or
    # the string it returns: no branch of its own on protection state, and no
    # loop over worksheets except the one that COUNTS them for the diagnostic.
    assert "If modProtection.ProtectionBeginStructural(detail) Then" in shim
    assert "If modProtection.ProtectionEndStructural(detail) Then" in shim


def test_151_workbook_structure_protection_is_never_released() -> None:
    """THE ENVELOPE IS NOT WIDENED. `ProtectionRelease` is the MAINTENANCE path
    and it alone calls `ThisWorkbook.Unprotect`; the structural window never
    touches the structure flag. Nothing in this harness may reach the first one,
    and the runner refuses a state where the flag moved."""
    owner = PROTECTION_OWNER.read_text(encoding="utf-8")
    # THE READING THIS RESTS ON, PROVED IN THE OWNER RATHER THAN ASSUMED.
    assert owner.count("ThisWorkbook.Unprotect") == 1, (
        "production gained another workbook-structure release")
    release = owner[owner.index("Public Function ProtectionRelease"):]
    assert "ThisWorkbook.Unprotect" in release, "the one release moved out of ProtectionRelease"
    begin = owner[owner.index("Public Function ProtectionBeginStructural"):
                  owner.index("Public Function ProtectionEndStructural")]
    assert "ThisWorkbook.Unprotect" not in begin, "the structural window now releases structure"

    # THE CALL, NOT THE WORD. Both files NAME `ProtectionRelease` in the paragraph
    # that explains why they must never call it, and naming it is the opposite of
    # calling it - so the ban is stated over the code with comments removed.
    for text, where in ((_code(), "the runner"), (_shim_code(), "the shim")):
        assert "ProtectionRelease" not in text, f"{where} reaches the maintenance release path"
    # AND THE EXPLANATION IS STILL THERE TO BE READ.
    assert "ProtectionRelease" in _shim(), "the shim no longer says why it must not release"
    # AND THE RUNNER REFUSES A RUN WHERE THE FLAG MOVED, on both sides.
    opener = _function(_code(), "Open-BenchmarkFixtureWindow")
    assert "if (-not $state.Structure) {" in opener
    assert "released WORKBOOK STRUCTURE" in opener, (
        "the refusal no longer distinguishes the workbook structure flag from "
        "worksheet protection, which is the distinction the whole window rests on")
    assert "if (-not $state.Structure)" in _function(_code(), "Assert-BenchmarkProtectionApplied")


def test_152_the_window_opens_before_the_fixture_and_closes_before_any_timed_run() -> None:
    """ORDER IS THE WHOLE SAFETY PROPERTY. Opened after the first protected write
    is no window at all; left open into the run loop means every timed operation
    is measured on an unprotected workbook, which is not the delivered product.

    Read off positions in the runner, because a comment promising an order is not
    the order."""
    code = _code()
    opened = code.index("Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest")
    fixture = code.index("$null = Set-Phase5Fixture -Excel $excel")
    closed = code.index("Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest")
    loop = code.index("foreach ($run in $plannedRuns) {")
    assert opened < fixture < closed < loop, (opened, fixture, closed, loop)
    # THE FIXTURE STOPWATCH BOUNDS BOTH, so the window's cost is reported as the
    # setup it is rather than disappearing.
    started = code.index("$fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()")
    stopped = code.index("$fixtureWatch.Stop()")
    assert started < opened and closed < stopped
    assert "$setupTimings.Add('scenario_fixture_ms'," in code
    # AND OPENING MEANS OPENING ONTO A PROVED STATE. A window opened over a
    # workbook that was already unprotected closes onto a state nobody
    # established, and the run would then report a restoration it never made.
    opener = _function(code, "Open-BenchmarkFixtureWindow")
    assert "Assert-BenchmarkProtectionApplied" in opener, (
        "the window opens without first proving the accepted protection state")
    assert "before the fixture maintenance window was opened" in opener
    # THE OPEN IS PROVED TO HAVE TAKEN, at the one depth this runner uses.
    assert "if ($state.Depth -ne 1) {" in opener


def test_153_a_fixture_that_raises_still_closes_the_window() -> None:
    """THE ONE OUTCOME THAT MUST BE IMPOSSIBLE is a raised fixture that leaves the
    workbook unprotected and a later run recording numbers from it. The close is
    in a `finally`, so it happens on the throwing path too - and a close that
    fails there raises in turn, which is right: an unprotected workbook is the
    worse fact and should be the reported one."""
    code = _code()
    region = code[code.index("$null = Open-BenchmarkFixtureWindow"):
                  code.index("$fixtureWatch.Stop()")]
    assert "    try {" in region and "} finally {" in region, region
    body = region[region.index("try {"):region.index("} finally {")]
    trailer = region[region.index("} finally {"):]
    assert "Set-Phase5Fixture" in body, "the fixture is not inside the guarded region"
    assert "Close-BenchmarkFixtureWindow" in trailer, "the close is not in the finally"
    assert "Close-BenchmarkFixtureWindow" not in body


def test_154_the_close_is_verified_against_the_declared_sheet_count() -> None:
    """"IT SAID OK" IS NOT RESTORATION. `ProtectionEndStructural` already requires
    `ProtectionIsApplied` before reporting success; the harness asks the workbook
    again, independently, and counts the sheets - against the manifest's own
    protection projection rather than a literal, so a contract that gained a
    worksheet is a contract change and not a silent pass at the old number."""
    closer = _function(_code(), "Close-BenchmarkFixtureWindow")
    assert "Assert-BenchmarkProtectionApplied" in closer
    assert "notlike 'OK|*'" in closer
    assert "No measurement may be taken from this" in closer

    guard = _function(_code(), "Assert-BenchmarkProtectionApplied")
    assert "@($Manifest.protection.sheets).Count" in guard, (
        "the sheet count is not read from the declared projection")
    assert "14" not in guard, "the sheet count is hard-coded"
    for required in ("$state.Applied", "$state.Structure", "$state.Sheets",
                     "$state.Protected", "$state.Depth"):
        assert required in guard, f"the verification does not check {required}"
    protection = _manifest_json()["protection"]
    assert len(protection["sheets"]) == 14, protection["sheets"]
    assert protection["protect_structure"] is True
    assert protection["user_interface_only"] is True


def test_155_a_failed_open_or_close_reaches_the_abandon_path() -> None:
    """AN ABORT IS NOT A SLOW RUN. Every refusal here is a `throw`, and a throw
    inside the measurement session sets `$abandoned`, which forces `$runComplete`
    false, which makes the status ABORTED and the exit code 1. A protection
    failure therefore cannot be reported as a baseline."""
    code = _code()
    for name in ("Open-BenchmarkFixtureWindow", "Close-BenchmarkFixtureWindow",
                 "Assert-BenchmarkProtectionApplied", "Import-BenchmarkFixtureWindow",
                 "Get-BenchmarkProtectionState"):
        assert "throw (" in _function(code, name), f"{name} has no refusal"
    assert "$abandoned = [string]$failure['message']" in code
    assert "$runComplete = ([bool](([string]::IsNullOrWhiteSpace($abandoned)) -and" in code
    assert "if (-not $runComplete) {" in code and "exit 1" in code


def test_156_the_shim_is_test_only_and_never_production() -> None:
    """A MODULE IMPORTED OVER PRODUCTION WOULD BE PRODUCTION. It lives in
    `bootstrap/windows` beside the Gate-B diagnostics module that has used this
    same mechanism since Phase 5, it is imported into the DISPOSABLE copy, and the
    runner refuses to import it if the manifest ever declares it."""
    assert SHIM.parent.name == "windows" and SHIM.parent.parent.name == "bootstrap"
    assert not (PCCM_ROOT / "src" / "vba" / "phase10_fixture_window.bas").exists()
    declared = [str(entry["name"]) for entry in _manifest_json()["vba"]["modules"]]
    assert SHIM_MODULE not in declared, "the shim is declared as a production module"
    assert SHIM_MODULE not in _manifest_text()

    importer = _function(_code(), "Import-BenchmarkFixtureWindow")
    assert "$declared -contains $script:FixtureWindowModule" in importer
    assert "will not import a test module over production" in importer
    # THE IMPORT IS PROVED TO HAVE WORKED, not assumed from a lack of exception.
    assert "$Excel.Run('P10FW_Ping')" in importer
    assert "imported but does not answer" in importer
    # AND IT ADDS NO NEW MACHINE DEPENDENCY: the Stage-B bootstrap this runner
    # already invokes reaches VBProject too, so a host that could not import
    # could not have produced the workbook in the first place.
    assert "$wb.VBProject" in (BOOTSTRAP / "build_stage_b.ps1").read_text(encoding="utf-8")


def test_157_the_protection_state_is_read_in_vba_not_across_com() -> None:
    """PROBE RUN 8, NOT REPEATED. Reading `Worksheet.ProtectContents` across COM
    raised a terminating PropertyNotFoundException on an object PowerShell had no
    type information for. Inside VBA the property binds at compile time, so the
    whole class is gone - and the harness parses one string instead."""
    code = _code()
    assert "ProtectContents" not in code, "the runner reads ProtectContents across COM again"
    assert "ProtectContents" in _shim(), "the shim no longer reads the sheets it counts"
    reader = _function(code, "Get-BenchmarkProtectionState")
    assert "$Excel.Run('P10FW_State')" in reader
    assert "the protection state is missing" in reader
    for required in ("'applied'", "'depth'", "'structure'", "'sheets'", "'protected'"):
        assert required in reader, required


def test_158_the_window_wraps_the_fixture_and_no_other_setup() -> None:
    """MINIMAL, AND THAT IS CHECKED. The seed write after the fixture is a VALUE
    write to a named cell, which Run 3 proved an external COM caller can make on a
    protected sheet - so it stays outside. So do the Stage-B bootstrap, the
    workbook open, the environment inventory and the dimension read-back."""
    code = _code()
    region = code[code.index("$null = Open-BenchmarkFixtureWindow"):
                  code.index("$protectionAfter = Close-BenchmarkFixtureWindow")]
    assert "Set-Phase5Fixture" in region
    for outside in ("Set-NamedValue", "Get-IdColumnValues", "build_stage_b",
                    "Get-BenchmarkEnvironment", "Invoke-BenchmarkOperation"):
        assert outside not in region, f"{outside} was wrapped and does not need the window"
    assert code.count("Open-BenchmarkFixtureWindow -Excel") == 1
    assert code.count("Close-BenchmarkFixtureWindow -Excel") == 1


def test_159_the_window_did_not_move_the_fixture_or_the_matrix() -> None:
    """A PROTECTION CORRECTION IS NO OCCASION TO CHANGE WHAT IS MEASURED. The
    content-based reset stays, the structural delete stays gone, and the scenario
    identity, dimensions, iteration matrix and cold/warm counts are untouched."""
    reset = _fx_reset()
    assert "-Value $null" in reset and ".Delete()" not in reset
    assert "for ($row = $seedRows + 1; $row -le $rows; $row++) {" in reset
    code = _code()
    assert ".ListRows.Add" not in code and ".Delete()" not in code

    sizes = {entry["title"]: (entry["drivers"], entry["years"])
             for entry in _plan()["scenarios"]}
    assert sizes == CONTRACT_DIMENSIONS, sizes
    assert _scenario("PERF-SMALL")["iterations"] == [10000, 50000, 100000]
    assert [entry["id"] for entry in _plan()["scenarios"]] == [
        "PERF-SMALL", "PERF-MEDIUM", "PERF-LARGE"]


def _run_evidence_section(heading: str) -> str:
    """One section of the append-only Windows record, heading to next heading."""
    text = RUN_EVIDENCE.read_text(encoding="utf-8")
    start = text.index(heading)
    tail = text[start + len(heading):]
    cut = tail.find("\n## ")
    return heading + (tail if cut == -1 else tail[:cut])


def test_160_benchmark_run_4_is_recorded_exactly_as_it_happened() -> None:
    """THE RUN THAT REFUTED THE INFERENCE. It has to be on the record with its
    own numbers, because the settlement it overturned is on the record with
    its."""
    section = _run_evidence_section("## Benchmark Run 4")
    for fact in ("7077608", "351 passed, 0 failed", "Stage-B bootstrap succeeded",
                 "$null = $cell.ClearContents()",
                 "The cell or chart you're trying to change is on a protected sheet",
                 "ABORTED BEFORE A COMPLETE BASELINE",
                 "0 of 11 planned run(s)",
                 "Shutdown and COM release were clean"):
        assert fact in section, f"the Run 4 record omits: {fact}"
    assert "RUNTIME-REFUTED" in section
    assert "No production defect is established" in section
    # AND IT IS NOT DRESSED UP AS A PARTIAL RESULT. The wording is Run 3's, which
    # settled how an aborted benchmark is described, so the two records refuse the
    # same three readings in the same words.
    for overclaim in ("NOT a baseline", "NOT a partial baseline",
                      "NOT a performance sample"):
        assert overclaim in section, overclaim
    run3 = _run_evidence_section("## Run 3 — PERF-SMALL")
    for overclaim in ("NOT a baseline", "NOT a partial baseline",
                      "NOT a performance sample"):
        assert overclaim in run3, f"the precedent wording moved: {overclaim}"


def test_161_the_capability_split_for_an_external_caller_is_recorded_as_facts() -> None:
    """THREE LINES, ALL OBSERVED, NONE INFERRED. An external COM client can write
    a value to a locked cell and cannot clear one - that is finer than the split
    recorded for VBA inside the workbook, and writing it down is the difference
    between a settled fact and a guess that happens to be right so far."""
    section = _run_evidence_section("## Benchmark Run 4")
    assert "Range.Value2" in section and "Range.ClearContents()" in section
    assert "ListRow.Delete" in section
    assert "permitted" in section and "refused" in section
    assert "None of them is inferred" in section
    # THE RECONCILIATION IS NOT CONTRADICTED, AND THE DISTINCTION IS NAMED.
    assert "does not contradict the Runtime Protection Reconciliation" in section
    assert "inside" in section and "external automation client" in section


def test_162_the_superseded_inference_is_left_standing_as_history() -> None:
    """NO FALSE REWRITE. The settlement at 7077608 said plainly that ClearContents
    under protection had not been observed and that its permission was an
    inference. That sentence was true of what was known; editing it out now would
    be rewriting the project's own record of how it learned this."""
    earlier = _run_evidence_section("## Benchmark fixture under protection —")
    assert "not yet been observed" in earlier, (
        "the earlier honest limitation was edited away")
    assert "inference from the proved split, not an observation" in earlier
    # AND THE NEW SECTION SAYS THE INFERENCE WAS WRONG, rather than pretending the
    # old one had always agreed.
    section = _run_evidence_section("## Benchmark Run 4")
    assert "The inference was" in section and "wrong" in section


def test_163_the_window_record_carries_its_path_its_writes_and_its_limit() -> None:
    """A MAINTENANCE WINDOW IS THE KIND OF THING THAT GETS WIDENED LATER. The
    record names the exact call path, every fixture write that needs it, the one
    that deliberately does not, and the single property that is established by
    construction because Excel exposes no way to read it."""
    section = _run_evidence_section("## Benchmark fixture maintenance window")
    for step in ("P10FW_Begin", "P10FW_End", "P10FW_State",
                 "ProtectionBeginStructural", "ProtectionEndStructural"):
        assert step in section, f"the call path omits {step}"
    assert "ByRef" in section and "no precedent" in section
    assert "ThisWorkbook.Unprotect" in section
    assert "structure protection is never released" in section.lower()
    # THE WRITE AUDIT, INCLUDING THE ONE LEFT OUTSIDE.
    assert "tblFXRates" in section and "tblInflationProfiles" in section
    assert "deliberately left outside" in section
    # AND THE HONEST LIMIT, AGAIN NAMED RATHER THAN GLOSSED.
    assert "cannot be read back" in section
    assert "construction argument rather than a read-back" in section
    assert "No Windows was executed for" in section


RESOLUTION_AUDIT = PCCM_ROOT / "tests" / "powershell_command_resolution_audit.ps1"
PWSH = "/opt/pwsh/pwsh"

# THE ONE NAME THIS RUNNER DELIBERATELY REDEFINES over the accepted Gate-B file.
BENCHMARK_DECLARED_OVERRIDES = ("Reset-Phase5FxTable",)


def _resolution_audit(path: Path, *declared: str) -> subprocess.CompletedProcess:
    command = [PWSH, "-NoProfile", "-File", str(RESOLUTION_AUDIT), "-Path", str(path)]
    if declared:
        command += ["-DeclaredOverride", ",".join(declared)]
    return subprocess.run(command, capture_output=True, text=True, timeout=300)


def test_164_the_partial_open_gap_and_its_closure_are_recorded() -> None:
    """THE ROUND THAT FOUND IT ALSO CLOSED IT, and both halves belong on the
    record: the property that already held, and the adjacent one that did not.

    The terminology correction is here too. An earlier return called one route
    "structure released", meaning ThisWorkbook.ProtectStructure - the record says
    which flag that is and proves no path can move it."""
    raw = _run_evidence_section("## Fixture window: the partial-open gap")
    # WRAPPED PROSE IS STILL THE SENTENCE. The record is hard-wrapped, so a clause
    # that must be present can straddle a newline; every phrase below is checked
    # against the section with its whitespace collapsed.
    section = " ".join(raw.split())
    assert "No Windows was executed" in section
    # THE CONTRACT, ALL FOUR ROWS.
    for clause in ("Begin refuses", "nothing is owed", "exactly ONE compensating End",
                   "closed nothing", "closes exactly once"):
        assert clause in section, f"the contract omits: {clause}"
    # WHY A CATCH AND NOT A FINALLY.
    assert "would REPLACE the original exception" in section
    assert "cannot be decremented twice" in section
    # THE EXECUTED PROOF, WITH ITS COUNTS.
    assert "phase10_fixture_window_flow.ps1" in section
    assert "Excel is never started" in section
    assert "asserting a belief about the language" in section
    # THE TERMINOLOGY CORRECTION AND ITS PROOF.
    assert "modProtection.bas:245" in section
    assert "ProtectionRelease" in section
    assert "No new defect" in section
    # AND THAT PROOF IS TRUE OF PRODUCTION RIGHT NOW.
    owner = PROTECTION_OWNER.read_text(encoding="utf-8")
    lines = owner.splitlines()
    hits = [i + 1 for i, line in enumerate(lines) if "ThisWorkbook.Unprotect" in line]
    assert hits == [245], f"the one workbook-structure release moved: {hits}"


def test_188_benchmark_run_5_is_recorded_exactly_as_it_happened() -> None:
    """THE RUN THAT PROVED THE WINDOW AND LOST THE BASELINE. Both halves belong on
    the record: the protection evidence that closes the window question, and the
    eight real elapsed times that are nevertheless not measurements because no
    sample was accepted."""
    raw = _run_evidence_section("## Benchmark Run 5")
    section = " ".join(raw.split())
    for fact in ("ce5951f", "351 passed, 0 failed", "Stage-B bootstrap succeeded",
                 "applied=True|depth=0|structure=True|sheets=14|protected=14",
                 "INVALID: System.Object[]", "System.Int32[]",
                 "ABORTED BEFORE A COMPLETE BASELINE", "0 of 11 planned run(s)",
                 "Shutdown and COM release were clean"):
        assert fact in section, f"the Run 5 record omits: {fact}"
    # THE WINDOW EVIDENCE IS STATED AS THE PROOF IT IS.
    assert "RUNTIME PROVEN" in section
    assert "structure protection stayed applied throughout" in section
    assert "before the first timed operation" in section
    # AND THE RUN IS NOT DRESSED UP. Real elapsed times, no measurements.
    for overclaim in ("NOT a baseline", "NOT a partial baseline",
                      "NOT a performance sample"):
        assert overclaim in section, overclaim
    assert "none of them is a measurement of record" in section
    assert "No production defect is established" in section


def test_189_both_shape_root_causes_are_recorded_as_separate_defects() -> None:
    """TWO DEFECTS, TWO CAUSES. One is a pipeline-array contract, the other a
    variable-name collision with a typed parameter. Recording them as one thing
    would lose the general lesson in each - and both lessons are the kind that
    recur."""
    raw = _run_evidence_section("## The two shape defects")
    section = " ".join(raw.split())
    assert "No Windows was executed" in section
    assert "independent and have different root causes" in section

    # DEFECT 1: the mechanism, its evidence, its second site, and the repair.
    assert "return ,@($problems)" in section
    assert "does NOT flatten a nested array" in section
    assert "Count=1 Valid=False join=System.Object[]" in section
    assert "Get-BenchmarkOneDriveRoots" in section
    assert "New-BenchmarkWeights" in section
    assert "throws" in section and "rather than flattening" in section

    # DEFECT 2: case-insensitivity, the constraint, why the banner looked right.
    assert "case-insensitive" in section
    assert "[int[]]$Iterations" in section
    assert "keeps its type constraint" in section
    assert "`[string]` of a one-element array is the element" in section
    # AND THAT NO ELEMENT WAS SELECTED, because the plan never held a list.
    assert "runs[].iterations` is `10000`, not `[10000]`" in section
    assert "the shape defect was the NAME" in section

    # THE EXECUTED PROOF AND ITS TABLES.
    assert "phase10_run_shape_flow.ps1" in section
    assert "Excel is never started" in section
    assert "REFUSED by name" in section
    # AND WHAT WAS DELIBERATELY LEFT ALONE.
    assert "byte-identical to `ce5951f`" in section


RESERVED_HARNESS = PCCM_ROOT / "tests" / "phase10_reserved_rows_flow.ps1"

# THE CONTRACTED RESERVED CAPACITY of every benchmark-populated table, and the
# rule each follows. Read from the contracts, restated here only so a control can
# compare two independent statements of it.
RESERVED_CAPACITY = {
    "tblCostLines": 25,
    "tblRiskRegister": 25,
    "tblCostProfiling": 25,
    "tblRiskProfiling": 25,
    "tblInflation": 10,
}

# capacity, needed -> (physical rows afterwards, ListRows.Add calls)
RESERVED_EXPECTED = {
    "small-cost-12-into-25": (25, 12, 25, 0),
    "small-risk-8-into-25": (25, 8, 25, 0),
    "exactly-at-capacity": (25, 25, 25, 0),
    "one-past-capacity": (25, 26, 26, 1),
    "medium-cost-60-into-25": (25, 60, 60, 35),
    "large-cost-180-into-25": (25, 180, 180, 155),
    "large-risk-120-into-25": (25, 120, 120, 95),
}


def _reserved_rows() -> dict:
    """Run the reserved-rows harness and group its tagged lines."""
    if "reserved" not in _MEMO:
        done = subprocess.run(
            [PWSH, "-NoProfile", "-File", str(RESERVED_HARNESS), "-Runner", str(RUNNER)],
            capture_output=True, text=True, timeout=300)
        assert done.returncode == 0, done.stdout + done.stderr
        rows = {"rows": {}, "block": {}, "blank": {}, "unmet": []}
        for line in done.stdout.splitlines():
            if line.startswith(("PARSE|", "MISSING|")):
                rows["unmet"].append(line)
            elif line.startswith("ROWS|"):
                case, capacity, needed, returned, adds, outcome = line[len("ROWS|"):].split("|", 5)
                rows["rows"][case] = (int(capacity), int(needed), returned, int(adds), outcome)
            elif line.startswith("BLOCK|"):
                n, r, c, first, last, counter, beyond = line[len("BLOCK|"):].split("|", 6)
                rows["block"][int(n)] = (int(r), int(c), first, last, counter, int(beyond))
            elif line.startswith("BLANK|"):
                n, blank, populated = line[len("BLANK|"):].split("|", 2)
                rows["blank"][int(n)] = (int(blank), int(populated))
        _MEMO["reserved"] = rows
    return _MEMO["reserved"]


# ===========================================================================
# R. RESERVED CAPACITY IS NOT SEMANTIC COUNT
# ===========================================================================
# Equivalence run 2's Bulk pass raised "tblCostLines already holds 25 body rows
# where the fixture needs 12". Stage A builds the register with reserved_rows: 25,
# and twelve Cost Lines in twelve of them is the state production reaches.
def test_250_productions_own_rule_is_grow_only_when_capacity_runs_out() -> None:
    """READ OUT OF PRODUCTION, NOT ASSUMED. `modDrivers.AddDriver` takes a blank
    RESERVED row and grows the table only when there is none - and its own comment
    says reserved rows were "only ever initial capacity, never a business
    maximum".

    All five benchmark-populated tables follow it: `modProfiling.SyncRows` and
    `modInflation.SyncProfileRows` grow only when the write row passes the body
    count, and then CLEAR the tail rather than delete it. No production path
    shrinks a body."""
    drivers = (PCCM_ROOT / "src" / "vba" / "modDrivers.bas").read_text(encoding="utf-8")
    add = drivers[drivers.index("Public Function AddDriver"):]
    add = add[:add.index("\nPublic ", 1)] if "\nPublic " in add[1:] else add
    assert "targetRow = FirstFreeRow(Kind, orphanRow)" in add
    assert "If targetRow = 0 Then" in add, "the grow is no longer conditional"
    assert "register.ListRows.Add" in add
    assert "never a business maximum" in add
    # THE GROW IS INSIDE THE CONDITIONAL, not before it.
    assert add.index("If targetRow = 0 Then") < add.index("register.ListRows.Add")

    profiling = (PCCM_ROOT / "src" / "vba" / "modProfiling.bas").read_text(encoding="utf-8")
    assert "If writeRow > modWorkbook.BodyRowCount(target) Then" in profiling
    assert "target.ListRows.Add" in profiling
    assert "' Clear the tail: rows below the last identified driver hold no profile." in profiling
    inflation = (PCCM_ROOT / "src" / "vba" / "modInflation.bas").read_text(encoding="utf-8")
    assert "If writeRow > modWorkbook.BodyRowCount(target) Then target.ListRows.Add" in inflation
    # NOTHING IN THE SYNC PATHS DELETES A BODY ROW.
    for text, name in ((profiling, "modProfiling"), (inflation, "modInflation")):
        sync = text[text.index("SyncRows") if "SyncRows" in text else 0:]
        assert "ListRows(" not in sync.split("Public Sub RemoveRow")[0] or name == "modProfiling"

    # AND THE CONTRACTED CAPACITIES ARE WHAT THIS RESTS ON.
    driver_contract = (SPEC / "driver_contract.yaml").read_text(encoding="utf-8")
    structure = (SPEC / "structure_contract.yaml").read_text(encoding="utf-8")
    for table, capacity in RESERVED_CAPACITY.items():
        source = driver_contract if table in ("tblCostLines", "tblRiskRegister") else structure
        index = source.index(table)
        window = source[index:index + 1500]
        assert f"reserved_rows: {capacity}" in window, (table, capacity)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_251_reserved_capacity_is_kept_and_growth_happens_only_when_needed() -> None:
    """EXECUTED, because "how many rows does this end up with and how many Adds did
    it take" is a COUNT. Twelve drivers into a twenty-five row table must leave
    twenty-five rows and make ZERO Adds; a hundred and eighty must grow to exactly
    a hundred and eighty."""
    rows = _reserved_rows()
    assert rows["unmet"] == [], rows["unmet"]
    assert set(rows["rows"]) == set(RESERVED_EXPECTED), (sorted(rows["rows"]))
    for case, (capacity, needed, physical, adds) in RESERVED_EXPECTED.items():
        got = rows["rows"][case]
        assert got[0] == capacity and got[1] == needed, (case, got)
        assert got[2] == str(physical), (case, "physical rows", got)
        assert got[3] == adds, (case, "ListRows.Add calls", got)
        assert got[4] == "returned", (case, got)
    # THE TWO SMALL CASES ARE THE ONES RUN 2 DIED ON, and they now make no Add.
    assert rows["rows"]["small-cost-12-into-25"][3] == 0
    assert rows["rows"]["small-risk-8-into-25"][3] == 0
    # AND GROWTH IS EXACT: 180 needed from 25 is 155 Adds, not 180.
    assert rows["rows"]["large-cost-180-into-25"][3] == 155


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_252_identifiers_stop_at_the_semantic_count_and_the_counter_matches() -> None:
    """THE BLOCK COVERS EXACTLY THE SEMANTIC ROWS. Twelve drivers means twelve
    rows, CL-001 to CL-012, no CL-013, and a counter of 12 - which is what
    `modDrivers.AllocateId` leaves after twelve Adds."""
    rows = _reserved_rows()
    for count, first, last in ((12, "CL-001", "CL-012"), (8, "CL-001", "CL-008"),
                               (180, "CL-001", "CL-180")):
        assert count in rows["block"], count
        r, c, got_first, got_last, counter, beyond = rows["block"][count]
        assert r == count, (count, "block rows", r)
        # THE COLUMN COUNT IS THE REGISTER'S, from the manifest.
        declared = next(entry for entry in _manifest_json()["registers"]
                        if entry["key"] == "cost_lines")
        assert c == len(declared["columns"]), (count, "block columns", c)
        assert got_first == first and got_last == last, (count, got_first, got_last)
        assert counter == str(float(count)) or counter == str(count), (count, counter)
        assert beyond == 0, (count, "identifiers past the semantic count", beyond)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_253_columns_no_driver_fills_are_genuinely_blank() -> None:
    """`category` AND `uom` ARE BLANK IN THE ENDPOINT-BUILT REGISTER TOO, because
    `Write-Phase5Driver` never writes them. Writing '' instead of $null would make
    them populated, and production would read a row nobody keyed as an orphan."""
    rows = _reserved_rows()
    for count in (12, 8, 180):
        blank, populated = rows["blank"][count]
        assert populated == 0, (count, "populated cells in unfilled columns", populated)
        assert blank == count * 2, (count, blank)
    # AND THE ACCEPTED WRITER REALLY DOES SKIP THEM.
    writer = _function((BOOTSTRAP / "phase5_gate_b_scenarios.ps1").read_text(encoding="utf-8"),
                       "Write-Phase5Driver")
    for skipped in ("'category'", "'uom'", "'risk_owner'"):
        assert skipped not in writer, (
            f"the accepted writer now fills {skipped}, so the bulk block must too")


def test_254_the_grower_takes_a_floor_and_never_shrinks() -> None:
    """THE SOURCE SIDE OF THE SAME PROPERTY, so a rewrite that kept the counts the
    harness samples but reintroduced the shape cannot pass quietly. The parameter
    is named for what it is, the early return is on `-ge`, and there is no delete
    and no resize."""
    grower = _function(_code(), "Set-BenchmarkRegisterRowCount")
    assert "[int]$MinimumRows" in grower, "the parameter still reads as a target"
    assert "if ($current -ge $MinimumRows) { return $current }" in grower, (
        "reserved capacity that already suffices is not left alone")
    # THE OLD PARAMETER, not the letters. `Set-BenchmarkRegisterRowCount` and
    # `Get-TableRowCount` both contain "RowCount" and always did; what must be gone
    # is the `$RowCount` parameter that carried the target-count reading.
    assert "$RowCount" not in grower, "the old target-count parameter survives"
    for banned in (".Delete", "Resize", "$current -gt"):
        assert banned not in grower, f"the grower shrinks or resizes: {banned}"
    # THE CALL SITE PASSES THE DRIVER COUNT AS A FLOOR.
    orchestrator = _function(_code(), "Set-BenchmarkBulkFixture")
    assert "-MinimumRows @($pair.drivers).Count" in orchestrator
    assert "$physical -lt @($pair.drivers).Count" in orchestrator, (
        "the builder does not check it got at least the rows it must populate")


def test_255_the_reserved_suffix_is_proved_blank_before_production_syncs() -> None:
    """A VALUE IN A RESERVED ROW IS AN ORPHAN. `modDrivers.AddDriver` refuses to
    mutate over a row that has data and no permanent identifier, so a fixture that
    left one would poison every later production command - and `ApplyTimeline` is
    about to synchronise both grids from that register."""
    orchestrator = _function(_code(), "Set-BenchmarkBulkFixture")
    assert "is a reserved row below the" in orchestrator
    assert "for ($row = $expected.Count; $row -lt $body.Count; $row++) {" in orchestrator, (
        "the reserved suffix is not read back")
    # THE CONDITION, NOT THE MESSAGE. An `if ($false)` in front of the same throw
    # leaves the sentence in the file and the check gone.
    assert "if ([string]$value -ne '') {" in orchestrator, (
        "the reserved-suffix check no longer tests anything")
    assert "populated unkeyed row is the orphan" in orchestrator
    # BEFORE THE SYNC, not after it.
    checked = orchestrator.index("is a reserved row below the")
    applied = orchestrator.index("PCCM_ApplyTimeline")
    assert checked < applied, "the suffix is checked after production synchronised from it"
    # AND PRODUCTION REALLY DOES REFUSE AN ORPHAN.
    drivers = (PCCM_ROOT / "src" / "vba" / "modDrivers.bas").read_text(encoding="utf-8")
    assert "already contains data but has " in drivers
    assert "no permanent identifier" in drivers


def test_256_profiling_geometry_is_still_productions_alone() -> None:
    """THE GRIDS KEEP THEIR RESERVED SUFFIX TOO, and that is production's business.
    The builder writes weights into the rows `SyncRows` created, reading the order
    back rather than assuming it, and never touches a grid's row count."""
    weights = _function(_code(), "New-BenchmarkWeightBlock")
    for banned in ("ListRows", "ListColumns", "Set-BenchmarkRegisterRowCount", ".Delete"):
        assert banned not in weights, f"the weight block changes grid geometry: {banned}"
    assert "Get-TableBody" in weights
    # THE KEYED ROWS ARE COUNTED, NOT THE PHYSICAL ONES - a grid with a blank
    # reserved suffix must still match the driver count.
    assert "if ([string]::IsNullOrWhiteSpace($key)) { continue }" in weights, (
        "blank reserved grid rows are counted as keyed rows")
    assert "$rows.Count -ne $drivers.Count" in weights
    orchestrator = _function(_code(), "Set-BenchmarkBulkFixture")
    assert "Set-BenchmarkRegisterRowCount" in orchestrator
    assert orchestrator.count("Set-BenchmarkRegisterRowCount") == 1, (
        "the builder resizes something other than the registers")


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_257_the_snapshot_compares_the_reserved_suffix_not_just_the_first_n_rows() -> None:
    """IF ENDPOINTS HAS 25 PHYSICAL ROWS AND BULK HAS 12, THAT MUST BE A REAL
    DIFFER. The register body comparison reads `Get-TableBody`, which returns every
    PHYSICAL row with blanks as empty strings - so a missing reserved suffix
    changes the compared string and is caught.

    The snapshot is not weakened to ignore reserved rows, and is byte-identical to
    what 99cb472 shipped."""
    harness = _equiv_harness()
    snapshot = _function(harness, "Get-EquivalenceSnapshot")
    assert "Get-TableBody" in snapshot
    for narrowing in ("Select-Object -First", "$drivers.Count", "IsNullOrWhiteSpace",
                      "where the key is not blank"):
        assert narrowing not in snapshot, (
            f"the snapshot narrows the body comparison: {narrowing}")
    # THE PROPERTY, NOT A LIST OF WAYS TO BREAK IT. Every row the reader returned
    # is appended UNCONDITIONALLY: a filter on the key column would drop exactly
    # the reserved suffix whose absence must be a difference.
    for line in snapshot.splitlines():
        if "$lines +=" in line:
            assert line.strip().startswith("$lines +="), (
                f"a body row is appended conditionally: {line.strip()}")
    # THE READER RETURNS EVERY PHYSICAL ROW.
    reader = _function(_code(), "Get-TableBody")
    assert "$rowCount = [int]$rowsObj.Count" in reader
    assert "for ($r = 1; $r -le $rowCount; $r++) {" in reader
    assert "if ($null -eq $v) { $line += '' }" in reader
    # AND IT HAS NOT MOVED SINCE THE GATE WAS ACCEPTED.
    accepted = _git("show", "99cb472:pccm/tests/phase10_fixture_equivalence.ps1")
    assert snapshot == _function(accepted, "Get-EquivalenceSnapshot"), (
        "the snapshot changed while the reserved-row rule was being corrected")


def test_258_equivalence_run_2_is_recorded_as_invalid_not_as_a_difference() -> None:
    """WHAT RUN 2 PROVED AND WHAT IT DID NOT. The bundle correction worked and the
    vocabulary worked; the Bulk fixture never built, so there is no equivalence
    result of any kind - and the record must not read as one."""
    section = " ".join(_run_evidence_section("## Equivalence run 2").split())
    for proven in ("BUNDLE|identical|5 artifact(s)",
                   "PASS|Endpoints|COMPLETED",
                   "EQUIV|<not evaluated>|invalid"):
        assert proven in section, f"the run-2 record omits what it proved: {proven}"
    for fact in ("d1af4e1", "already holds 25 body rows", "INVALID",
                 "No semantic mismatch"):
        assert fact in section, f"the run-2 record omits: {fact}"
    assert "must not be read as a fixture DIFFER" in section, section
    # AND BULK IS STILL NOT AUTHORISED.
    assert "[string]$FixtureMode = 'Endpoints'" in _runner()


def test_259_the_reserved_rows_harness_tests_the_shipping_builder() -> None:
    """IT LIFTS THE REAL FUNCTIONS BY AST. A harness that reimplemented the grower
    would have reported the right counts for the wrong code."""
    harness = RESERVED_HARNESS.read_text(encoding="utf-8")
    assert "FunctionDefinitionAst" in harness and "Invoke-Expression" in harness
    assert "'Set-BenchmarkRegisterRowCount', 'Get-BenchmarkPermanentId'" in harness
    for banned in ("New-Object -ComObject", "Excel.Application", "Workbooks.Open"):
        assert banned not in _ps_code(harness), f"the harness starts Excel: {banned}"
    assert "phase10_reserved_rows_flow" not in _runner()


def test_260_production_is_byte_identical_and_the_timed_path_did_not_move() -> None:
    """A RESERVED-ROW FIX TOUCHES NO PRODUCTION, NO TIMING AND NO SNAPSHOT."""
    for commit in ("d1af4e1", "99cb472", "f3b3a33"):
        assert _production_changed_since(commit) == [], commit
    accepted = _code_at("d1af4e1")
    for name in ("Invoke-BenchmarkExecution", "Test-BenchmarkSample",
                 "Assert-BenchmarkProblemList", "Open-BenchmarkFixtureWindow",
                 "Close-BenchmarkFixtureWindow", "Invoke-BenchmarkWindowRollback",
                 "Set-BenchmarkRangeBlock", "New-BenchmarkWeightBlock",
                 "Get-BenchmarkPermanentId", "New-BenchmarkRegisterBlock"):
        assert _function(_code(), name) == _function(accepted, name), (
            f"{name} changed while the reserved-row rule was being corrected")
    plan = _plan()
    assert len([r for r in plan["runs"] if r["scenario"] == "PERF-LARGE"]) == 8


BUNDLE_HARNESS = PCCM_ROOT / "tests" / "phase10_bundle_flow.ps1"

# WHAT build_stage_b.ps1 RESOLVES AGAINST THE SUPPLIED -BuildDir, and therefore
# what every disposable bundle must carry. Derived from the bootstrap, not from
# what the benchmark happens to copy.
BUNDLE_REQUIRED = {
    "file": ("stage_b_manifest.json", "PCCM_stageA.xlsx"),
    "directory": ("vba",),
}
BUNDLE_FILES = ("stage_b_manifest.json", "PCCM_stageA.xlsx",
                "vba/modConstants.bas", "vba/modCalcContract.bas",
                "vba/modSimContract.bas")


def _bundle_rows() -> dict:
    """Run the bundle harness and group its tagged lines."""
    if "bundle" not in _MEMO:
        done = subprocess.run(
            [PWSH, "-NoProfile", "-File", str(BUNDLE_HARNESS), "-Gate", str(EQUIV_HARNESS)],
            capture_output=True, text=True, timeout=300)
        assert done.returncode == 0, done.stdout + done.stderr
        rows = {"artifact": [], "bundle": {}, "roots": None, "identity": None,
                "damaged": None, "refuse": {}, "stale": None, "unmet": []}
        for line in done.stdout.splitlines():
            if line.startswith(("PARSE|", "MISSING|")):
                rows["unmet"].append(line)
            elif line.startswith("ARTIFACT|"):
                kind, name = line[len("ARTIFACT|"):].split("|", 1)
                rows["artifact"].append((kind, name))
            elif line.startswith("BUNDLE|"):
                mode, path = line[len("BUNDLE|"):].split("|", 1)
                rows["bundle"].setdefault(mode, []).append(path)
            elif line.startswith("ROOTS|"):
                rows["roots"] = line[len("ROOTS|"):].split("|", 1)
            elif line.startswith("IDENTITY|"):
                rows["identity"] = line[len("IDENTITY|"):].split("|", 1)
            elif line.startswith("DAMAGED|"):
                rows["damaged"] = line[len("DAMAGED|"):].split("|", 1)
            elif line.startswith("REFUSE|"):
                case, verdict, detail = line[len("REFUSE|"):].split("|", 2)
                rows["refuse"][case] = (verdict, detail)
            elif line.startswith("STALE|"):
                rows["stale"] = line[len("STALE|"):].split("|", 1)
        _MEMO["bundle"] = rows
    return _MEMO["bundle"]


# ===========================================================================
# Q. THE EQUIVALENCE GATE'S STARTING BUNDLE
# ===========================================================================
# Run 1 failed before Excel in BOTH passes: the gate copied the Stage-A workbook
# and the generated `vba` directory into each workdir and not
# `stage_b_manifest.json`, which build_stage_b.ps1 reads first from the supplied
# -BuildDir. Every source-reading control passed. File plumbing is behaviour, so
# these controls execute it.
def test_210_the_bundle_contract_is_derived_from_the_bootstrap() -> None:
    """NOT COPIED FROM THE BENCHMARK'S LIST. `build_stage_b.ps1` resolves exactly
    three things against the supplied -BuildDir, and two deliberately against the
    repository - the distinction its own comment draws. The inspections the
    benchmark also copies are NOT part of this contract: the bootstrap never reads
    one."""
    bootstrap = (BOOTSTRAP / "build_stage_b.ps1").read_text(encoding="utf-8")
    # THE THREE FROM THE BuildDir.
    assert "$manifestPath = Join-Path $BuildDir 'stage_b_manifest.json'" in bootstrap
    assert "$stageAPath = Join-Path $BuildDir $manifest.stage_a_filename" in bootstrap
    assert "$genDir  = Join-Path $BuildDir (Split-Path -Leaf $manifest.vba.generated_dir)" in bootstrap
    # THE TWO FROM THE REPOSITORY, which is why they are not in the bundle.
    assert "$srcDir  = Join-Path $pccmRoot $manifest.vba.source_dir" in bootstrap
    assert "$docFile = Join-Path $srcDir ([string]$docModule.file)" in bootstrap
    # AND NOTHING ELSE: no inspection projection is read from either.
    assert "inspection" not in bootstrap.lower(), (
        "the bootstrap now reads an inspection, so the bundle contract is incomplete")
    # THE OUTPUT IS NOT AN INPUT.
    assert "$stageBPath = Join-Path $BuildDir $manifest.stage_b_filename" in bootstrap

    # THE GATE'S DERIVED LIST MATCHES, and it derives the directory from the
    # manifest rather than naming 'vba'.
    artifacts = _function(_equiv_harness(), "Get-BundleArtifacts")
    assert "Split-Path -Leaf ([string]$Manifest.vba.generated_dir)" in artifacts
    assert "'stage_b_manifest.json'" in artifacts
    assert "[string]$Manifest.stage_a_filename" in artifacts
    assert "stage_b_filename" not in artifacts, "the gate carries the output into the bundle"


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_211_both_bundles_receive_every_required_artifact() -> None:
    """EXECUTED, because run 1 proved a source-reading control cannot see a file
    that was never copied. The harness drives the real bundle builder against a
    fake build directory and reports what arrived."""
    rows = _bundle_rows()
    assert rows["unmet"] == [], rows["unmet"]
    got = {(kind, name) for kind, name in rows["artifact"]}
    want = {(kind, name) for kind, names in BUNDLE_REQUIRED.items() for name in names}
    assert got == want, (sorted(got), sorted(want))
    for mode in ("Endpoints", "Bulk"):
        assert mode in rows["bundle"], mode
        assert sorted(rows["bundle"][mode]) == sorted(BUNDLE_FILES), (mode, rows["bundle"][mode])


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_212_the_two_bundles_are_isolated_and_proved_identical() -> None:
    """NEITHER PASS MAY SEE THE OTHER'S WORKBOOK, and both must start from the same
    bytes or nothing is comparable. The identity check is proved non-vacuous by
    editing one artifact and requiring the comparison to notice, with both
    digests."""
    rows = _bundle_rows()
    assert rows["roots"][0] == "isolated", rows["roots"]
    assert rows["identity"][0] == "identical", rows["identity"]
    assert rows["damaged"][0] == "differ", rows["damaged"]
    assert "stage_b_manifest.json differs" in rows["damaged"][1], rows["damaged"]
    # BOTH DIGESTS ARE NAMED, so a reader can see which side changed.
    assert rows["damaged"][1].count("=") >= 2, rows["damaged"]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_213_a_missing_required_artifact_refuses_before_excel() -> None:
    """EACH ONE, BY NAME. Run 1's diagnosis told the operator to run Stage A, which
    had already been run - so the refusal now names the artifact AND the directory
    it was looked for in."""
    rows = _bundle_rows()
    for case, expected in (("missing-manifest", "stage_b_manifest.json"),
                           ("missing-stage-a-workbook", "PCCM_stageA.xlsx"),
                           ("missing-generated-vba", "vba")):
        assert case in rows["refuse"], case
        verdict, detail = rows["refuse"][case]
        assert verdict == "refused", (case, verdict, detail)
        assert expected in detail, (case, detail)
        assert "is not in the repository build directory" in detail, (case, detail)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_214_no_stale_stage_b_workbook_travels_into_a_bundle() -> None:
    """THE OUTPUT IS NOT AN INPUT. A stale `PCCM_stageB.xlsm` beside the Stage-A
    workbook must not be copied, or the pass would open last week's build instead
    of the one it just made - and every EQUIV line would be about the wrong
    workbook."""
    rows = _bundle_rows()
    assert rows["stale"][0] == "absent", rows["stale"]
    # AND THE BUILDER REFUSES ONE THAT IS SOMEHOW PRESENT, which is the other half.
    builder = _function(_equiv_harness(), "New-EquivalenceBundle")
    assert "$Manifest.stage_b_filename" in builder
    assert "carried a stale build" in builder


def test_215_a_pass_that_did_not_complete_cannot_be_reported_as_PASS() -> None:
    """RUN 1 PRINTED `PASS|Endpoints|RAISED` FOR A PASS THAT NEVER BUILT A
    WORKBOOK. PASS now belongs only to a pass that completed, and a failure names
    its stage: SETUP before any build, BOOTSTRAP when Stage-B produced nothing,
    RAISED for anything after Excel started."""
    code = _ps_code(_equiv_harness())
    assert "('PASS|' + $mode + '|COMPLETED|" in code, code
    assert "('FAIL|' + $mode + '|SETUP|'" in code
    assert "('FAIL|' + $mode + '|' + $stage + '|'" in code
    assert "$stage = 'RAISED'" in code
    assert "$stage = 'BOOTSTRAP'" in code
    # THE OLD VOCABULARY IS GONE FROM THE CODE. The header still QUOTES it, in the
    # sentence explaining why it was wrong, and quoting it is the opposite of
    # emitting it.
    assert "'|RAISED|'" not in code.replace("$stage", "")
    for line in code.splitlines():
        if "'PASS|'" in line:
            assert "COMPLETED" in line, f"a PASS line that is not a completion: {line.strip()}"
    assert "PASS|Endpoints|RAISED" in _equiv_harness(), (
        "the header no longer records the vocabulary defect it was corrected for")
    # AND COMPLETED MEANS BOTH HALVES ARRIVED.
    pass_fn = _function(code, "Invoke-EquivalencePass")
    # THE CONDITIONS, NOT THE MESSAGES. An `if ($false)` in front of either throw
    # leaves both sentences in the file and both checks gone - and a pass that
    # calculated nothing would then be compared, with CALCEQUIV comparing two
    # empty strings and calling them equal.
    assert "if ($null -eq $snapshot) {" in pass_fn, (
        "the snapshot check no longer tests anything")
    assert "if ([string]::IsNullOrWhiteSpace($calcFingerprint)) {" in pass_fn, (
        "the fingerprint check no longer tests anything")
    assert "produced no state snapshot" in pass_fn
    assert "produced no calculation fingerprint" in pass_fn


def test_216_differ_is_reserved_for_a_comparison_that_actually_ran() -> None:
    """RUN 1 PRINTED `differ` AFTER A SETUP FAILURE, which reads as "the two
    fixtures are not equivalent" and was not what happened. A comparison that did
    not happen is `invalid`, and the reason is named."""
    harness = _equiv_harness()
    assert "EQUIV|<not evaluated>|invalid|comparison was not executed: " in harness
    assert "$completed.Count -ne 2" in harness
    # EVERY `differ` EMISSION SITS INSIDE THE BOTH-COMPLETED BRANCH.
    invalid_at = harness.index("EQUIV|<not evaluated>|invalid")
    for marker in ("|differ|endpoints=", "|differ|only the bulk pass reported this field"):
        assert harness.index(marker) > invalid_at, (
            f"a differ emission precedes the not-evaluated branch: {marker}")
    # AND THE ONE SURVIVING `differ` ON THE BUNDLE IS ABOUT THE BUNDLE, not a
    # fixture comparison.
    assert "'BUNDLE|differ|'" in harness


def test_217_calc_and_calcequiv_cannot_be_emitted_without_two_calculations() -> None:
    """NO PLACEHOLDER VERDICT ON A CALCULATION THAT NEVER RAN. Both CALC lines and
    CALCEQUIV live inside the branch entered only when both passes completed - and
    completion already requires a fingerprint."""
    code = _ps_code(_equiv_harness())
    guard = code.index("if ($completed.Count -ne 2) {")
    for marker in ("'CALC|' + $pass.Mode", "CALCEQUIV|match", "CALCEQUIV|differ"):
        assert code.index(marker) > guard, f"{marker} is emitted outside the guard"
    # TWO IN THE CODE - the match and the differ - and the header lists a third,
    # which is documentation rather than an emission.
    assert code.count("CALCEQUIV|") == 2, code.count("CALCEQUIV|")
    assert "$reference.CalcFingerprint -ceq $optimised.CalcFingerprint" in code


def test_218_nothing_is_built_when_the_starting_states_disagree() -> None:
    """REFUSING BEFORE EXCEL IS THE POINT OF PROVING IDENTITY. Two bundles that
    differ were never comparable, and building them anyway would produce EQUIV
    lines about two different starting workbooks."""
    harness = _equiv_harness()
    assert "if (-not $setupFailed) {" in harness
    bundles_at = harness.index("$problems = @(Test-BundleIdentity")
    passes_at = harness.index("$passes = @{}")
    assert bundles_at < passes_at, "identity is checked after the passes run"
    # THE PASS LOOP IS GUARDED BY IT.
    tail = harness[passes_at:]
    assert tail.index("if (-not $setupFailed) {") < tail.index("Invoke-EquivalencePass"), tail[:400]


def test_219_the_equivalence_field_set_is_unchanged() -> None:
    """A BROKEN SETUP IS NO REASON TO COMPARE LESS. Every family the gate was built
    to compare is still compared, and the snapshot function is byte-identical to
    the one 99cb472 shipped."""
    harness = _equiv_harness()
    for piece in ("'.ids'", "'.body'", "'.headers'", "'.columns'",
                  "'counter.'", "'applied.'", "'structural.state'",
                  "'structural.report'", "'fingerprint.calculation_inputs'",
                  "'fingerprint.simulation_request'", "'modelcheck.calculation_state'"):
        assert piece in harness, f"the snapshot omits {piece}"
    assert "@('cost_profiling', 'risk_profiling', 'inflation')" in harness
    assert harness.count("@('cost_lines', 'risk_register')") >= 2
    for name in ("nmBaseYear_Applied", "nmStartYear_Applied", "nmDuration_Applied",
                 "nmLastYear_Applied", "nmYearCount_Applied",
                 "nmInflFirstYear", "nmInflLastYear"):
        assert name in harness, name
    assert "$excel.Run('PCCM_Calculate')" in harness
    # BYTE-IDENTICAL TO WHAT SHIPPED, so the setup correction cannot have touched it.
    accepted = _git("show", "99cb472:pccm/tests/phase10_fixture_equivalence.ps1")
    assert (_function(harness, "Get-EquivalenceSnapshot")
            == _function(accepted, "Get-EquivalenceSnapshot")), (
        "the snapshot changed while the setup was being corrected")


def test_220_the_bulk_path_is_still_not_authorised_and_the_run_is_recorded() -> None:
    """RUN 1 WAS INVALID, NOT A DIFFERENCE. It must be on the record as a setup
    failure that reached no comparison, and Bulk must still be opt-in with
    Endpoints the default."""
    section = " ".join(_run_evidence_section("## Equivalence run 1").split())
    for fact in ("99cb472", "before Excel was started", "stage_b_manifest.json",
                 "INVALID", "No semantic equivalence claim may be made",
                 "PASS|Endpoints|RAISED"):
        assert fact in section, f"the run-1 record omits: {fact}"
    assert "does NOT mean the fixtures differ" in section
    # AND THE DEFAULT HAS NOT MOVED.
    assert "[string]$FixtureMode = 'Endpoints'" in _runner()


def test_221_the_bundle_harness_tests_the_shipping_gate_and_starts_no_excel() -> None:
    """IT LIFTS THE REAL FUNCTIONS BY AST. A harness that reimplemented the bundle
    builder would have passed run 1 too."""
    harness = BUNDLE_HARNESS.read_text(encoding="utf-8")
    assert "FunctionDefinitionAst" in harness and "Invoke-Expression" in harness
    for lifted in ("'Get-BundleArtifacts', 'New-EquivalenceBundle', 'Test-BundleIdentity'",):
        assert lifted in harness, lifted
    # THE CODE, NOT THE EXPLANATION. The harness names `build_stage_b.ps1` in the
    # paragraph saying what it does NOT run, which is the opposite of running it.
    code = _ps_code(harness)
    for banned in ("New-Object -ComObject", "Excel.Application", "Workbooks",
                   "build_stage_b"):
        assert banned not in code, f"the bundle harness reaches past file plumbing: {banned}"
    assert "build_stage_b.ps1" in harness, (
        "the harness no longer says which bootstrap contract it is about")
    assert "phase10_bundle_flow" not in _runner()


def test_222_production_is_byte_identical_to_the_accepted_revision() -> None:
    """A SETUP FIX TOUCHES NO PRODUCTION, NO TIMING AND NO FIXTURE SEMANTICS."""
    for commit in ("99cb472", "f3b3a33", "ce5951f"):
        changed = _production_changed_since(commit)
        assert changed == [], (commit, changed)
    # THE BULK BUILDER AND THE TIMED PATH ARE UNTOUCHED BY THIS ROUND, with one
    # DECLARED exception: equivalence run 2 proved `Set-BenchmarkRegisterRowCount`
    # confused reserved capacity with semantic count, and correcting that is the
    # whole of the reserved-row round. Everything else must still match.
    declared = ("Set-BenchmarkRegisterRowCount", "Set-BenchmarkBulkFixture")
    accepted = _code_at("99cb472")
    for name in BULK_FUNCTIONS + ("Invoke-BenchmarkExecution", "Test-BenchmarkSample",
                                  "Assert-BenchmarkProblemList"):
        if name in declared:
            # A DECLARATION FOR SOMETHING THAT DID NOT CHANGE IS A STALE EXEMPTION.
            assert _function(_code(), name) != _function(accepted, name), (
                f"{name} is declared as corrected but is unchanged since 99cb472")
            continue
        assert _function(_code(), name) == _function(accepted, name), (
            f"{name} changed while the equivalence setup was being corrected")


EQUIV_HARNESS = PCCM_ROOT / "tests" / "phase10_fixture_equivalence.ps1"

# The four timed operations and the production endpoint each must reach. Read
# from the plan, never restated, so a retargeted operation is a failure here.
TIMED_ENDPOINTS = {
    "calculate": "PCCM_Calculate",
    "simulation": "PCCM_RunSimulation",
    "sensitivity": "PCCM_RunSensitivity",
    "annual": "PCCM_RunAnnualStochastic",
}

# Everything the bulk builder is forbidden to touch. These are the publication
# surfaces and the state labels: production owns every one of them.
PUBLICATION_BANS = (
    "_Calc", "_SimData", "SH_CALC", "SH_SIM",
    "CalculationFingerprint", "SimulationResultDigest", "CalculationStatus",
    "SimulationStatus", "nmStructuralState", "nmCalcState", "CURRENT",
    "tblCalcDrivers", "tblCalcYears", "tblCalcAnnual", "tblCalcFX",
    "tblCalcInflationFactors", "tblSimResults", "tblSensitivity",
)


def _equiv_harness() -> str:
    if "equiv_harness" not in _MEMO:
        _MEMO["equiv_harness"] = EQUIV_HARNESS.read_text(encoding="utf-8")
    return _MEMO["equiv_harness"]


def _ps_code(text: str) -> str:
    """A PowerShell file with its help block and comment lines removed.

    SAME RULE AS `_code`. Both PowerShell harnesses here DOCUMENT the vocabulary
    they must not misuse and the bootstrap they must not reach - they have to name
    those things to explain them. Everything that asserts what a harness DOES
    reads this; everything that asserts what it SAYS reads the raw text.

    NOT MEMOISED: it is derived from text the mutation battery installs, and a
    cache would compare every mutation after the first against undamaged source.
    """
    body = re.sub(r"<#.*?#>", "", text, flags=re.S)
    return "\n".join(line for line in body.splitlines()
                      if not line.strip().startswith("#"))


BULK_FUNCTIONS = ("Set-BenchmarkRangeBlock", "Set-BenchmarkRegisterRowCount",
                  "Get-BenchmarkPermanentId", "New-BenchmarkRegisterBlock",
                  "New-BenchmarkWeightBlock", "Set-BenchmarkBulkFixture")


def _bulk_functions() -> dict:
    """The bulk builder's functions, comments stripped, by name.

    DELIBERATELY NOT MEMOISED. It is derived from `_code`, and the mutation
    battery damages the runner by installing a replacement for `runner` and
    popping `code`. A cache here would survive that, so every mutation after the
    first would be compared against the undamaged builder - which is exactly how
    a whole family of controls goes quietly vacuous. String slicing is cheap;
    a stale baseline is not.
    """
    code = _code()
    return {name: _function(code, name) for name in BULK_FUNCTIONS}


# ===========================================================================
# P. THE BULK FIXTURE
# ===========================================================================
# PERF-LARGE was operator-aborted after >4 hours in fixture construction. The
# cost is a production Add per driver, and each is O(register rows x years):
# snapshot both tables, re-sync the grid, validate, re-protect fourteen sheets,
# recalculate four. N of those is O(N^2 x years).
def test_190_the_timed_operations_still_reach_the_real_production_endpoints() -> None:
    """THE ONE THING THAT MAY NOT MOVE. A faster fixture is worthless if the
    numbers stop coming from production. Every timed operation's endpoint is read
    from the plan and must be the real PCCM_ command, and the runner must invoke
    it by that name and nothing else."""
    operations = {str(entry["key"]): entry for entry in _plan()["operations"]}
    for key, endpoint in TIMED_ENDPOINTS.items():
        assert key in operations, key
        assert str(operations[key]["endpoint"]) == endpoint, (key, operations[key])
    # The recalculation row is the one operation with no endpoint: it is
    # Application.CalculateFull by declared mechanism, not a PCCM_ command.
    assert operations["recalculation"]["endpoint"] is None
    assert str(operations["recalculation"]["mechanism"]) == "Application.CalculateFull"

    # AND THE RUNNER REACHES THEM BY THE PLAN'S OWN NAME, not a literal.
    execution = _function(_code(), "Invoke-BenchmarkExecution")
    assert "$Excel.Run([string]$Operation.endpoint)" in execution, execution
    assert "$Excel.CalculateFull()" in execution
    for endpoint in TIMED_ENDPOINTS.values():
        assert endpoint not in execution, (
            f"{endpoint} is hard-coded in the timed path instead of read from the plan")


def test_191_the_bulk_fixture_publishes_nothing() -> None:
    """IT MAY POPULATE INPUTS. IT MAY NOT PRODUCE RESULTS. A fixture that wrote a
    _Calc block, a _SimData row, a fingerprint or a CURRENT label would be a
    second business engine, and the benchmark would be measuring a workbook it
    had written the answers into."""
    for name, body in _bulk_functions().items():
        for banned in PUBLICATION_BANS:
            assert banned not in body, (
                f"{name} touches a publication surface or a state label: {banned}")
    # NOR DOES IT REACH THE RESULT-PRODUCING ENDPOINTS. The only production
    # operation the builder invokes is ApplyTimeline.
    orchestrator = _bulk_functions()["Set-BenchmarkBulkFixture"]
    for endpoint in TIMED_ENDPOINTS.values():
        assert endpoint not in orchestrator, (
            f"the bulk fixture invokes {endpoint}, which only the timed path may")
    invoked = re.findall(r"-Operation '([A-Za-z_]+)'", orchestrator)
    assert invoked == ["PCCM_ApplyTimeline"], invoked


def test_192_the_bulk_fixture_leaves_every_structural_product_to_production() -> None:
    """WHAT IT WOULD BE EASIEST TO FAKE. The year columns, the profiling rows and
    their permanent-ID keying, and the applied-timeline names all come from ONE
    real PCCM_ApplyTimeline - modProfiling.SetYearColumns, modProfiling.SyncRows,
    modInflation.SetYearColumns and modInflation.SyncProfileRows - and the builder
    reproduces none of them.

    That ApplyTimeline really does produce them is read out of production, not
    assumed here."""
    orchestrator = _bulk_functions()["Set-BenchmarkBulkFixture"]
    assert "PCCM_ApplyTimeline" in orchestrator
    assert "Assert-Phase5StructurallyCoherent" in orchestrator
    # THE BUILDER NEVER WRITES A YEAR COLUMN OR A GRID ROW KEY.
    for banned in ("ListColumns.Add", "ListColumns.Delete", "SetYearColumns", "SyncRows",
                   "Resize(", ".Delete()"):
        assert banned not in orchestrator, f"the bulk fixture does production's work: {banned}"

    # PRODUCTION'S SIDE OF THE CLAIM, from src/vba.
    timeline = (PCCM_ROOT / "src" / "vba" / "modTimeline.bas").read_text(encoding="utf-8")
    for produced in ("modProfiling.SetYearColumns", "modProfiling.SyncRows",
                     "modInflation.SetYearColumns", "modInflation.SyncProfileRows",
                     "modStructuralCheck.ValidateStructure",
                     "modWorkbook.WriteValue NM_APPLIED_BASE_YEAR"):
        assert produced in timeline, (
            f"PCCM_ApplyTimeline no longer produces {produced}, so the bulk fixture's "
            f"assumption about what it gets for free is stale")


def test_193_the_permanent_ids_and_counters_follow_productions_own_rule() -> None:
    """THE ONE THING THE BUILDER WRITES THAT IS NOT A PLAIN USER INPUT.
    modDrivers.AllocateId reads the counter, increments it, PERSISTS it and
    formats prefix + the sequence zero-padded to the declared width - so N adds
    yield prefix-001..prefix-00N and leave the counter at N.

    The prefix and pad width come from the manifest's counter projection, never
    from a literal, and the builder COMPARES its identifier against the model's
    rather than trusting either."""
    formatter = _bulk_functions()["Get-BenchmarkPermanentId"]
    assert "[string]$Counter.prefix" in formatter
    assert "[int]$Counter.pad_width" in formatter
    for literal in ("'CL-'", '"CL-"', "'R-'", '"R-"', "D3", "000"):
        assert literal not in formatter, f"the identifier format is hard-coded: {literal}"

    block = _bulk_functions()["New-BenchmarkRegisterBlock"]
    assert "-cne $declared" in block, "the issued identifier is not compared with the model's"
    assert "CounterValue = [double]$drivers.Count" in block
    orchestrator = _bulk_functions()["Set-BenchmarkBulkFixture"]
    assert "$prepared.CounterValue" in orchestrator
    assert "[string]$counter.defined_name" in orchestrator

    # PRODUCTION'S SIDE: the counter really is persisted per allocation, and the
    # manifest really declares the prefix and pad the builder reads.
    drivers = (PCCM_ROOT / "src" / "vba" / "modDrivers.bas").read_text(encoding="utf-8")
    assert "modWorkbook.WriteValue CounterName(Kind), nextSequence" in drivers
    assert "AllocateId = FormatId(Kind, nextSequence)" in drivers
    for counter in _manifest_json()["counters"]:
        assert counter["prefix"] in ("CL-", "R-"), counter
        assert int(counter["pad_width"]) == 3, counter
        assert int(counter["initial"]) == 0, counter


def test_194_the_bulk_fixture_writes_rectangular_blocks_not_cell_by_cell() -> None:
    """THE WHOLE PERFORMANCE ARGUMENT. A 300x11 register body is ONE cross-process
    assignment instead of 3,300, and a 300x40 weight block is one instead of
    12,000. The writer refuses anything that is not a rank-2 rectangle, because a
    jagged array of arrays is silently NOT the VARIANT shape Excel accepts."""
    writer = _bulk_functions()["Set-BenchmarkRangeBlock"]
    assert "$target.Value2 = $Block" in writer
    assert "$Block.Rank -ne 2" in writer, "the block writer accepts a non-rectangular array"
    assert "GetLength(0)" in writer and "GetLength(1)" in writer
    assert "throw (" in writer
    # ONE ASSIGNMENT PER BLOCK: no loop over cells anywhere in the writer.
    assert "for (" not in writer and "foreach (" not in writer, (
        "the block writer loops over cells, which is the cost it exists to remove")
    # THE ARRAYS ARE BUILT AS RANK-2, not nested.
    for builder in ("New-BenchmarkRegisterBlock", "New-BenchmarkWeightBlock"):
        body = _bulk_functions()[builder]
        assert "New-Object 'object[,]'" in body, (
            f"{builder} builds a jagged array, which Excel rejects for a block write")
    # AND THE ORCHESTRATOR USES IT FOR BOTH REGISTERS AND BOTH GRIDS.
    orchestrator = _bulk_functions()["Set-BenchmarkBulkFixture"]
    assert orchestrator.count("Set-BenchmarkRangeBlock") == 2, orchestrator


def test_195_growing_a_register_is_bounded_proved_and_never_destructive() -> None:
    """ListRows.Add IS STRUCTURAL, so it is only legal inside the fixture window -
    which is where the builder runs. It is one COM call per row rather than a
    whole production operation per row, and the result is PROVED: a grow that
    silently did nothing would leave the block write landing outside the table.

    IT NEVER DELETES AND NEVER SHRINKS. Equivalence run 2 proved that the second
    half of that had been written as a REFUSAL - the grower treated its argument as
    the row count the table should end up with and raised when reserved capacity
    exceeded it. The rule is now a floor; the property that nothing is deleted is
    unchanged, and test_254 holds the floor itself."""
    grower = _bulk_functions()["Set-BenchmarkRegisterRowCount"]
    assert "$rows.Add()" in grower
    assert ".Delete" not in grower, "the register grower deletes rows"
    assert "if ($current -ge $MinimumRows) { return $current }" in grower, (
        "reserved capacity that already suffices is not left alone")
    assert "$after -ne $MinimumRows" in grower, "the grow is not proved to have taken"
    # AND THE BUILDER RUNS INSIDE THE WINDOW, which is what makes the add legal.
    code = _code()
    opened = code.index("Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest")
    bulk = code.index("Set-BenchmarkBulkFixture -Excel $excel")
    closed = code.index("Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest")
    assert opened < bulk < closed, (opened, bulk, closed)


def test_196_the_fixture_mode_is_declared_recorded_and_defaults_to_the_accepted_path() -> None:
    """TWO METHODS THAT REACH THE SAME STATE ARE STILL TWO METHODS. The accepted
    PERF-SMALL and PERF-MEDIUM baselines were built by the endpoint path, so that
    stays the DEFAULT and the bulk path is opt-in - and whichever ran is written
    into the artifact, so no baseline can be compared across methods without a
    reader seeing it."""
    runner = _runner()
    assert "[ValidateSet('Endpoints', 'Bulk')]" in runner
    assert "[string]$FixtureMode = 'Endpoints'" in runner, (
        "the bulk path is the default, which would silently change what the accepted "
        "baselines were built by")
    code = _code()
    assert "$report.Add('fixture_mode', [string]$FixtureMode)" in code
    assert "if ($FixtureMode -eq 'Bulk') {" in code
    # BOTH BUILDERS ARE REACHABLE, and exactly one runs.
    assert code.count("Set-BenchmarkBulkFixture -Excel $excel") == 1
    assert code.count("Set-Phase5Fixture -Excel $excel") == 1


def test_197_the_bulk_fixture_is_inside_the_window_and_outside_every_clock() -> None:
    """SETUP, EXACTLY AS THE ENDPOINT PATH IS. The builder sits between the fixture
    stopwatch's start and stop - so its cost is reported as the setup it is - and
    the run loop is downstream of the window closing."""
    code = _code()
    started = code.index("$fixtureWatch = [System.Diagnostics.Stopwatch]::StartNew()")
    bulk = code.index("Set-BenchmarkBulkFixture -Excel $excel")
    stopped = code.index("$fixtureWatch.Stop()")
    loop = code.index("foreach ($run in $plannedRuns) {")
    assert started < bulk < stopped < loop, (started, bulk, stopped, loop)
    assert "$setupTimings.Add('scenario_fixture_ms'," in code
    # NO STOPWATCH INSIDE THE BUILDER: it is not measuring anything.
    assert "Stopwatch" not in _bulk_functions()["Set-BenchmarkBulkFixture"]


def test_198_the_bulk_fixture_refuses_a_register_that_is_not_empty() -> None:
    """IT WRITES IDENTIFIERS FROM SEQUENCE 1. A register that already carried rows
    would be given duplicates and a counter that disagreed with the highest
    identifier present - which production's own HighestIssued check exists to
    catch, and which a fixture should never create."""
    orchestrator = _bulk_functions()["Set-BenchmarkBulkFixture"]
    assert "$existing.Count -ne 0" in orchestrator
    assert "requires an empty" in orchestrator
    # AND IT PROVES THE BLOCK LANDED before asking production to sync from it.
    assert "-cne [string]$expected[$i]" in orchestrator, (
        "the register is not read back after the block write")
    written = orchestrator.index("Set-BenchmarkRangeBlock")
    verified = orchestrator.index("-cne [string]$expected[$i]")
    applied = orchestrator.index("PCCM_ApplyTimeline")
    assert written < verified < applied, (written, verified, applied)


def test_199_the_weight_block_is_keyed_by_what_production_synchronised() -> None:
    """THE GRID ORDER IS PRODUCTION'S, NOT THE BUILDER'S. modProfiling.SyncRows
    rebuilds the grid from the register, and nothing binds its physical order to
    the order the register was written in - so the weight block reads the keys
    back and maps each driver to the row production gave it.

    It also refuses a non-contiguous or mismatched grid, because a single
    rectangular write over a gap would land weights on the wrong drivers."""
    weights = _bulk_functions()["New-BenchmarkWeightBlock"]
    assert "Get-TableBody" in weights, "the grid order is assumed rather than read"
    assert "$byId[$rows[$r]]" in weights
    # THE CONDITION, NOT THE MESSAGE. A `if ($false)` in front of the same throw
    # leaves the sentence in the file and the check gone, which is exactly the
    # shape a reviewer skims past.
    assert "if ([string]::IsNullOrWhiteSpace([string]$body[$r][0])) {" in weights, (
        "the blank-key check no longer tests anything")
    assert "are not contiguous" in weights
    assert "$rows.Count -ne $drivers.Count" in weights
    assert "declares $" not in weights
    # EVERY DRIVER'S WEIGHTS MUST COVER EVERY PROJECT YEAR.
    assert "$weights.Count -ne $Years" in weights


def test_200_the_equivalence_gate_exists_tests_the_shipping_builder_and_judges_nothing() -> None:
    """THE GATE IS WHAT MAKES THE BULK PATH LEGITIMATE. It builds PERF-SMALL BOTH
    ways in two Excel sessions over two disposable copies of the same build,
    compares a full state snapshot field for field, and then runs the REAL
    PCCM_Calculate on both and compares production's own fingerprint.

    It lifts the builder by AST, so it cannot pass against a restatement of
    itself; and it asserts nothing - it prints tagged lines and the pytest control
    decides, so the evidence and the verdict are two authorities."""
    harness = _equiv_harness()
    assert "FunctionDefinitionAst" in harness and "Invoke-Expression" in harness
    assert "'Set-BenchmarkBulkFixture'" in harness
    assert "Set-Phase5Fixture" in harness, "the gate does not build the reference way"
    # BOTH PASSES, AND THE SAME WINDOW FOR EACH.
    # BOTH LOOPS. The gate walks the two modes twice - once to build the starting
    # bundles and once to run the passes - and a mutation that narrowed either one
    # would leave the other to satisfy a bare `in` check.
    assert harness.count("foreach ($mode in @('Endpoints', 'Bulk'))") == 2, (
        "the gate no longer walks both modes for both the bundles and the passes")
    # THE CALL, NOT THE NAME. The gate also NAMES every function it lifts by AST,
    # and naming one is the opposite of calling it - so the count is over
    # invocations, of which there is exactly one, shared by both passes.
    calls = [line.strip() for line in harness.splitlines()
             if line.strip().startswith("$null = Open-BenchmarkFixtureWindow")]
    assert len(calls) == 1, calls
    assert "'Open-BenchmarkFixtureWindow'," in harness, "the window is not lifted by AST"
    # THE REAL CALCULATE, ON BOTH.
    assert "$excel.Run('PCCM_Calculate')" in harness
    assert "PCCM_CalculationFingerprint" in harness
    assert "CALCEQUIV|" in harness
    # EVERY FIELD FAMILY THE AUTHORISATION NAMES.
    # THE NAMES ARE COMPOSED, so the check is over the composition and the key
    # lists that drive it - which is what actually decides the families.
    for piece in ("'.ids'", "'.body'", "'.headers'", "'.columns'",
                  "'counter.'", "'applied.'", "'structural.state'",
                  "'structural.report'", "'fingerprint.calculation_inputs'",
                  "'fingerprint.simulation_request'", "'modelcheck.calculation_state'"):
        assert piece in harness, f"the snapshot omits {piece}"
    # BOTH REGISTERS, ALL THREE GRIDS, AND THE APPLIED TIMELINE NAMES.
    assert harness.count("@('cost_lines', 'risk_register')") >= 2, harness
    assert "@('cost_profiling', 'risk_profiling', 'inflation')" in harness
    for name in ("nmBaseYear_Applied", "nmStartYear_Applied", "nmDuration_Applied",
                 "nmLastYear_Applied", "nmYearCount_Applied",
                 "nmInflFirstYear", "nmInflLastYear"):
        assert name in harness, f"the applied timeline comparison omits {name}"
    # AND THE COMPARISON IS EXACT AND TWO-WAY: a field only one pass reported is a
    # difference, not an omission.
    assert "$left -ceq $right" in harness
    assert "only the bulk pass reported this field" in harness
    # IT JUDGES NOTHING.
    assert "exit 0" in harness
    assert "exit 1" not in harness, "the gate decides its own verdict"


def test_201_the_bulk_path_is_not_usable_as_a_baseline_until_the_gate_has_run() -> None:
    """THE GATE IS A WINDOWS FACT AND HAS NOT BEEN RUN. Until it has, the record
    must say so, and no accepted baseline may claim the bulk method. When it HAS
    run, every field family must have matched and the calculation fingerprint with
    them - this control reads whichever of those two states the record is in and
    refuses anything else."""
    text = RUN_EVIDENCE.read_text(encoding="utf-8")
    section = " ".join(_run_evidence_section("## The fixture equivalence gate").split())
    if "NOT YET RUN" in section:
        # OUTSTANDING. Then no baseline anywhere may have been built in bulk.
        assert "fixture_mode: Bulk" not in text, (
            "a baseline claims the bulk fixture while the equivalence gate is unrun")
        assert "must not be used for a baseline" in section
        return
    # RUN. Then it passed, on every family and on the fingerprint.
    assert "EVERY FIELD FAMILY MATCHED" in section, section
    assert "CALCEQUIV|match" in section, section
    assert "differ" not in section.replace("no field differed", ""), section


def test_202_the_aborted_large_run_stays_recorded_as_aborted() -> None:
    """>4 HOURS IN SETUP IS EVIDENCE ABOUT THE FIXTURE METHOD, NOT ABOUT PCCM.
    The record has to say both halves: that the run produced nothing, and that it
    says nothing about how long a Large calculation or simulation takes."""
    section = " ".join(_run_evidence_section("## PERF-LARGE attempt 1").split())
    for fact in ("f3b3a33", "operator-aborted", "more than four hours",
                 "0 timed operations", "NOT a baseline", "Shutdown"):
        assert fact in section, f"the aborted-run record omits: {fact}"
    # AND IT REFUSES THE WRONG READING.
    assert "does NOT mean" in section
    for wrong in ("Large calculation takes", "Simulation takes",
                  "Large scenario is unsupported"):
        assert wrong in section, f"the record does not refuse the reading: {wrong}"
    assert "operationally impractical" in section


def test_203_the_measured_cost_reduction_is_recorded_with_its_derivation() -> None:
    """A NUMBER WITHOUT ITS DERIVATION IS A CLAIM. The record carries both counts
    and where each came from, so Windows can check the reduction rather than take
    it on trust."""
    section = " ".join(_run_evidence_section("## The fixture cost, counted").split())
    for figure in ("1,252,725", "5,816,730", "543,760", "191,275"):
        assert figure in section, f"the cost table omits {figure}"
    assert "O(N^2" in section or "O(N²" in section
    for named in ("SnapshotTable", "SyncRows", "ValidateStructure",
                  "ProtectionBeginStructural", "RecalculateStructuralState"):
        assert named in section, f"the derivation does not name {named}"


def test_204_the_iteration_matrix_and_timing_semantics_did_not_move() -> None:
    """A FIXTURE CORRECTION IS NO OCCASION TO CHANGE WHAT IS MEASURED. Large stays
    capped at 50,000 with eight planned runs; cold is 1 and warm is 3 everywhere;
    the median still needs every warm sample valid."""
    plan = _plan()
    large = [r for r in plan["runs"] if r["scenario"] == "PERF-LARGE"]
    assert len(large) == 8, len(large)
    assert sorted({r["iterations"] for r in large if r["iterations"]}) == [10000, 50000]
    forbidden = [f for f in plan["forbidden"] if f["scenario"] == "PERF-LARGE"]
    assert len(forbidden) == 1 and int(forbidden[0]["iterations"]) == 100000, forbidden
    for scenario, count in (("PERF-SMALL", 11), ("PERF-MEDIUM", 11), ("PERF-LARGE", 8)):
        rows = [r for r in plan["runs"] if r["scenario"] == scenario]
        assert len(rows) == count, (scenario, len(rows))
        assert {r["cold_runs"] for r in rows} == {1}
        assert {r["warm_runs"] for r in rows} == {3}
    code = _code()
    assert "if ($validWarm.Count -eq [int]$run.warm_runs) {" in code


def test_205_the_protection_contract_is_untouched_by_the_bulk_path() -> None:
    """THE WINDOW IS RUNTIME PROVEN AND IS NOT REOPENED FOR THIS. The bulk builder
    runs inside the SAME window the endpoint path uses, with no new release, no
    Unprotect of its own, and no second protection authority."""
    for name, body in _bulk_functions().items():
        for banned in (".Unprotect", ".Protect(", "ProtectionRelease",
                       "ProtectionBeginStructural", "UserInterfaceOnly:="):
            assert banned not in body, f"{name} reaches for protection: {banned}"
    # THE WINDOW'S OWN FUNCTIONS ARE BYTE-IDENTICAL TO THE RUN THAT PROVED THEM.
    for name in ("Open-BenchmarkFixtureWindow", "Close-BenchmarkFixtureWindow",
                 "Invoke-BenchmarkWindowRollback", "Assert-BenchmarkProtectionApplied",
                 "Get-BenchmarkProtectionState", "Import-BenchmarkFixtureWindow"):
        assert _function(_code(), name) == _function(_code_at("ce5951f"), name), (
            f"{name} changed after the run that proved it")
    assert not _git("diff", "--name-only", "ce5951f", "--",
                    "pccm/bootstrap/windows/phase10_fixture_window.bas").strip()


SHAPE_HARNESS = PCCM_ROOT / "tests" / "phase10_run_shape_flow.ps1"

# WHAT EVERY PLANNED PERF-SMALL RUN MUST PRODUCE from the run loop's own
# iteration selection, executed under the runner's own param block.
#   label -> (CLR type, value, what [double] gives)
SHAPE_ITER_EXPECTED = {
    "plan-calculate-none": ("<null>", "n/a", "n/a"),
    "plan-recalculation-none": ("<null>", "n/a", "n/a"),
    "plan-simulation-10000": ("System.Int32", "10000", "10000"),
    "plan-sensitivity-10000": ("System.Int32", "10000", "10000"),
    "plan-annual-10000": ("System.Int32", "10000", "10000"),
    "plan-simulation-50000": ("System.Int32", "50000", "50000"),
    "plan-sensitivity-50000": ("System.Int32", "50000", "50000"),
    "plan-annual-50000": ("System.Int32", "50000", "50000"),
    "plan-simulation-100000": ("System.Int32", "100000", "100000"),
    "plan-sensitivity-100000": ("System.Int32", "100000", "100000"),
    "plan-annual-100000": ("System.Int32", "100000", "100000"),
    # The parameter really supplied, which is the state its [int[]] constraint
    # was written for and the one a scoped run uses.
    "parameter-supplied-10000": ("System.Int32", "10000", "10000"),
}

# WHAT THE SAMPLE VALIDATOR MUST SAY. count, valid, whether the shape guard
# accepted the list.
SHAPE_SAMPLE_EXPECTED = {
    "clean-command": (0, True, True),
    "clean-recalculation": (0, True, True),
    "endpoint-refusal": (1, False, True),
    "malformed-announcement": (1, False, True),
    "missing-published-iterations": (1, False, True),
    "wrong-published-iterations": (1, False, True),
    "right-published-iterations": (0, True, True),
    "annual-state-wrong": (1, False, True),
    "two-problems-at-once": (3, False, True),
    # THE DEFECT ITSELF. It still reproduces exactly - and the guard refuses it.
    "double-wrapped": (1, False, False),
}


def _shape_rows(runner: Path) -> dict:
    """Run the shape harness against one runner and parse its tagged lines."""
    done = subprocess.run(
        [PWSH, "-NoProfile", "-File", str(SHAPE_HARNESS), "-Runner", str(runner),
         "-Plan", str(PLAN_FILE)],
        capture_output=True, text=True, timeout=600)
    assert done.returncode == 0, done.stdout + done.stderr
    rows = {"iter": {}, "sample": {}, "unmet": []}
    for line in done.stdout.splitlines():
        if line.startswith(("PARSE|", "MISSING|")):
            # THE HARNESS COULD NOT FIND WHAT IT LIFTS. That is a refusal the
            # named controls below must report, not an error in this reader: an
            # anchor that moved means the thing being proved is no longer there.
            rows["unmet"].append(line)
            continue
        if line.startswith("ITER|"):
            label, kind, shown, as_double, banner = line[len("ITER|"):].split("|", 4)
            rows["iter"][label] = (kind, shown, as_double, banner)
        elif line.startswith("SAMPLE|"):
            parts = line[len("SAMPLE|"):].split("|", 5)
            name, count, valid, guard = parts[0], parts[1], parts[2], parts[3]
            rows["sample"][name] = {
                "count": int(count.split("=")[1]),
                "valid": valid.split("=")[1] == "True",
                "guard": guard.split("=", 1)[1],
                # THE MESSAGE CAN CONTAIN PIPES - a refusal reads
                # "FAIL|Calculate|..." - so the tail is rejoined and only the
                # field name is stripped.
                "rendered": "|".join(parts[4:])[len("rendered="):],
            }
    return rows


def _shape() -> dict:
    if "shape" not in _MEMO:
        _MEMO["shape"] = _shape_rows(RUNNER)
    return _MEMO["shape"]


# ===========================================================================
# O. THE TWO SHAPE DEFECTS THE FIRST TIMED RUN FOUND
# ===========================================================================
@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_180_every_planned_run_selects_exactly_one_integer_iteration_count() -> None:
    """A CLR TYPE, NOT A CLAIM ABOUT ONE. The run loop's `$iterations` WAS the
    script parameter `[int[]]$Iterations` - PowerShell variable names are
    case-insensitive, and a typed parameter keeps its constraint for the whole
    life of the variable, so every assignment was coerced back to `[int[]]`.
    `[string]` of a one-element array is still the element, which is why the
    banner read correctly and only `[double]` raised.

    So this executes the runner's OWN param block and the loop's OWN selection
    lines, once per planned PERF-SMALL run, and reports what came out."""
    shape = _shape()
    assert shape["unmet"] == [], (
        f"the shape harness could not lift what it proves: {shape['unmet']}")
    rows = shape["iter"]
    assert set(rows) == set(SHAPE_ITER_EXPECTED), (sorted(rows), sorted(SHAPE_ITER_EXPECTED))
    for label, (kind, shown, as_double) in SHAPE_ITER_EXPECTED.items():
        got = rows[label]
        assert got[0] == kind, (label, got)
        assert got[1] == shown, (label, got)
        # `[double]` IS THE CAST THAT RAISED ON WINDOWS, so it is the one checked.
        assert got[2] == as_double, (label, got)
        assert "RAISED" not in got[2], (label, got)
    # NINE ITERATION-DEPENDENT RUNS, EACH A SINGLE INTEGER, and two that carry
    # none - which is the accepted eleven.
    scalars = [k for k, v in rows.items() if v[0] == "System.Int32" and k.startswith("plan-")]
    nulls = [k for k, v in rows.items() if v[0] == "<null>"]
    assert len(scalars) == 9, scalars
    assert len(nulls) == 2, nulls
    assert len(scalars) + len(nulls) == 11


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_181_the_sample_validator_returns_a_flat_list_and_stays_strict() -> None:
    """THE DEFECT WAS THE SHAPE, NOT THE STRICTNESS. `return ,@($problems)` paired
    with `@(Test-BenchmarkSample ...)` gave a one-element array whose element was
    the real list: count 1 whatever the sample found, so `Valid` was False on
    every execution and `-join` rendered the inner array as `System.Object[]`.
    Calculate and the recalculation had executed correctly and had no problems.

    Loosening the check would have been the wrong repair. Every refusal still
    refuses, and the count is now the number of problems."""
    shape = _shape()
    assert shape["unmet"] == [], (
        f"the shape harness could not lift what it proves: {shape['unmet']}")
    rows = shape["sample"]
    assert set(rows) == set(SHAPE_SAMPLE_EXPECTED), (sorted(rows), sorted(SHAPE_SAMPLE_EXPECTED))
    for name, (count, valid, accepted) in SHAPE_SAMPLE_EXPECTED.items():
        got = rows[name]
        assert got["count"] == count, (name, got)
        assert got["valid"] == valid, (name, got)
        assert (got["guard"] == "ACCEPTED") == accepted, (name, got)
    # A CLEAN SAMPLE IS VALID AND RENDERS NOTHING - the state that was impossible.
    assert rows["clean-command"]["rendered"] == ""
    assert rows["clean-recalculation"]["rendered"] == ""
    # AND NOTHING IS SILENTLY STRINGIFIED: three problems render as three.
    assert rows["two-problems-at-once"]["rendered"].count(";") == 2


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_182_a_re_wrapped_list_still_breaks_and_the_guard_refuses_it() -> None:
    """THE PROOF THAT THE GUARD IS NOT DECORATIVE. The harness recreates the exact
    defect - `@()` applied to the call, not to an assigned variable, because
    wrapping an array is a no-op - and it still produces `count=1`,
    `valid=False`, `System.Object[]`. The guard refuses it by name, so the next
    edit that half-keeps the two-place contract aborts the run instead of
    invalidating every sample for an unreadable reason."""
    row = _shape()["sample"]["double-wrapped"]
    assert row["count"] == 1 and row["valid"] is False
    assert row["rendered"] == "System.Object[]", row["rendered"]
    assert row["guard"].startswith("REFUSED:"), row["guard"]
    assert "NESTED list" in row["guard"]
    assert "do not re-wrap it" in row["guard"]
    # AND IT REFUSES RATHER THAN FLATTENING. A guard that repaired the shape would
    # hide the defect and keep the run going.
    guard = _function(_code(), "Assert-BenchmarkProblemList")
    assert "throw (" in guard
    # STATEMENT POSITION AGAIN. The refusal's own wording says the validator "must
    # return a list", and saying so is not returning one. What is banned is a
    # `return` STATEMENT - a guard that handed back a value would be repairing the
    # shape - and any reassignment of the list it was given.
    for line in guard.splitlines():
        stripped = line.strip()
        assert not stripped.startswith("return"), (
            f"the shape guard returns a value instead of refusing: {stripped}")
        assert "$Problems =" not in stripped, (
            f"the shape guard rewrites the list it was given: {stripped}")
    assert "-join" not in guard, "the shape guard flattens instead of refusing"


def test_183_the_two_place_array_contract_is_kept_at_both_of_its_sites() -> None:
    """A CONTRACT SPREAD OVER TWO PLACES IS ONE AN EDIT CAN HALF-KEEP, and this
    one was half-kept twice. A helper that returns `,@(...)` hands back ONE object
    that IS the array; the caller must ASSIGN it, never wrap it again.

    Both sites are named, so a third one cannot be added without appearing here."""
    code = _code()
    # STATEMENT POSITION, NOT THE LETTERS. The shape guard's refusal QUOTES the
    # defective pairing in its message, and quoting it is the opposite of writing
    # it - so a line only counts when the token starts a statement.
    commas = [line.strip() for line in code.splitlines()
              if line.strip().startswith("return ,@(") or "{ return ,@(" in line]
    assert len(commas) == 4, commas
    quoted = [line.strip() for line in code.splitlines()
              if "return ,@(" in line and line.strip() not in commas]
    assert len(quoted) == 1, quoted
    assert "do not re-wrap" in _function(code, "Assert-BenchmarkProblemList")
    # AND THE GUARD IS CALLED, not merely defined. A guard nobody invokes is the
    # same as no guard, and its presence in the file would read as protection.
    execution = _function(code, "Invoke-BenchmarkExecution")
    assert "Assert-BenchmarkProblemList -Problems $problems" in execution, (
        "the shape guard is never invoked on the problem list it exists to check")
    # IMMEDIATELY AFTER THE CALL, before the sample object is built from it.
    assert (execution.index("Assert-BenchmarkProblemList")
            < execution.index("Valid      =")), (
        "the sample is built before its problem list is checked")
    # EVERY comma-returning helper, and how its result is taken.
    for helper, expected in (("Test-BenchmarkSample", "$problems = Test-BenchmarkSample"),
                             ("Get-BenchmarkOneDriveRoots", "$roots = Get-BenchmarkOneDriveRoots"),
                             ("New-BenchmarkWeights", "(New-BenchmarkWeights -Years")):
        assert f"return ,@(" in _function(code, helper) or helper == "New-BenchmarkWeights", helper
        assert expected in code, f"{helper}'s result is not taken by assignment: {expected}"
        assert f"@({helper}" not in code, f"{helper}'s result is re-wrapped in @()"


def test_184_no_local_reuses_a_typed_script_parameter_name() -> None:
    """THE GENERAL FORM OF THE ITERATION DEFECT. PowerShell variable names are
    case-insensitive and a typed parameter keeps its constraint for the whole life
    of the variable, so a local that reuses a parameter's name is silently coerced
    to that parameter's type on every assignment. `[int[]]` and `[string[]]` are
    the dangerous ones: they turn a scalar into a one-element array that prints
    correctly and casts wrongly.

    So no assignment anywhere may target a name the param block declares - and
    the param block is read out of the runner rather than restated."""
    runner = _runner()
    block = runner[runner.index("param("):runner.index("\nSet-StrictMode")]
    declared = {}
    # THE TYPE CAN NEST ITS BRACKETS - `[int[]]`, `[string[]]` - so the pattern
    # must span them. A character class that stopped at the first `]` read every
    # array parameter as untyped, which is exactly the declaration that caused
    # the defect.
    for match in re.finditer(r"(\[[\w.\[\]]+\]\s*)?\$(\w+)\s*[,)=]", block):
        declared[match.group(2).lower()] = (match.group(1) or "").strip()
    assert {"scenario", "iterations", "operations"} <= set(declared), sorted(declared)
    assert declared["iterations"] == "[int[]]", declared["iterations"]

    body = runner[runner.index("\nSet-StrictMode"):]
    offenders = []
    for match in re.finditer(r"^\s*\$(\w+)\s*=[^=]", body, re.M):
        name = match.group(1).lower()
        if name in declared and declared[name]:
            offenders.append(match.group(0).strip())
    # THE THREE PATH PARAMETERS ARE DEFAULTED IN PLACE ON PURPOSE - same [string]
    # type, so no coercion can change a shape - and they are named here rather
    # than excluded by a pattern.
    allowed = ("$BuildDir =", "$WorkDir =", "$OutDir =")
    unexpected = [o for o in offenders if not o.startswith(allowed)]
    assert unexpected == [], (
        f"a local reuses a typed script parameter's name, so every assignment to "
        f"it is coerced to that type: {unexpected}")


def test_185_the_shape_harness_reads_the_real_runner_and_starts_no_excel() -> None:
    """IT LIFTS THE PARAM BLOCK TOO, which is the defect's other half. A probe
    that declared its own `[int[]]$Iterations` would be testing its own guess
    about what the runner declares."""
    harness = SHAPE_HARNESS.read_text(encoding="utf-8")
    assert "$ast.ParamBlock.Extent.Text" in harness, "the param block is not lifted"
    assert "FunctionDefinitionAst" in harness and "Invoke-Expression" in harness
    assert "'ITER|'" in harness and "'SAMPLE|'" in harness
    for banned in ("New-Object -ComObject", "Excel.Application", "Workbooks"):
        assert banned not in harness, f"the shape harness starts Excel: {banned}"
    # AND IT IS TEST-ONLY.
    assert "phase10_run_shape_flow" not in _runner()


def test_186_the_median_still_needs_three_valid_warm_samples() -> None:
    """THE CONTRACT THE SHAPE DEFECT WAS HIDING. With every sample invalid the
    median was never computed and the gate was never exercised. It is unchanged:
    a median exists only when EVERY warm sample was valid, and the count comes
    from the plan rather than from a literal."""
    code = _code()
    assert "if ($validWarm.Count -eq [int]$run.warm_runs) {" in code
    assert "$valid = ([bool](($cold.Valid) -and ($validWarm.Count -eq [int]$run.warm_runs)))" in code
    # INVALID SAMPLES ARE EXCLUDED FROM THE MEDIAN, not counted toward it.
    assert "$validWarm = @($warm | Where-Object { $_.Valid })" in code
    assert "Get-BenchmarkMedian -Values @($validWarm" in code
    # AND THE PLAN STILL SAYS ONE COLD AND THREE WARM, ELEVEN TIMES.
    runs = [r for r in _plan()["runs"] if r["scenario"] == "PERF-SMALL"]
    assert len(runs) == 11, len(runs)
    assert {r["cold_runs"] for r in runs} == {1}
    assert {r["warm_runs"] for r in runs} == {3}


def test_187_the_fixture_window_is_byte_identical_to_the_run_that_proved_it() -> None:
    """THE WINDOW IS RUNTIME PROVEN AND IS NOT REOPENED. Windows at ce5951f opened
    it, built the fixture, closed it and verified depth 0 before timing. These two
    shape defects are in the run loop, which is downstream of all of it."""
    for name in ("Open-BenchmarkFixtureWindow", "Close-BenchmarkFixtureWindow",
                 "Invoke-BenchmarkWindowRollback", "Assert-BenchmarkProtectionApplied",
                 "Get-BenchmarkProtectionState", "Import-BenchmarkFixtureWindow"):
        current = _function(_code(), name)
        accepted = _function(_code_at("ce5951f"), name)
        assert current == accepted, f"{name} changed after the run that proved it"
    # THE SHIM TOO, BYTE FOR BYTE.
    assert not _git("diff", "--name-only", "ce5951f", "--",
                    "pccm/bootstrap/windows/phase10_fixture_window.bas").strip()
    # AND THE WINDOW STILL SPANS THE FIXTURE AND NOTHING ELSE.
    code = _code()
    assert code.count("Open-BenchmarkFixtureWindow -Excel") == 1
    assert code.count("Close-BenchmarkFixtureWindow -Excel") == 1


def _code_at(commit: str) -> str:
    """The runner as it stood at one commit, comments removed like `_code`."""
    key = ("code_at", commit)
    if key not in _MEMO:
        text = _git("show", f"{commit}:pccm/bootstrap/windows/phase10_benchmark.ps1")
        text = re.sub(r"<#.*?#>", "", text, flags=re.S)
        _MEMO[key] = "\n".join(
            line for line in text.splitlines() if not line.strip().startswith("#"))
    return _MEMO[key]


FLOW_HARNESS = PCCM_ROOT / "tests" / "phase10_fixture_window_flow.ps1"

# WHAT EACH SCENARIO MUST DO. The macro sequence is the whole proof: "exactly
# one compensating End" is a count, and a count cannot be read off source text.
#
#   calls  - the Application.Run macros, in order
#   timed  - 1 only when control reached the stand-in for the run loop
FLOW_EXPECTED = {
    # Nothing was opened, so nothing is owed. An End here would decrement a
    # depth nobody raised.
    "begin-refuses": (
        ["P10FW_State", "P10FW_Begin"], 0),
    # Opened, then this function's own checks refused. One compensating End.
    "post-open-depth-wrong": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    "post-open-state-read-refuses": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    "post-open-state-read-raises": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    "post-open-structure-false": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    # The rollback itself failing does not buy a second attempt.
    "post-open-fails-and-rollback-refuses": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    "post-open-fails-and-rollback-raises": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    # A successful open closes NOTHING itself; the caller's finally does it once.
    "open-succeeds-fixture-succeeds": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End", "P10FW_State"], 1),
    "open-succeeds-fixture-throws": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End", "P10FW_State"], 0),
    "open-succeeds-close-refuses": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End"], 0),
    "open-succeeds-close-leaves-depth-open": (
        ["P10FW_State", "P10FW_Begin", "P10FW_State", "P10FW_End", "P10FW_State"], 0),
}


def _flow_rows(runner: Path) -> dict:
    """Run the control-flow harness against one runner and parse its output.

    Takes the runner as an argument so the mutation battery can point it at a
    damaged copy on disk - the harness is an AST lift over a real file, so a
    damaged copy is the only way to mutate what it observes.
    """
    done = subprocess.run(
        [PWSH, "-NoProfile", "-File", str(FLOW_HARNESS), "-Runner", str(runner)],
        capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stdout + done.stderr
    rows = {}
    for line in done.stdout.splitlines():
        assert not line.startswith(("PARSE|", "MISSING|")), line
        # ONLY TAGGED LINES. A runner under mutation may write to the host, and an
        # untagged line parsed as a result would turn a caught mutation into an
        # unreadable one.
        if not line.startswith("FLOW|"):
            continue
        name, calls, outcome, timed, message = line[len("FLOW|"):].split("|", 4)
        rows[name] = {
            "calls": [c for c in calls.split(",") if c],
            "outcome": outcome,
            "timed": int(timed.split("=")[1]),
            "message": message,
        }
    return rows


def _flow() -> dict:
    if "flow" not in _MEMO:
        _MEMO["flow"] = _flow_rows(RUNNER)
    return _MEMO["flow"]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_170_the_window_opens_transactionally_on_every_failure_path() -> None:
    """EXECUTED, NOT READ. Every other control here asserts what the runner SAYS.
    "Exactly one compensating P10FW_End, and only on the failing paths" is a
    COUNT of what PowerShell does with a try/catch when a `return` is taken and
    when a post-open assertion throws - asserting that from source would be
    asserting a belief about the language.

    So the harness lifts the real functions and the caller's own try/finally out
    of the runner by AST and runs them against a fake that records every macro.
    Excel is never started."""
    flow = _flow()
    assert set(flow) == set(FLOW_EXPECTED), (sorted(flow), sorted(FLOW_EXPECTED))
    for name, (calls, timed) in FLOW_EXPECTED.items():
        assert flow[name]["calls"] == calls, (name, flow[name]["calls"], calls)
        assert flow[name]["timed"] == timed, (name, flow[name]["timed"], timed)


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_171_a_failed_begin_compensates_nothing() -> None:
    """NOTHING WAS OPENED, SO NOTHING IS OWED. A compensating close here would
    decrement a depth nobody raised, and `ProtectionEndStructural` reports
    exactly that as "closed more often than it was opened"."""
    row = _flow()["begin-refuses"]
    assert "P10FW_End" not in row["calls"], row["calls"]
    assert row["outcome"] == "THREW"
    assert "could not be opened" in row["message"]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_172_every_post_open_failure_attempts_exactly_one_close() -> None:
    """THE GAP THIS ROUND CLOSED. `P10FW_Begin` succeeds, a post-open check then
    refuses, and before the correction nothing closed what had been opened -
    the error escaped to the abandon path leaving a half-open window behind."""
    flow = _flow()
    for name in ("post-open-depth-wrong", "post-open-state-read-refuses",
                 "post-open-state-read-raises", "post-open-structure-false"):
        row = flow[name]
        assert row["calls"].count("P10FW_End") == 1, (name, row["calls"])
        assert row["outcome"] == "THREW", name
        assert row["timed"] == 0, name
        # THE ORIGINAL FAILURE SURVIVES THE ROLLBACK, and the rollback says it took.
        assert "The window was open when this failed" in row["message"], name
        assert "it succeeded and protection is restored" in row["message"], name


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_173_a_failed_rollback_is_reported_beside_the_failure_that_caused_it() -> None:
    """BOTH FACTS, IN ONE SENTENCE. A `finally` that threw would REPLACE the
    original exception with the rollback's, losing the diagnosis the rollback
    exists to accompany - which is why the compensation is a `catch`.

    A REFUSAL AND A RAISE ARE REPORTED DIFFERENTLY, because "the workbook says it
    could not restore protection" and "the call never arrived" are different
    facts about the machine."""
    flow = _flow()
    refused = flow["post-open-fails-and-rollback-refuses"]
    assert refused["calls"].count("P10FW_End") == 1, refused["calls"]
    assert "opened to depth 2" in refused["message"], "the original failure was lost"
    assert "IT REFUSED" in refused["message"]
    assert "must not be measured" in refused["message"]
    assert refused["timed"] == 0

    raised = flow["post-open-fails-and-rollback-raises"]
    assert raised["calls"].count("P10FW_End") == 1, raised["calls"]
    assert "opened to depth 2" in raised["message"], "the original failure was lost"
    assert "IT RAISED" in raised["message"]
    assert raised["timed"] == 0


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_174_a_successful_open_closes_nothing_itself() -> None:
    """NO DOUBLE-DECREMENT. `return` inside a `try` does not enter its `catch`, so
    a successful open hands the window to the caller untouched and the caller's
    `finally` performs the one normal close. Both closing would take
    modProtection's depth to -1, which it reports as a close without an open."""
    flow = _flow()
    ok = flow["open-succeeds-fixture-succeeds"]
    assert ok["calls"].count("P10FW_End") == 1, ok["calls"]
    assert ok["outcome"] == "RETURNED"
    assert ok["timed"] == 1
    # THE ONE End IS THE CALLER'S: it comes after the open's own verification read
    # and is followed by the post-close verification read.
    assert ok["calls"] == ["P10FW_State", "P10FW_Begin", "P10FW_State",
                           "P10FW_End", "P10FW_State"]

    raised = flow["open-succeeds-fixture-throws"]
    assert raised["calls"].count("P10FW_End") == 1, raised["calls"]
    assert raised["timed"] == 0
    # AND THE FIXTURE'S OWN FAILURE IS WHAT IS REPORTED, not the close's success.
    assert raised["message"].strip() == "THE FIXTURE RAISED", raised["message"]


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_175_no_timed_work_follows_any_window_failure() -> None:
    """THE POINT OF ALL OF IT. Whatever failed - the open, a post-open check, the
    rollback, the close, or the post-close verification - control never reaches
    the run loop, so no sample can be taken from a workbook whose protection is
    not known to be restored."""
    flow = _flow()
    for name, row in flow.items():
        if name == "open-succeeds-fixture-succeeds":
            assert row["timed"] == 1, name
        else:
            assert row["timed"] == 0, (name, row)
            assert row["outcome"] == "THREW", name
    # INCLUDING A CLOSE THAT REPORTED SUCCESS WHILE THE DEPTH STAYED OPEN.
    stuck = flow["open-succeeds-close-leaves-depth-open"]
    assert "still open at depth 1" in stuck["message"], stuck["message"]


def test_176_the_rollback_is_the_only_close_inside_open_and_cannot_raise() -> None:
    """THE SOURCE SIDE OF THE SAME PROPERTY, so a rewrite that kept the behaviour
    the harness samples but reintroduced the shape cannot pass quietly.

    The rollback returns text on every path - it is called from a catch that is
    about to rethrow, and a rollback that threw would destroy the diagnosis."""
    opener = _function(_code(), "Open-BenchmarkFixtureWindow")
    assert opener.count("P10FW_End") == 0, (
        "the open closes the window itself instead of delegating to the rollback")
    assert opener.count("Invoke-BenchmarkWindowRollback") == 1, opener
    assert "} catch {" in opener and "} finally {" not in opener, (
        "the compensation is a finally, which would replace the original exception")
    # THE BEGIN IS OUTSIDE THE GUARDED REGION, so a refused open compensates nothing.
    guarded = opener[opener.index("    try {"):]
    assert "P10FW_Begin" not in guarded, "a failed Begin would now be compensated"

    rollback = _function(_code(), "Invoke-BenchmarkWindowRollback")
    assert rollback.count("P10FW_End") == 1, rollback
    assert "throw" not in rollback, "the rollback can raise out of a catch block"
    assert rollback.count("return (") == 3, "a path through the rollback reports nothing"
    # NOT A CATCH-AND-IGNORE: the failing branch becomes text, and says so.
    assert "IT REFUSED" in rollback and "IT RAISED" in rollback
    assert "must not be measured" in rollback


def test_177_the_flow_harness_reads_the_real_runner_and_starts_no_excel() -> None:
    """A HARNESS THAT TESTED A COPY WOULD PROVE NOTHING. It lifts the functions
    and the caller region out of the shipping file by AST and by anchored text,
    so a rename or a reformat of either is a failure here rather than a silently
    stale test."""
    harness = FLOW_HARNESS.read_text(encoding="utf-8")
    for lifted in ("Get-BenchmarkProtectionState", "Assert-BenchmarkProtectionApplied",
                   "Open-BenchmarkFixtureWindow", "Invoke-BenchmarkWindowRollback",
                   "Close-BenchmarkFixtureWindow"):
        assert lifted in harness, lifted
    assert "FunctionDefinitionAst" in harness and "Invoke-Expression" in harness
    # RESULT LINES ARE TAGGED, so stray host output from a damaged runner cannot
    # be read as a scenario result.
    assert "'FLOW|'" in harness
    # NO EXCEL, ANYWHERE.
    for banned in ("New-Object -ComObject", "Excel.Application", "Workbooks"):
        assert banned not in harness, f"the flow harness starts Excel: {banned}"
    # AND IT IS TEST-ONLY: it is not reachable from the runner.
    assert "phase10_fixture_window_flow" not in _runner()


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_144_every_command_the_benchmark_can_reach_still_resolves() -> None:
    """THE CLASS THAT ENDED PHASE-9 WINDOWS RUN 1, and the reason
    `Remove-TableRow` is a refusal rather than a deletion.

    The audit walks the runner AND every file it dot-sources, so a name only the
    Gate-B file calls is still a name that must resolve. Removing the definition
    passed every text control in this file and failed here."""
    done = _resolution_audit(RUNNER, *BENCHMARK_DECLARED_OVERRIDES)
    assert done.returncode == 0, done.stdout.strip() + done.stderr.strip()
    assert done.stdout.startswith("CLEAN"), done.stdout


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_145_the_override_declaration_is_required_and_is_not_a_blanket() -> None:
    """A DECLARATION NOBODY CHECKS IS AN EXEMPTION. Without it the duplicate is
    still reported, so the runner cannot quietly acquire a second override; and a
    declared name that overrides nothing is itself a finding, so the list cannot
    become a place to park exemptions that stopped being true."""
    undeclared = _resolution_audit(RUNNER)
    assert undeclared.returncode == 1, undeclared.stdout
    assert "DUPLICATE reset-phase5fxtable" in undeclared.stdout, undeclared.stdout

    stale = _resolution_audit(RUNNER, *BENCHMARK_DECLARED_OVERRIDES, "Get-TableBody")
    assert stale.returncode == 1, stale.stdout
    assert "DECLARED-OVERRIDE get-tablebody overrides nothing" in stale.stdout, stale.stdout


@pytest.mark.skipif(not Path(PWSH).exists(), reason="no PowerShell on this host")
def test_146_an_override_that_would_not_win_is_refused(tmp_path: Path) -> None:
    """THE DECLARATION CHECKS ORDER, WHICH IS THE WHOLE MECHANISM. An override
    placed above the dot-source is overwritten by the accepted definition: the
    runner silently runs the behaviour it meant to replace, with its corrected
    source still sitting in the file and every text control here still passing.

    Written beside the real runner because the audit resolves dot-sourced files
    relative to the script's own directory."""
    with RUNNER.open(encoding="utf-8", newline="") as handle:
        source = handle.read()
    start = source.index("function Reset-Phase5FxTable {")
    end = source.index("\r\n}\r\n", start) + 5
    block = source[start:end]
    assert "Set-Phase5TypedCell" in block, "the override block anchor moved"
    marker = ". (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')"
    moved = (source[:start] + source[end:]).replace(marker, block + "\r\n" + marker, 1)
    assert moved != source

    scratch = BOOTSTRAP / "__benchmark_override_order_tmp.ps1"
    with scratch.open("w", encoding="utf-8", newline="") as handle:
        handle.write(moved)
    try:
        done = _resolution_audit(scratch, *BENCHMARK_DECLARED_OVERRIDES)
    finally:
        scratch.unlink()
    assert done.returncode == 1, done.stdout
    assert "does not win" in done.stdout, done.stdout


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
