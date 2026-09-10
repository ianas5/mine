#!/usr/bin/env python3
"""P10-4A MUTATION CONTROLS.

A HARNESS CONTROL THAT CANNOT FAIL IS THE WORST KIND HERE, because nobody will
find out on Linux. The benchmark runs on a machine this repository never sees,
for hours, and the numbers it produces become the release baseline. If a control
over it is vacuous, the first evidence of that is a baseline nobody can defend.

So every control below breaks ONE thing - a dimension, an iteration count, a
gate, a policy sentence, a shutdown step - reruns the WHOLE conformance battery
against the damaged copy, and requires a NAMED control among the refusers.

WHAT IS MUTATED, AND WHY THESE

  the matrix        sizes redefined, the split changed, Large let up to 100,000,
                    Calculate multiplied across the iteration counts. Each is a
                    plausible "tidy-up" that would silently change what the
                    baseline means.

  the timing        four warm runs, the mean instead of the median, the median
                    computed from an invalid sample, the fixture build brought
                    inside the clock, and the clock widened to include the
                    automation envelope.

  the gates         a refusal counted as a sample, the iteration count not
                    verified, the annual state not required.

  the policy        an absolute pass mark invented, the ratios altered, the
                    Phase-7 numbers promoted from context to threshold.

  the lifecycle     the workbook saved, Excel killed instead of quit, the
                    shutdown ledger dropped from the artifact.

Nothing here writes to the repository: damaged copies live in memory and the
conformance module's cache is pointed at them for exactly one control.

Runs standalone or under pytest.
"""
from __future__ import annotations

import copy
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

import pytest  # noqa: E402

import test_phase10_benchmark_harness as conformance  # noqa: E402


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance) if name.startswith("test_"))
    assert len(names) >= 50, names
    return names


def _run_battery() -> list[str]:
    refused: list[str] = []
    for name in _tests():
        try:
            getattr(conformance, name)()
        except BaseException:  # noqa: BLE001 - any refusal counts
            refused.append(name)
    return refused


def _install(damaged: dict):
    saved = dict(conformance._MEMO)
    conformance._MEMO.update(damaged)

    def restore() -> None:
        conformance._MEMO.clear()
        conformance._MEMO.update(saved)

    return restore


def _detects(expected: str, **damaged) -> None:
    restore = _install(damaged)
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _plan_mutation(expected: str, mutate) -> None:
    plan = copy.deepcopy(conformance._plan())
    mutate(plan)
    _detects(expected, plan=plan)


def _runner_mutation(expected: str, before: str, after: str) -> None:
    """Damage the runner's SOURCE, and prove the anchor still matched.

    NOT AN AssertionError WHEN THE ANCHOR MISSES. A no-op mutation that raised
    one would be indistinguishable from the control it is meant to provoke, and
    the whole battery would rot silently the first time the runner was
    reformatted.
    """
    original = conformance._runner()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the runner")
    # `_code` is derived from `_runner`, so the derived entry must go too.
    restore = _install({"runner": damaged})
    try:
        conformance._MEMO.pop("code", None)
        refused = _run_battery()
    finally:
        restore()
        conformance._MEMO.pop("code", None)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# A. THE MATRIX
# ===========================================================================
def test_01_redefining_a_model_size_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["scenarios"][2]["drivers"] = 250

    _plan_mutation("test_01", mutate)


def test_02_changing_the_project_years_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["scenarios"][1]["years"] = 10

    _plan_mutation("test_01", mutate)


def test_03_changing_the_cost_risk_split_is_rejected() -> None:
    """THE SPLIT IS THE ACCEPTED PHASE-7 ONE. A benchmark built on a different
    one would not be comparable with the historical context beside it."""
    def mutate(plan: dict) -> None:
        plan["scenarios"][0]["cost_lines"] = 10
        plan["scenarios"][0]["risks"] = 10

    _plan_mutation("test_03", mutate)


def test_04_a_split_that_does_not_add_up_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["scenarios"][1]["risks"] = 30

    _plan_mutation("test_03", mutate)


def test_05_adding_an_iteration_count_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["scenarios"][0]["iterations"].append(250_000)

    _plan_mutation("test_04", mutate)


def test_06_letting_large_reach_one_hundred_thousand_is_rejected() -> None:
    """THE CAP THE CONTRACT SET. This is the mutation the whole of control 3
    exists for, and it is caught in the matrix rather than at run time."""
    def mutate(plan: dict) -> None:
        plan["scenarios"][2]["iterations"].append(100_000)
        plan["runs"].append({"scenario": "PERF-LARGE", "operation": "simulation",
                             "iterations": 100_000, "cold_runs": 1, "warm_runs": 3})

    _plan_mutation("test_05", mutate)


def test_07_quietly_dropping_the_forbidden_record_is_rejected() -> None:
    """NOT RUNNING IT AND RECORDING WHY ARE DIFFERENT THINGS. Without the
    record, the missing cell reads as missing evidence."""
    def mutate(plan: dict) -> None:
        plan["forbidden"] = []

    _plan_mutation("test_05", mutate)


def test_08_a_forbidden_reason_that_stops_naming_the_cap_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["forbidden"][0]["reason"] = "we did not run it"

    _plan_mutation("test_05", mutate)


def test_09_repeating_calculate_per_iteration_count_is_rejected() -> None:
    """THREE MEASUREMENTS OF IDENTICAL WORK would make the matrix look richer
    than the evidence is."""
    def mutate(plan: dict) -> None:
        plan["runs"].append({"scenario": "PERF-SMALL", "operation": "calculate",
                             "iterations": 50_000, "cold_runs": 1, "warm_runs": 3})

    _plan_mutation("test_06", mutate)


def test_10_making_calculate_iteration_dependent_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        for operation in plan["operations"]:
            if operation["key"] == "calculate":
                operation["iteration_dependent"] = True

    _plan_mutation("test_06", mutate)


def test_11_dropping_a_stochastic_operation_from_a_size_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["runs"] = [run for run in plan["runs"]
                        if not (run["scenario"] == "PERF-MEDIUM"
                                and run["operation"] == "annual"
                                and run["iterations"] == 100_000)]

    _plan_mutation("test_07", mutate)


# ===========================================================================
# B. THE TIMING METHOD
# ===========================================================================
def test_20_a_fourth_warm_run_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["timing"]["warm_runs"] = 4
        for run in plan["runs"]:
            run["warm_runs"] = 4

    _plan_mutation("test_10", mutate)


def test_21_dropping_the_cold_run_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["timing"]["cold_runs"] = 0
        for run in plan["runs"]:
            run["cold_runs"] = 0

    _plan_mutation("test_10", mutate)


def test_22_making_the_mean_the_comparison_statistic_is_rejected() -> None:
    """THE MEAN WOULD AVERAGE AWAY THE OUTLIER, and the outlier is information
    about the machine."""
    def mutate(plan: dict) -> None:
        plan["timing"]["comparison_statistic"] = "mean_of_warm_runs"

    _plan_mutation("test_11", mutate)


def test_23_a_runner_that_averages_instead_of_taking_the_median_is_rejected() -> None:
    _runner_mutation(
        "test_11",
        "    $sorted = @($Values | Sort-Object)",
        "    $sorted = @($Values | Measure-Object -Average)")


def test_24_computing_a_median_from_fewer_than_three_valid_samples_is_rejected() -> None:
    _runner_mutation(
        "test_12",
        "        if ($validWarm.Count -eq [int]$run.warm_runs) {",
        "        if ($validWarm.Count -ge 1) {")


def test_25_bringing_the_fixture_build_inside_the_clock_is_rejected() -> None:
    """SETUP INSIDE A MEASUREMENT is the defect that would make every number in
    the report wrong in the same direction."""
    _runner_mutation(
        "test_15",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        $Excel.Run('PCCM_AutomationBegin', $true, '') | Out-Null\n"
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null")


def test_26_reading_the_evidence_inside_the_clock_is_rejected() -> None:
    _runner_mutation(
        "test_15",
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null\n"
        "        $watch.Stop()",
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null\n"
        "        $result = [string]$Excel.Run('PCCM_AutomationResult')\n"
        "        $watch.Stop()")


def test_27_timing_with_a_wall_clock_is_rejected() -> None:
    _runner_mutation(
        "test_14",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        $Excel.CalculateFull()",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        $started = Get-Date\n"
        "        $Excel.CalculateFull()")


def test_28_dropping_the_setup_timings_from_the_artifact_is_rejected() -> None:
    _runner_mutation(
        "test_16",
        "$report.Add('setup_ms', $setupTimings)",
        "$report.Add('setup_ms', 'not recorded')")


def test_29_a_cold_definition_that_means_the_workbook_open_is_rejected() -> None:
    """COLD MUST NOT BE FAKED. The one thing it is most easily made to mean is
    the cost of starting Excel."""
    def mutate(plan: dict) -> None:
        plan["timing"]["cold_definition"] = "the time to open the workbook"
        plan["timing"]["excluded_from_elapsed"] = ["Stage-B bootstrap"]

    _plan_mutation("test_13", mutate)


# ===========================================================================
# C. THE ENVIRONMENT RECORD
# ===========================================================================
def test_30_dropping_the_excel_bitness_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["environment_fields"].remove("excel_bitness")

    _plan_mutation("test_20", mutate)


def test_31_dropping_the_cpu_or_the_memory_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["environment_fields"].remove("installed_ram_gb")

    _plan_mutation("test_21", mutate)


def test_32_dropping_the_sync_classification_is_rejected() -> None:
    """THE REPOSITORY IS UNDER ONEDRIVE ON THIS USER'S MACHINE, and a report
    that recorded only a temp path would let that be read as a local
    unsynchronised one."""
    def mutate(plan: dict) -> None:
        plan["environment_fields"].remove("repository_location_type")

    _plan_mutation("test_22", mutate)


def test_33_a_runner_that_stops_classifying_onedrive_is_rejected() -> None:
    _runner_mutation(
        "test_22",
        "            return 'onedrive'",
        "            return 'local'")


def test_34_dropping_the_commit_from_the_record_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["environment_fields"].remove("git_commit")

    _plan_mutation("test_23", mutate)


def test_35_a_declared_field_the_runner_never_populates_is_rejected() -> None:
    """A PROMISE IN ONE FILE AND SILENCE IN THE OTHER."""
    def mutate(plan: dict) -> None:
        plan["environment_fields"].append("gpu_model")

    _plan_mutation("test_24", mutate)


def test_36_a_runner_that_omits_a_declared_field_is_rejected() -> None:
    _runner_mutation(
        "test_24",
        "    $record.Add('cpu_model', $(if ($null -ne $cpu) { [string]$cpu.Name } else { $unknown }))",
        "")


def test_37_theorising_about_onedrive_performance_is_rejected() -> None:
    """RECORD IT, DO NOT EXPLAIN IT. The harness observes where the file is; a
    claim about what that does to a timing is not evidence."""
    _runner_mutation(
        "test_22",
        "# WHAT KIND OF PLACE THIS PATH IS. An observation about the path and nothing",
        "# WHAT KIND OF PLACE THIS PATH IS. A synced path is slower, so this is")


# ===========================================================================
# D. DETERMINISM
# ===========================================================================
def test_40_a_random_scenario_is_rejected() -> None:
    _runner_mutation(
        "test_30",
        "        $shape = [double](1 + ((($year + $Offset) % 5)))",
        "        $shape = [double](Get-Random -Minimum 1 -Maximum 5)")


def test_41_an_auto_seed_is_rejected() -> None:
    """AUTO SAMPLES A DIFFERENT SEQUENCE EVERY TIME, so a difference between two
    benchmark runs would be the sample rather than the machine."""
    _runner_mutation(
        "test_30",
        "function Get-BenchmarkSeed { return 20260101 }",
        "function Get-BenchmarkSeed { return $null }")


def test_42_weights_that_do_not_sum_to_one_are_rejected() -> None:
    """PRODUCTION WOULD REFUSE THEM, CORRECTLY, and the harness must never
    provoke that: a scenario that cannot be calculated measures nothing."""
    _runner_mutation(
        "test_32",
        "    $weights += [double](1.0 - $running)",
        "    $weights += [double]([Math]::Round([double]$raw[$Years - 1] / $total, 6))")


def test_43_a_degenerate_driver_population_is_rejected() -> None:
    """A ZERO-VARIANCE DRIVER IS NOT WHAT IS BEING TIMED. If every driver were
    degenerate the sensitivity measurement would be of the refusal path."""
    _runner_mutation(
        "test_33",
        "        most_likely       = [double]($base * 1.35)",
        "        most_likely       = $base")


def test_44_a_single_currency_scenario_is_rejected() -> None:
    """FX, INFLATION AND PROFILING ARE PART OF THE WORK. A model without them
    would be a cheap benchmark of something the user does not have."""
    _runner_mutation(
        "test_34",
        "            [pscustomobject]@{ currency = 'USD'; rate = 3.75 }",
        "            [pscustomobject]@{ currency = 'SAR'; rate = 1.0 }")


def test_45_hard_coding_a_dimension_in_the_runner_is_rejected() -> None:
    _runner_mutation(
        "test_31",
        "    $costCount = [int]$ScenarioSpec.cost_lines",
        "    $costCount = 12")


# ===========================================================================
# E. THE CORRECTNESS GATES
# ===========================================================================
def test_50_counting_a_refusal_as_a_sample_is_rejected() -> None:
    """A FAST FAILURE IS NOT A FAST OPERATION, and it is the failure mode that
    produces the most attractive number in the report."""
    _runner_mutation(
        "test_40",
        "        if ($result -notlike 'OK|*') {",
        "        if ($false) {")


def test_51_dropping_the_validity_filter_is_rejected() -> None:
    _runner_mutation(
        "test_40",
        "        $validWarm = @($warm | Where-Object { $_.Valid })",
        "        $validWarm = @($warm)")


def test_52_not_verifying_the_iteration_count_is_rejected() -> None:
    """A RUN THAT EXECUTED TEN THOUSAND ITERATIONS WHEN A HUNDRED THOUSAND WERE
    ASKED FOR would be the fastest and most useless measurement in the file."""
    _runner_mutation(
        "test_41",
        "        } elseif (([int]$published) -ne ([int]$RequestedIterations)) {",
        "        } elseif ($false) {")


def test_53_treating_an_unreported_iteration_count_as_acceptable_is_rejected() -> None:
    _runner_mutation(
        "test_41",
        "            $problems += 'the operation did not report how many iterations it executed'",
        "            $problems += @()")


def test_54_dropping_the_annual_state_gate_is_rejected() -> None:
    _runner_mutation(
        "test_44",
        "        if ($distribution -ne 'CURRENT') {",
        "        if ($false) {")


def test_55_a_gate_policy_that_admits_invalid_samples_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["correctness_gates"]["invalid_sample_policy"] = (
            "a refused execution is reported with its time")

    _plan_mutation("test_40", mutate)


def test_56_not_reading_the_dimensions_back_is_rejected() -> None:
    _runner_mutation(
        "test_45",
        "    if (([int]$actual['drivers']) -ne ([int]$scenarioSpec.drivers)) {",
        "    if ($false) {")


# ===========================================================================
# F. IT MEASURES PRODUCTION
# ===========================================================================
def test_60_benchmarking_something_that_is_not_a_user_command_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        for operation in plan["operations"]:
            if operation["key"] == "sensitivity":
                operation["endpoint"] = "PCCM_SensitivityFast"

    _plan_mutation("test_50", mutate)


def test_61_a_harness_that_implements_a_statistic_of_its_own_is_rejected() -> None:
    """A BENCHMARK-ONLY CODE PATH WOULD MEASURE THE HARNESS."""
    _runner_mutation(
        "test_52",
        "function Get-BenchmarkMedian {",
        "function Get-BenchmarkSpearman { param($A, $B) return [Math]::Sqrt(1) }\n"
        "function Get-BenchmarkMedian {")


def test_62_calling_the_endpoint_by_a_hard_coded_name_is_rejected() -> None:
    _runner_mutation(
        "test_51",
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null",
        "        $Excel.Run('PCCM_RunSimulation') | Out-Null")


def test_63_a_runner_that_declares_its_own_matrix_is_rejected() -> None:
    _runner_mutation(
        "test_81",
        "$declaredIterations = @($scenarioSpec.iterations | ForEach-Object { [int]$_ })",
        "$declaredIterations = @(10000, 50000, 100000)")


# ===========================================================================
# G. THE POLICY
# ===========================================================================
def test_70_inventing_an_absolute_pass_mark_is_rejected() -> None:
    """THERE IS NOTHING TO COMPARE AGAINST YET. A threshold invented now would
    make the first run pass or fail for a reason nobody could defend."""
    def mutate(plan: dict) -> None:
        plan["regression_policy"]["absolute_threshold_seconds"] = 120

    _plan_mutation("test_63", mutate)


def test_71_a_runner_that_judges_a_timing_is_rejected() -> None:
    _runner_mutation(
        "test_63",
        "$script:BenchmarkSupportedSchema = 1",
        "$script:BenchmarkSupportedSchema = 1\n$script:BudgetSeconds = 300")


def test_72_altering_the_investigate_ratio_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["regression_policy"]["investigate_ratio"] = 3.0

    _plan_mutation("test_64", mutate)


def test_73_altering_the_blocking_ratio_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["regression_policy"]["blocking_ratio"] = 5.0
        plan["regression_policy"]["rules"][2] = (
            "above 5.0x: blocking regression unless explained and explicitly accepted")

    _plan_mutation("test_64", mutate)


def test_74_making_a_ratio_apply_to_the_cold_run_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["regression_policy"]["applies_to"] = "cold against cold"

    _plan_mutation("test_64", mutate)


def test_75_promoting_the_phase_seven_timings_to_a_threshold_is_rejected() -> None:
    """THEY ARE THE OPERATOR'S REPORT OF A DIFFERENT MEASUREMENT, on a machine
    whose record was never settled, against code that has since changed."""
    def mutate(plan: dict) -> None:
        plan["historical_context"]["status"] = "the baseline to beat"

    _plan_mutation("test_65", mutate)


def test_76_a_runner_that_compares_against_a_historical_number_is_rejected() -> None:
    _runner_mutation(
        "test_65",
        "Write-BenchmarkLine 'SUMMARY'",
        "Write-BenchmarkLine 'SUMMARY'\n"
        "if ($sessionWatch.Elapsed.TotalSeconds -gt 105.4) { Write-BenchmarkLine 'slower than Phase 7' }")


def test_77_altering_the_baseline_identifier_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["baseline_id"] = "adhoc"

    _plan_mutation("test_66", mutate)


def test_78_an_unversioned_artifact_schema_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["schema_version"] = 2

    _plan_mutation("test_60", mutate)


def test_79_a_runner_that_accepts_any_plan_schema_is_rejected() -> None:
    _runner_mutation(
        "test_60",
        "if ([int]$plan.schema_version -ne $script:BenchmarkSupportedSchema) {",
        "if ($false) {")


# ===========================================================================
# H. LIFECYCLE
# ===========================================================================
def test_80_saving_the_workbook_is_rejected() -> None:
    """A CELL HOLDING A TIMING WOULD BE A STATE AUTHORITY NOBODY ASKED FOR."""
    _runner_mutation(
        "test_62",
        "            try { $wb.Close($false); $rel.WorkbookClosed = $true }",
        "            try { $wb.Save(); $wb.Close($false); $rel.WorkbookClosed = $true }")


def test_81_killing_excel_instead_of_quitting_it_is_rejected() -> None:
    _runner_mutation(
        "test_71",
        "            try { $excel.Quit(); $rel.QuitCalled = $true }",
        "            try { Stop-Process -Name EXCEL -Force; $rel.QuitCalled = $true }")


def test_82_dropping_the_wait_for_a_natural_exit_is_rejected() -> None:
    _runner_mutation(
        "test_70",
        "        $rel.NaturalExit = Wait-ExcelExit -Identity $excelIdentity",
        "        $rel.NaturalExit = $true")


def test_83_releasing_the_parent_before_the_leaf_is_rejected() -> None:
    _runner_mutation(
        "test_70",
        "        Invoke-NamedRelease $rel $wb        'Workbook';  $wb        = $null",
        "        $wb = $null")


def test_84_dropping_the_shutdown_record_from_the_artifact_is_rejected() -> None:
    """SO A RUN THAT LEFT A ZOMBIE SAYS SO IN THE EVIDENCE, rather than only in
    a console line nobody kept."""
    _runner_mutation(
        "test_72",
        "$report.Add('shutdown', $shutdownRecord)",
        "$report.Add('shutdown', 'clean')")


def test_85_opening_the_real_build_output_is_rejected() -> None:
    _runner_mutation(
        "test_73",
        "$stageBPath = Join-Path $tempRoot ([string]$manifest.stage_b_filename)",
        "$stageBPath = Join-Path $BuildDir ([string]$manifest.stage_b_filename)")


def test_86_running_without_attributing_the_workbook_is_rejected() -> None:
    _runner_mutation(
        "test_74",
        "    $revision = Get-BenchmarkSourceRevision -RepoRoot $repoRoot",
        "    $revision = [pscustomobject]@{ Head = 'unknown'; Dirty = @() }")


def test_87_letting_one_invocation_run_every_scenario_is_rejected() -> None:
    """THESE RUNS ARE LONG. One command at a time is what lets the operator stop
    after any of them and still hold complete evidence for the rest."""
    _runner_mutation(
        "test_75",
        "    [ValidateSet('PERF-SMALL', 'PERF-MEDIUM', 'PERF-LARGE')]",
        "    [ValidateSet('PERF-SMALL', 'PERF-MEDIUM', 'PERF-LARGE', 'All')]")


def test_88_drifting_from_the_accepted_com_primitives_is_rejected() -> None:
    """A DRIFTED COPY IS A SECOND, WORSE IMPLEMENTATION of the same COM access,
    and the drift would show up as a benchmark that behaved differently from
    every accepted scenario."""
    _runner_mutation(
        "test_57",
        "function Get-IdColumnValues {",
        "function Get-IdColumnValues {\n    # drifted\n")


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
