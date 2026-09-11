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


# CONTROLS THIS BATTERY DOES NOT RERUN, and exactly why each one.
#
# NOT AN EXEMPTION FROM BEING CHECKED. Each of these audits the runner ON DISK
# through a PowerShell subprocess, so a damaged copy held in memory is invisible
# to it - it would pass every mutation below while proving nothing, and it would
# cost three process launches on each of a hundred-odd reruns. The property they
# hold is instead mutated once, directly, by
# `test_144_deleting_a_definition_a_dot_sourced_file_calls_is_rejected`.
FILESYSTEM_CONTROLS = (
    "test_144_every_command_the_benchmark_can_reach_still_resolves",
    "test_145_the_override_declaration_is_required_and_is_not_a_blanket",
    "test_146_an_override_that_would_not_win_is_refused",
)


def _tests() -> list[str]:
    names = sorted(name for name in dir(conformance)
                   if name.startswith("test_") and name not in FILESYSTEM_CONTROLS)
    assert len(names) >= 50, names
    # THE SKIP LIST IS NOT ALLOWED TO NAME SOMETHING THAT NO LONGER EXISTS - a
    # renamed control would otherwise drop silently out of every mutation here.
    for skipped in FILESYSTEM_CONTROLS:
        assert hasattr(conformance, skipped), skipped
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


def _detects_source(expected: str, damaged: str) -> None:
    """Run the battery against a whole damaged runner source.

    For mutations that MOVE code rather than rewrite it - where a
    before/after replacement cannot express the change.
    """
    assert damaged != conformance._runner(), "the mutation changed nothing"
    restore = _install({"runner": damaged})
    try:
        conformance._MEMO.pop("code", None)
        refused = _run_battery()
    finally:
        restore()
        conformance._MEMO.pop("code", None)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


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
        "    $record.Add('cpu_model', (Format-BenchmarkFact "
        "(Get-BenchmarkProperty -InputObject $cpu -Name 'Name')))",
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


# ===========================================================================
# I. THE WINDOWS RUN 1 DEFECT
# ===========================================================================
def test_90_restoring_the_windows_run_1_value_assumption_is_rejected() -> None:
    """THE EXACT STATEMENT THAT ABORTED WINDOWS RUN 1, put back.

    `Get-Item Env:OneDriveCommercial -ErrorAction SilentlyContinue` emits
    nothing on a machine with no work account; the parenthesised pipeline is
    then $null, and `$null.Value` under StrictMode 2.0 is a terminating
    PropertyNotFoundException. This is the mutation the round exists to catch.
    """
    _runner_mutation(
        "test_90",
        "        foreach ($item in @(Get-Item -LiteralPath ('Env:' + $name) "
        "-ErrorAction SilentlyContinue)) {\n"
        "            $value = [string](Get-BenchmarkProperty -InputObject $item -Name 'Value')\n"
        "            if (-not [string]::IsNullOrWhiteSpace($value)) { $roots += $value }\n"
        "        }",
        "        $value = [string](Get-Item -LiteralPath ('Env:' + $name) "
        "-ErrorAction SilentlyContinue).Value\n"
        "        if (-not [string]::IsNullOrWhiteSpace($value)) { $roots += $value }")


def test_91_reading_a_cim_property_off_an_unnormalised_query_is_rejected() -> None:
    """THE SAME DEFECT, ONE CLASS WIDER. A CIM query that matches nothing is an
    empty pipeline exactly as an absent environment variable is."""
    _runner_mutation(
        "test_91",
        "    foreach ($item in @(Get-CimInstance -ClassName Win32_ComputerSystem "
        "-ErrorAction SilentlyContinue)) { $computer = $item }",
        "    $computer = (Get-CimInstance -ClassName Win32_ComputerSystem "
        "-ErrorAction SilentlyContinue)")


def test_92_removing_the_property_guard_is_rejected() -> None:
    _runner_mutation(
        "test_91",
        "    $property = $InputObject.PSObject.Properties[$Name]\n"
        "    if ($null -eq $property) { return $null }\n"
        "    return $property.Value",
        "    return $InputObject.$Name")


def test_93_turning_strict_mode_off_is_rejected() -> None:
    """THE FIX IS NOT TO STOP CHECKING. StrictMode is what turned a silent $null
    into a loud failure; without it the environment record would have carried
    empty strings and nobody would have known the machine was never asked."""
    _runner_mutation(
        "test_93",
        "Set-StrictMode -Version 2.0",
        "Set-StrictMode -Off")


def test_94_defaulting_a_missing_fact_is_rejected() -> None:
    """A FAKE DEFAULT IS WORSE THAN AN ABSENT FIELD, because it reads as an
    answer the machine gave."""
    _runner_mutation(
        "test_94",
        "    if ($null -eq $Value) { return (Get-BenchmarkUnavailable) }",
        "    if ($null -eq $Value) { return '' }")


def test_95_swallowing_a_failure_silently_is_rejected() -> None:
    _runner_mutation(
        "test_94",
        "    } catch {\n"
        "        # UNCLASSIFIABLE IS NOT LOCAL.",
        "    } catch { }\n"
        "    if ($false) {\n"
        "        # UNCLASSIFIABLE IS NOT LOCAL.")


def test_96_classifying_an_unreadable_path_as_local_is_rejected() -> None:
    """THE CONTRACT ASKED THAT A SYNCED REPOSITORY NEVER BE SILENTLY RECORDED AS
    AN UNSYNCHRONISED PATH, and a failed drive lookup falling through to 'local'
    would do exactly that."""
    _runner_mutation(
        "test_22",
        "        return 'unknown'\n"
        "    }\n"
        "    return 'local'",
        "        return 'local'\n"
        "    }\n"
        "    return 'local'")


def test_97_dropping_the_stage_cursor_is_rejected() -> None:
    """WITHOUT IT, THE NEXT FAILURE SAYS ONLY WHAT THE EXCEPTION SAYS - which is
    what made Windows Run 1 take a round trip to diagnose."""
    _runner_mutation(
        "test_95",
        "function New-BenchmarkFailureRecord {",
        "function Build-BenchmarkFailureNote {")


def test_98_failing_to_set_the_cursor_in_setup_is_rejected() -> None:
    _runner_mutation(
        "test_96",
        "    Set-BenchmarkStage -Stage 'setup' -Action 'capturing the environment inventory'",
        "")


def test_99_moving_diagnostics_inside_the_clock_is_rejected() -> None:
    """THE CURSOR MUST COST A TIMING NOTHING."""
    _runner_mutation(
        "test_97",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null",
        "        $watch = [System.Diagnostics.Stopwatch]::StartNew()\n"
        "        Set-BenchmarkStage -Stage 'measurement' -Action 'running'\n"
        "        $Excel.Run([string]$Operation.endpoint) | Out-Null")


def test_100_letting_an_aborted_run_look_like_a_baseline_is_rejected() -> None:
    """WINDOWS RUN 1 PRODUCED NO TIMED OPERATION AT ALL. A run like it must not
    be able to look like evidence."""
    _runner_mutation(
        "test_98",
        "$runComplete = ([bool](([string]::IsNullOrWhiteSpace($abandoned)) -and\n"
        "                       (@($completed).Count -eq $plannedCount) -and ($plannedCount -gt 0)))",
        "$runComplete = $true")


def test_101_exiting_zero_after_an_abort_is_rejected() -> None:
    """A CALLER, A SCHEDULED TASK OR A TRANSCRIPT READER would otherwise record
    an abort as a success."""
    _runner_mutation(
        "test_98",
        "if (-not $runComplete) {\n"
        "    Write-Host ''\n"
        "    Write-Host $baselineStatus -ForegroundColor Red\n"
        "    exit 1\n"
        "}",
        "if ($false) {\n"
        "    Write-Host ''\n"
        "}")


def test_102_counting_an_unsupplied_parameter_as_a_scope_is_rejected() -> None:
    """`@($null).Count` IS 1. Without the null check every full run would have
    been reported as a scoped one, and no run could ever be a baseline."""
    _runner_mutation(
        "test_99",
        "$scoped = ([bool]((($null -ne $Iterations) -and (@($Iterations).Count -gt 0)) -or\n"
        "                  (($null -ne $Operations) -and (@($Operations).Count -gt 0))))",
        "$scoped = ([bool]((@($Iterations).Count -gt 0) -or (@($Operations).Count -gt 0)))")


def test_103_recording_the_hosts_bitness_as_excels_is_rejected() -> None:
    """A 64-BIT HOST AUTOMATING A 32-BIT EXCEL would have been recorded as
    64-bit Excel, in the one field a reader uses to say which build was
    exercised."""
    _runner_mutation(
        "test_101",
        "    $image = Get-BenchmarkExcelImage -Identity $Identity",
        "    $image = [pscustomobject]@{ Path = ''; "
        "Bitness = [string][System.Environment]::Is64BitProcess }")


def test_104_a_powershell_7_only_construct_is_rejected() -> None:
    """THE TARGET IS WINDOWS POWERSHELL 5.1, where none of this parses."""
    _runner_mutation(
        "test_92",
        "    if ($null -eq $InputObject) { return $null }",
        "    if ($null -eq $InputObject) { return $InputObject?.Value }")


def test_105_treating_the_bootstrap_time_as_a_measurement_is_rejected() -> None:
    """WINDOWS RUN 1'S ONLY NUMBER WAS A 68.4 s STAGE-B BOOTSTRAP, and it is
    setup under every reading."""
    _runner_mutation(
        "test_100",
        "'time is evidence about a build, never a baseline.'))",
        "'time is the run\'s first performance figure.'))")


# ===========================================================================
# J. THE WINDOWS RUN 2 DEFECT - THE WRONG AUTHORITY
# ===========================================================================
def test_110_restoring_the_manifest_builder_version_read_is_rejected() -> None:
    """THE EXACT STATEMENT THAT ABORTED WINDOWS RUN 2, put back.

    `stage_b_manifest.json` is a projection of the MODEL side of the
    specification and has never carried a builder version - it cannot, because
    the two are independent authorities. Reading it is a terminating
    PropertyNotFoundException under StrictMode 2.0.
    """
    _runner_mutation(
        "test_23",
        "    $record.Add('builder_version', [string](Get-BenchmarkRequiredProperty `\n"
        "        -InputObject $ReleaseIdentity -Name 'builder_version' `\n"
        "        -Where 'the benchmark plan release identity'))",
        "    $record.Add('builder_version', [string]$Manifest.builder_version)")


def test_111_restoring_the_manifest_build_phase_read_is_rejected() -> None:
    """THE SAME DEFECT ON THE NEXT LINE. Windows Run 2 never reached it because
    the builder version threw first; it would have thrown next."""
    _runner_mutation(
        "test_23",
        "    $record.Add('build_phase', [string](Get-BenchmarkRequiredProperty `\n"
        "        -InputObject $ReleaseIdentity -Name 'build_phase' `\n"
        "        -Where 'the benchmark plan release identity'))",
        "    $record.Add('build_phase', [string]$Manifest.build_phase)")


def test_112_deriving_the_builder_version_from_the_model_version_is_rejected() -> None:
    """THE COLLAPSE P10-3 FORBIDS. They read the same value for this release,
    so a line that made one read the other would pass every value comparison in
    the project - until the first builder-only change."""
    def mutate(plan: dict) -> None:
        plan["release_identity"]["builder_version"] = plan["release_identity"]["model_version"]
        plan["release_identity"]["authorities"]["builder_version"] = (
            "spec/workbook.yaml: model.model_version")

    _plan_mutation("test_114", mutate)


def test_113_a_second_builder_version_authority_in_the_manifest_is_rejected() -> None:
    """ADDING IT TO workbook.yaml WOULD SATISFY POWERSHELL and destroy the
    settlement: the builder would then have two answers and no rule for
    choosing."""
    original = conformance._MEMO.get("manifest_text")
    damaged = (conformance.SPEC / "workbook.yaml").read_text(encoding="utf-8").replace(
        '  model_version: "1.0.0"', '  model_version: "1.0.0"\n  builder_version: "1.0.0"', 1)
    restore = _install({"manifest_text": damaged})
    try:
        # `test_114` reads the manifest through the module's own accessor, so the
        # damaged copy is what it sees.
        refused = _run_battery()
    finally:
        restore()
        if original is not None:
            conformance._MEMO["manifest_text"] = original
    assert any(name.startswith("test_114") for name in refused), refused


def test_114_hard_coding_the_builder_version_in_powershell_is_rejected() -> None:
    """A SECOND LITERAL IS A SECOND AUTHORITY the day one of them moves."""
    _runner_mutation(
        "test_113",
        "    $record.Add('builder_version', [string](Get-BenchmarkRequiredProperty `\n"
        "        -InputObject $ReleaseIdentity -Name 'builder_version' `\n"
        "        -Where 'the benchmark plan release identity'))",
        "    $record.Add('builder_version', '1.0.0')")


def test_115_silently_defaulting_a_missing_builder_version_is_rejected() -> None:
    """REQUIRED RELEASE IDENTITY MAY NOT BE 'unavailable'. An optional
    environment fact may; a value a later comparison depends on may not."""
    _runner_mutation(
        "test_116",
        "    $record.Add('builder_version', [string](Get-BenchmarkRequiredProperty `\n"
        "        -InputObject $ReleaseIdentity -Name 'builder_version' `\n"
        "        -Where 'the benchmark plan release identity'))",
        "    $record.Add('builder_version', (Format-BenchmarkFact "
        "(Get-BenchmarkProperty -InputObject $ReleaseIdentity -Name 'builder_version')))")


def test_116_dropping_the_preflight_release_check_is_rejected() -> None:
    """W2 SPENT SIXTY-EIGHT SECONDS building a workbook it could not attribute.
    The check belongs before the bootstrap, not in the middle of setup."""
    _runner_mutation(
        "test_115",
        "if ($releaseProblems.Count -gt 0) {",
        "if ($false) {")


def test_117_moving_the_release_check_after_the_bootstrap_is_rejected() -> None:
    _runner_mutation(
        "test_115",
        "    Write-Host 'REFUSED, BEFORE ANYTHING WAS BUILT OR MEASURED.' -ForegroundColor Red",
        "    Write-Host 'refused' -ForegroundColor Red")


def test_118_a_plan_with_no_release_identity_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan.pop("release_identity")

    _plan_mutation("test_112", mutate)


def test_119_a_blank_release_value_is_rejected() -> None:
    def mutate(plan: dict) -> None:
        plan["release_identity"]["builder_version"] = ""

    _plan_mutation("test_112", mutate)


def test_120_renaming_an_authority_in_the_projection_is_rejected() -> None:
    """THE ARTIFACT SAYS WHICH FILE ANSWERS FOR EACH VALUE, and that label is
    what makes the independence auditable without this source."""
    def mutate(plan: dict) -> None:
        plan["release_identity"]["authorities"]["builder_version"] = "somewhere"

    _plan_mutation("test_112", mutate)


def test_121_the_runner_reading_python_source_is_rejected() -> None:
    """THE RUNNER CONSUMES A GENERATED PROJECTION. Parsing `workbook_builder.py`
    from PowerShell would make the harness a second reader of an authority it
    has no business interpreting."""
    _runner_mutation(
        "test_112",
        "$plan          = Get-Content -LiteralPath $planPath       -Raw | ConvertFrom-Json",
        "$plan          = Get-Content -LiteralPath $planPath       -Raw | ConvertFrom-Json\n"
        "$builderSource = Get-Content -LiteralPath (Join-Path $pccmRoot "
        "'builder/pccm_builder/workbook_builder.py') -Raw")


def test_122_dropping_the_projection_from_the_artifact_is_rejected() -> None:
    _runner_mutation(
        "test_117",
        "$report.Add('release_identity', $releaseIdentity)",
        "$report.Add('release_identity', 'see the plan')")


# ===========================================================================
# L. THE FIXTURE ON A PROTECTED WORKBOOK
# ===========================================================================
# Every mutation here is a plausible way to "simplify" the correction back into
# the defect Benchmark Run 3 died on, or to make the reset look right while it
# leaves a previous scenario's rates in the table.
def test_130_reintroducing_the_direct_listrow_delete_is_rejected() -> None:
    """THE EXACT CALL BENCHMARK RUN 3 DIED ON, put back into the reset."""
    _runner_mutation(
        "test_130",
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {",
        "    for ($row = $rows; $row -gt $seedRows; $row--) {\n"
        "        $lo.ListRows.Item($row).Delete()\n"
        "    }\n"
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {")


def test_131_restoring_the_structural_row_delete_primitive_is_rejected() -> None:
    """A HELPER NOBODY CALLS IS STILL A HELPER SOMEBODY WILL. The primitive was
    removed rather than stubbed precisely so it cannot come back as a body
    filled into a shell that was left waiting."""
    _runner_mutation(
        "test_130",
        "function Get-IdColumnValues {",
        "function Remove-TableRow {\n"
        "    param($Workbook, [string]$TableName, [int]$RowIndex)\n"
        "    $Workbook.Worksheets.Item(1).ListObjects.Item($TableName)"
        ".ListRows.Item($RowIndex).Delete()\n"
        "}\n\n"
        "function Get-IdColumnValues {")


def test_132_growing_the_table_instead_of_using_a_reserved_row_is_rejected() -> None:
    """THE ADD IS REFUSED ON A PROTECTED SHEET, and putting it back turns the
    next abort - the FX append at step C, which Run 3 never reached - into the
    first one."""
    _runner_mutation(
        "test_130",
        "    $body = @(Get-TableBody -Workbook $Workbook -SheetName $SheetName "
        "-TableName $TableName)",
        "    $lo = $Workbook.Worksheets.Item($SheetName).ListObjects.Item($TableName)\n"
        "    $null = $lo.ListRows.Add()\n"
        "    $body = @(Get-TableBody -Workbook $Workbook -SheetName $SheetName "
        "-TableName $TableName)")


def test_133_clearing_only_the_first_row_below_the_seed_is_rejected() -> None:
    """CLEARING TOO LITTLE IS THE QUIET FAILURE. Row 2 is the only row the
    benchmark's own model writes, so a reset that stops there looks correct on
    every run of this scenario - and carries any other row straight into a
    measurement, where production resolves it as a real rate."""
    _runner_mutation(
        "test_131",
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {",
        "    for ($row = $seedRows + 1; $row -le $seedRows + 1; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {")


def test_134_clearing_only_the_currency_column_is_rejected() -> None:
    """A ROW WITH NO CURRENCY AND A SURVIVING RATE is not a blank row. It is
    invisible to `MatchingFxRows` today and visible to anything that reads the
    column, and half a reset is not a reset."""
    _runner_mutation(
        "test_131",
        "        for ($column = 1; $column -le $columns.Count; $column++) {\n"
        "            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name `",
        "        for ($column = 1; $column -le 1; $column++) {\n"
        "            Set-TableCell -Workbook $Workbook -SheetName $fx.sheet "
        "-TableName $fx.table_name `")


def test_135_accepting_an_empty_string_as_a_genuine_blank_is_rejected() -> None:
    """ClearContents LEAVES Value2 $null. A check that also accepts the empty
    string would pass on a cell production reads as POPULATED, because
    `RawCellText` exits on IsEmpty and an empty string is not empty."""
    _runner_mutation(
        "test_131",
        "            if ($null -ne $body[$row - 1][$column - 1]) {",
        "            if (([string]$body[$row - 1][$column - 1]) -ne '') {")


def test_136_dropping_the_proof_that_the_blanking_took_is_rejected() -> None:
    """A BLANKING NOBODY CHECKED IS AN INTENTION. It is also the one place this
    correction could discover that ClearContents behaves differently under
    protection than the value write does - which has not been observed yet."""
    _runner_mutation(
        "test_131",
        "    for ($row = $seedRows + 1; $row -le $body.Count; $row++) {",
        "    for ($row = $seedRows + 1; $row -le $seedRows; $row++) {")


def test_137_clearing_the_table_header_is_rejected() -> None:
    """BLANKING A BODY IS NOT REWRITING A TABLE. A reset that reached the header
    row would change the contract's shape under a measurement, and the numbers
    would describe a workbook the specification never declared."""
    _runner_mutation(
        "test_133",
        "    $columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName $fx.sheet `",
        "    $Workbook.Worksheets.Item($fx.sheet).ListObjects.Item($fx.table_name)"
        ".HeaderRowRange.ClearContents()\n"
        "    $columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName $fx.sheet `")


def test_138_blanking_the_locked_seed_row_as_well_is_rejected() -> None:
    """THE SEED IS RESTORED FROM A CAPTURE, NOT REBUILT. Blanking row 1 on the
    way past would work by accident today and would destroy the very row the
    capture exists to preserve if the restoration ever moved."""
    _runner_mutation(
        "test_133",
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {",
        "    for ($row = 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {")


def test_139_altering_the_accepted_seed_restoration_is_rejected() -> None:
    """ONE HALF OF THE RESET CHANGED, AND ONLY ONE. Casting the captured rate
    would repair a defective text seed into agreement with the contract, which
    is exactly what the capture rule forbids - and a correction to the blanking
    is no occasion to do it."""
    _runner_mutation(
        "test_132",
        "        -RowIndex 1 -ColumnIndex 2 -Value $Seed.Rate",
        "        -RowIndex 1 -ColumnIndex 2 -Value ([double]$Seed.Rate)")


def test_140_refusing_without_naming_the_table_is_rejected() -> None:
    """A REFUSAL THAT DOES NOT SAY WHICH TABLE sends the next run back to the
    bare 1004 this correction exists to replace."""
    _runner_mutation(
        "test_134",
        "        throw ('the benchmark fixture needs a blank row in ' + $TableName + ' on ' +\n"
        "               $SheetName + ', and the table has no body row at all. Creating one is a ' +",
        "        throw ('the benchmark fixture needs a blank row, ' +\n"
        "               'and the table has no body row at all. Creating one is a ' +")


def test_141_reaching_for_a_protection_release_from_the_harness_is_rejected() -> None:
    """THE BOUNDARY STAYS CLOSED. Option 2 was explicitly NOT authorised, and a
    harness that opened production's structural window for itself would reopen an
    architecture closed on Windows evidence - for a measurement."""
    _runner_mutation(
        "test_137",
        "    $columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName $fx.sheet `",
        "    $null = $Excel.Run('PCCM_ProtectionRelease')\n"
        "    $columns = @(Get-TableColumnNames -Workbook $Workbook -SheetName $fx.sheet `")


def test_142_defining_the_override_above_the_dot_source_is_rejected() -> None:
    """THE WHOLE MECHANISM IS ORDER. An override above the dot-source is
    overwritten by the accepted definition and the runner silently goes back to
    deleting rows - with the corrected source still sitting in the file, and
    every source-reading control above still passing."""
    original = conformance._runner()
    marker = ". (Join-Path $scriptDir 'phase5_gate_b_scenarios.ps1')"
    start = original.index("function Reset-Phase5FxTable {")
    end = original.index("\n}", start) + 3
    moved = original[:start] + original[end:]
    moved = moved.replace(marker, original[start:end] + "\n" + marker, 1)
    _detects_source("test_136", moved)


def test_143_measuring_the_fixture_build_as_a_sample_is_rejected() -> None:
    """A CORRECTION TO SETUP MUST NOT MOVE THE CLOCK. The reset is a longer
    sequence of cell writes than a delete loop was, and the one way that could
    corrupt a baseline is by being timed."""
    _runner_mutation(
        "test_138",
        "    $setupTimings.Add('scenario_fixture_ms', "
        "[double]$fixtureWatch.Elapsed.TotalMilliseconds)",
        "    $sampleMs = [double]$fixtureWatch.Elapsed.TotalMilliseconds")


def test_144_deleting_a_definition_a_dot_sourced_file_calls_is_rejected() -> None:
    """THE DEFECT THAT KILLED PHASE-9 WINDOWS RUN 1, recreated deliberately.

    Deleting `Remove-TableRow` outright was the first attempt at this correction
    and it passed every text control in the conformance module: nothing in the
    benchmark calls it. Two functions in the DOT-SOURCED Gate-B file do, and a
    name a reachable file can call must resolve.

    Driven through a real file because the audit is a cross-file AST closure and
    has nothing to read in an in-memory copy. Run once rather than on every
    mutation above, which is what the skip list pays for."""
    with conformance.RUNNER.open(encoding="utf-8", newline="") as handle:
        source = handle.read()
    start = source.index("function Remove-TableRow {")
    end = source.index("\r\n}\r\n", start) + 5
    damaged = source[:start] + source[end:]
    assert "function Remove-TableRow {" not in damaged
    assert damaged != source

    scratch = conformance.BOOTSTRAP / "__benchmark_missing_definition_tmp.ps1"
    with scratch.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    try:
        done = conformance._resolution_audit(
            scratch, *conformance.BENCHMARK_DECLARED_OVERRIDES)
    finally:
        scratch.unlink()
    assert done.returncode == 1, done.stdout
    assert "UNRESOLVED Remove-TableRow" in done.stdout, done.stdout


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
