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
    # The EXECUTED control-flow proof. It lifts the real functions out of the
    # runner ON DISK through a PowerShell subprocess, so an in-memory damaged
    # copy is invisible to it and it would pass every mutation below while
    # proving nothing. It is mutated instead by `_flow_mutation`, which writes a
    # damaged runner to disk and points the harness at that.
    "test_170_the_window_opens_transactionally_on_every_failure_path",
    "test_171_a_failed_begin_compensates_nothing",
    "test_172_every_post_open_failure_attempts_exactly_one_close",
    "test_173_a_failed_rollback_is_reported_beside_the_failure_that_caused_it",
    "test_174_a_successful_open_closes_nothing_itself",
    "test_175_no_timed_work_follows_any_window_failure",
    # The EXECUTED shape proof, for the same reason: it reads the runner on disk
    # through a subprocess. Mutated by `_shape_mutation` below.
    "test_180_every_planned_run_selects_exactly_one_integer_iteration_count",
    "test_181_the_sample_validator_returns_a_flat_list_and_stays_strict",
    "test_182_a_re_wrapped_list_still_breaks_and_the_guard_refuses_it",
    "test_185_the_shape_harness_reads_the_real_runner_and_starts_no_excel",
    # The EXECUTED bundle proof, same reason again: a subprocess over the gate on
    # disk. Mutated by `_gate_mutation` below.
    "test_211_both_bundles_receive_every_required_artifact",
    "test_212_the_two_bundles_are_isolated_and_proved_identical",
    "test_213_a_missing_required_artifact_refuses_before_excel",
    "test_214_no_stale_stage_b_workbook_travels_into_a_bundle",
    # The EXECUTED reserved-row proof, same reason: a subprocess over the runner on
    # disk. Mutated by `_reserved_mutation` below.
    "test_251_reserved_capacity_is_kept_and_growth_happens_only_when_needed",
    "test_252_identifiers_stop_at_the_semantic_count_and_the_counter_matches",
    "test_253_columns_no_driver_fills_are_genuinely_blank",
    "test_257_the_snapshot_compares_the_reserved_suffix_not_just_the_first_n_rows",
    "test_259_the_reserved_rows_harness_tests_the_shipping_builder",
)

# The control-flow controls, which `_flow_mutation` reruns against damaged rows.
FLOW_CONTROLS = (
    "test_170_the_window_opens_transactionally_on_every_failure_path",
    "test_171_a_failed_begin_compensates_nothing",
    "test_172_every_post_open_failure_attempts_exactly_one_close",
    "test_173_a_failed_rollback_is_reported_beside_the_failure_that_caused_it",
    "test_174_a_successful_open_closes_nothing_itself",
    "test_175_no_timed_work_follows_any_window_failure",
    # Source-side, so it runs here too: a rewrite that kept the observed
    # behaviour but reintroduced the shape is still a failure.
    "test_176_the_rollback_is_the_only_close_inside_open_and_cannot_raise",
)


def _flow_mutation(expected: str, before: str, after: str) -> None:
    """Damage the runner ON DISK, re-run the executed control-flow harness
    against the damaged copy, and require a named control-flow check to refuse.

    THE COUNTS ARE THE PROOF. "Exactly one compensating close, and only on the
    failing paths" cannot be mutated in memory, because the harness is an AST
    lift over a real file. So the damaged runner is written beside the real one -
    the harness resolves dot-sourced paths relative to the script directory - run
    once, and removed.
    """
    with conformance.RUNNER.open(encoding="utf-8", newline="") as handle:
        source = handle.read()
    damaged = source.replace(before.replace("\n", "\r\n"), after.replace("\n", "\r\n"), 1)
    if damaged == source:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the runner")

    scratch = conformance.BOOTSTRAP / "__flow_mutation_tmp.ps1"
    with scratch.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    try:
        rows = conformance._flow_rows(scratch)
    finally:
        scratch.unlink()

    restore = _install({"flow": rows, "runner": damaged.replace("\r\n", "\n")})
    refused = []
    try:
        conformance._MEMO.pop("code", None)
        for name in FLOW_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        restore()
        conformance._MEMO.pop("code", None)
    assert refused, "the mutation survived every control-flow check"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


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


# ===========================================================================
# M. THE FIXTURE MAINTENANCE WINDOW
# ===========================================================================
# Every mutation here is a plausible way to make the window look right while it
# leaves the workbook unprotected under a measurement, or hands the fixture a
# second protection authority.
def _shim_mutation(expected: str, before: str, after: str) -> None:
    """Damage the VBA shim's SOURCE, and prove the anchor still matched."""
    original = conformance._shim()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the shim")
    restore = _install({"shim": damaged})
    try:
        conformance._MEMO.pop("shim_code", None)
        refused = _run_battery()
    finally:
        restore()
        conformance._MEMO.pop("shim_code", None)
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_150_unprotecting_from_powershell_instead_of_the_owner_is_rejected() -> None:
    """A SECOND PROTECTION AUTHORITY. Two of them disagree the first time one is
    not reached, which is precisely the state the single-owner rule exists to
    prevent - and it would also skip the depth counter production relies on."""
    _runner_mutation(
        "test_150",
        "    $reply = [string]$Excel.Run('P10FW_Begin')",
        "    foreach ($ws in $Excel.ActiveWorkbook.Worksheets) { $ws.Unprotect() }\n"
        "    $reply = [string]$Excel.Run('P10FW_Begin')")


def test_151_reaching_the_maintenance_release_path_is_rejected() -> None:
    """`ProtectionRelease` RELEASES WORKBOOK STRUCTURE - `ThisWorkbook.Unprotect`.
    The structural window never touches it, no evidence asks for it, and a
    benchmark is not the place to widen the envelope."""
    _shim_mutation(
        "test_151",
        "    If modProtection.ProtectionBeginStructural(detail) Then",
        "    If modProtection.ProtectionRelease(detail) Then")


def test_152_letting_the_structure_flag_move_is_rejected() -> None:
    """STRUCTURE PROTECTION MUST BE TRUE ON BOTH SIDES OF THE WINDOW. A run that
    stopped checking would not notice the one change this correction promised
    never to make."""
    _runner_mutation(
        "test_151",
        "        if (-not $state.Structure) {\n"
        "            throw ('opening the fixture maintenance window released WORKBOOK STRUCTURE ' +",
        "        if ($false) {\n"
        "            throw ('opening the fixture maintenance window released WORKBOOK STRUCTURE ' +")


def test_153_closing_the_window_after_the_timed_runs_is_rejected() -> None:
    """EVERY TIMED OPERATION WOULD THEN BE MEASURED ON AN UNPROTECTED WORKBOOK,
    which is not the product that ships. The numbers would be real and would
    describe something nobody delivers.

    MOVED BY LINE, NOT BY TEXT. The close sits inside a `finally` whose exact
    indentation a reformat could change; matching the statement rather than the
    block is what makes this mutation survive an edit to its surroundings - and a
    mutation that silently added a SECOND close instead of moving the first would
    prove nothing at all."""
    original = conformance._runner()
    lines = original.split("\n")
    call = "Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest"
    held = [line for line in lines if call in line]
    assert len(held) == 1, held
    kept = [line for line in lines if call not in line]
    assert len(kept) == len(lines) - 1

    after_loop = "    $excel.Run('PCCM_AutomationEnd') | Out-Null"
    assert kept.count(after_loop) == 1, kept.count(after_loop)
    moved = []
    for line in kept:
        if line == after_loop:
            moved.append("    $protectionAfter = Close-BenchmarkFixtureWindow "
                         "-Excel $excel -Manifest $manifest")
        moved.append(line)
    damaged = "\n".join(moved)
    assert damaged.count(call) == 1, "the close was duplicated rather than moved"
    _detects_source("test_152", damaged)


def test_154_dropping_the_finally_that_closes_on_a_raise_is_rejected() -> None:
    """A FIXTURE THAT RAISES WOULD LEAVE THE WORKBOOK UNPROTECTED, and the next
    thing to touch it would be a timed run."""
    # ANCHORED ON THE TWO KEYWORDS, not on the builder call between them: the
    # region gained a branch when the bulk builder arrived, and a mutation pinned
    # to the old body would silently stop changing anything.
    _runner_mutation(
        "test_153",
        "    $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest\n"
        "    try {\n",
        "    $null = Open-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest\n")


def test_155_trusting_the_owners_ok_without_re_verifying_is_rejected() -> None:
    """"IT SAID OK" IS NOT RESTORATION. The independent count is the only thing
    that would catch a sheet that came back unprotected while the owner still
    reported success."""
    _runner_mutation(
        "test_154",
        "    return (Assert-BenchmarkProtectionApplied -Excel $Excel -Manifest $Manifest `\n"
        "        -Stage 'after the fixture maintenance window was closed')",
        "    return (Get-BenchmarkProtectionState -Excel $Excel)")


def test_156_hard_coding_the_sheet_count_is_rejected() -> None:
    """THE COUNT COMES FROM THE MANIFEST'S PROTECTION PROJECTION. A literal would
    keep passing at fourteen after the contract declared a fifteenth sheet, and
    the unprotected one would be the new one."""
    _runner_mutation(
        "test_154",
        "    $expected = @($Manifest.protection.sheets).Count",
        "    $expected = 14")


def test_157_continuing_after_a_failed_close_is_rejected() -> None:
    """A PROTECTION FAILURE MUST ABORT BEFORE TIMING. Warning and carrying on
    would publish a baseline taken from a workbook whose protection nobody
    restored."""
    _runner_mutation(
        "test_155",
        "        throw ('the fixture maintenance window could not be closed and protection was ' +",
        "        Write-Host ('the fixture maintenance window could not be closed: ' + $reply)\n"
        "        return (Get-BenchmarkProtectionState -Excel $Excel)\n"
        "        $unreachable = ('the fixture maintenance window could not be closed: ' +")


def test_158_skipping_the_pre_open_protection_check_is_rejected() -> None:
    """OPENING A WINDOW OVER AN ALREADY-UNPROTECTED WORKBOOK would close onto a
    state nobody established, and the run would report a restoration it never
    performed."""
    _runner_mutation(
        "test_152",
        "    $null = Assert-BenchmarkProtectionApplied -Excel $Excel -Manifest $Manifest `\n"
        "        -Stage 'before the fixture maintenance window was opened'",
        "    $null = Get-BenchmarkProtectionState -Excel $Excel")


def test_159_importing_the_shim_over_a_declared_production_module_is_rejected() -> None:
    """A MODULE THE MANIFEST DECLARES IS PRODUCTION, and importing one at runtime
    would be replacing production rather than adding a test surface."""
    _runner_mutation(
        "test_156",
        "    if ($declared -contains $script:FixtureWindowModule) {",
        "    if ($false) {")


def test_160_assuming_the_import_worked_is_rejected() -> None:
    """AN IMPORT THAT RAISED NOTHING AND A PROJECT THAT WILL CALL IT are different
    things - Gate-B run 7 met a project that answered one call while an unreached
    procedure still held a declaration the parser rejected."""
    _runner_mutation(
        "test_156",
        "    $ping = [string]$Excel.Run('P10FW_Ping')",
        "    $ping = 'OK|' + $script:FixtureWindowModule")


def test_161_reading_protectcontents_across_com_again_is_rejected() -> None:
    """PROBE RUN 8'S DEFECT, PUT BACK. `Worksheet.ProtectContents` on an object
    PowerShell has no type information for is a terminating
    PropertyNotFoundException under StrictMode 2.0."""
    _runner_mutation(
        "test_157",
        "    $raw = [string]$Excel.Run('P10FW_State')",
        "    $raw = [string]$Excel.ActiveWorkbook.Worksheets.Item(1).ProtectContents")


def test_162_defaulting_a_missing_protection_field_is_rejected() -> None:
    """A MISSING FIELD IS NOT A FALSE ONE. Defaulting `protected` would let a
    state string that never carried the count read as a full house."""
    _runner_mutation(
        "test_157",
        "        if (-not $state.ContainsKey($required)) {\n"
        "            throw ('the protection state is missing ' + $required + ': ' + $raw)\n"
        "        }",
        "        if (-not $state.ContainsKey($required)) { $state[$required] = 'True' }")


def test_163_wrapping_the_whole_run_in_the_window_is_rejected() -> None:
    """THE WINDOW IS SETUP ONLY. Widening it to cover work that does not need it -
    here the seed write, which Run 3 proved an external COM caller can make on a
    protected sheet - is how a maintenance window becomes an unprotected run."""
    _runner_mutation(
        "test_158",
        "    } finally {\n"
        "        $protectionAfter = Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest\n"
        "    }",
        "        Set-NamedValue -Workbook $wb -DefinedName 'nmSeed' -Value ([double]1)\n"
        "    } finally {\n"
        "        $protectionAfter = Close-BenchmarkFixtureWindow -Excel $excel -Manifest $manifest\n"
        "    }")


def test_164_restoring_the_structural_delete_under_the_window_is_rejected() -> None:
    """THE WINDOW IS NOT A LICENCE TO GO BACK. With protection released the old
    `ListRow.Delete` would now SUCCEED - and would silently resume destroying the
    contract's reserved, validated rows."""
    _runner_mutation(
        "test_130",
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {",
        "    for ($row = $rows; $row -gt $seedRows; $row--) {\n"
        "        $lo.ListRows.Item($row).Delete()\n"
        "    }\n"
        "    for ($row = $seedRows + 1; $row -le $rows; $row++) {\n"
        "        for ($column = 1; $column -le $columns.Count; $column++) {")


def test_165_the_shim_deciding_instead_of_forwarding_is_rejected() -> None:
    """A SHIM THAT REPORTED SUCCESS ON ITS OWN AUTHORITY would hide exactly the
    failure the harness exists to abort on."""
    _shim_mutation(
        "test_150",
        "    If modProtection.ProtectionEndStructural(detail) Then\n"
        "        P10FW_End = \"OK|depth=\" & CStr(modProtection.ProtectionStructuralDepth())\n"
        "    Else\n"
        "        P10FW_End = \"FAIL|\" & detail\n"
        "    End If",
        "    detail = detail\n"
        "    P10FW_End = \"OK|depth=0\"")


# ===========================================================================
# N. THE TRANSACTIONAL OPEN
# ===========================================================================
# Each of these damages the runner ON DISK and re-runs the EXECUTED control-flow
# harness against the damaged copy, so what refuses is a macro call COUNT rather
# than a text match. A mutation that leaves the source looking right and changes
# what PowerShell actually does is caught here and nowhere else.
def test_170_dropping_the_compensating_close_is_rejected() -> None:
    """THE GAP THIS ROUND CLOSED, REOPENED. `P10FW_Begin` succeeds, a post-open
    check refuses, and nothing closes what was opened: the error escapes to the
    abandon path leaving a half-open window behind."""
    _flow_mutation(
        "test_172",
        "    } catch {\n"
        "        $original = [string]$_.Exception.Message\n"
        "        $recovery = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "        throw ($original + ' The window was open when this failed, so a compensating ' +\n"
        "               'close was attempted: ' + $recovery)\n"
        "    }",
        "    } catch {\n"
        "        throw ([string]$_.Exception.Message)\n"
        "    }")


def test_171_compensating_a_failed_begin_is_rejected() -> None:
    """NOTHING WAS OPENED, SO NOTHING IS OWED. Moving the Begin inside the guarded
    region makes a refused open decrement a depth nobody raised, which
    `ProtectionEndStructural` reports as a close without an open."""
    _flow_mutation(
        "test_171",
        "    $reply = [string]$Excel.Run('P10FW_Begin')\n"
        "    if ($reply -notlike 'OK|*') {\n"
        "        throw ('the fixture maintenance window could not be opened: ' + $reply)\n"
        "    }\n",
        "    try {\n"
        "    $reply = [string]$Excel.Run('P10FW_Begin')\n"
        "    if ($reply -notlike 'OK|*') {\n"
        "        throw ('the fixture maintenance window could not be opened: ' + $reply)\n"
        "    }\n"
        "    } catch {\n"
        "        $null = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "        throw ([string]$_.Exception.Message)\n"
        "    }\n")


def test_172_closing_inside_a_successful_open_is_rejected() -> None:
    """A DOUBLE-DECREMENT. Both the open and the caller's finally would close the
    same window, taking modProtection's depth to -1 - and the second close would
    re-apply protection over a fixture that had not been built yet."""
    _flow_mutation(
        "test_174",
        "        return $state\n"
        "    } catch {",
        "        $null = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "        return $state\n"
        "    } catch {")


def test_173_compensating_in_a_finally_instead_of_a_catch_is_rejected() -> None:
    """A `finally` RUNS ON THE SUCCESS PATH TOO, so it closes a window the caller
    is still expecting to hold - and if it threw, its exception would REPLACE the
    original failure, losing the diagnosis the rollback exists to accompany."""
    _flow_mutation(
        "test_17",
        "    } catch {\n"
        "        $original = [string]$_.Exception.Message\n"
        "        $recovery = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "        throw ($original + ' The window was open when this failed, so a compensating ' +\n"
        "               'close was attempted: ' + $recovery)\n"
        "    }",
        "    } finally {\n"
        "        $null = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "    }")


def test_174_losing_the_original_failure_behind_the_rollback_is_rejected() -> None:
    """THE ROLLBACK IS NOT THE NEWS. Reporting only that the window was closed
    would hide WHY it had to be, and the run would abort with a sentence that
    explains nothing."""
    _flow_mutation(
        "test_17",
        "        throw ($original + ' The window was open when this failed, so a compensating ' +\n"
        "               'close was attempted: ' + $recovery)",
        "        throw ('the fixture maintenance window was rolled back: ' + $recovery)")


def test_175_a_rollback_that_can_raise_is_rejected() -> None:
    """IT IS CALLED FROM A CATCH THAT IS ABOUT TO RETHROW. A rollback that threw
    would destroy the composed diagnosis and replace it with its own."""
    _flow_mutation(
        "test_17",
        "        if ($reply -notlike 'OK|*') {\n"
        "            return ('IT REFUSED, so worksheet protection is NOT restored and this ' +\n"
        "                    'workbook must not be measured - ' + $reply)\n"
        "        }",
        "        if ($reply -notlike 'OK|*') {\n"
        "            throw ('the compensating close refused: ' + $reply)\n"
        "        }")


def test_176_swallowing_a_raised_rollback_is_rejected() -> None:
    """A ROLLBACK THAT RAISED AND REPORTED SUCCESS is the worst of both: the
    workbook is unprotected and the transcript says it is fine."""
    _flow_mutation(
        "test_173",
        "        return ('IT RAISED, so worksheet protection is NOT restored and this workbook ' +\n"
        "                'must not be measured - ' + [string]$_.Exception.Message)",
        "        return 'it succeeded and protection is restored - recovered'")


def test_177_attempting_the_compensating_close_twice_is_rejected() -> None:
    """EXACTLY ONE. A second attempt over a window the first one already closed
    decrements modProtection's depth below zero, which it reports as a close
    without an open - and the retry would look like a fix."""
    _flow_mutation(
        "test_172",
        "        $recovery = Invoke-BenchmarkWindowRollback -Excel $Excel\n",
        "        $recovery = Invoke-BenchmarkWindowRollback -Excel $Excel\n"
        "        $recovery = $recovery + '; retried: ' + "
        "(Invoke-BenchmarkWindowRollback -Excel $Excel)\n")


def test_178_swallowing_a_post_open_failure_and_proceeding_is_rejected() -> None:
    """A CLOSED WINDOW AND A CONTINUING RUN. The fixture would then be built
    against a protected workbook, and the caller's finally would close a window
    that was already closed."""
    _flow_mutation(
        "test_17",
        "        throw ($original + ' The window was open when this failed, so a compensating ' +\n"
        "               'close was attempted: ' + $recovery)",
        "        Write-Host ($original + ' rolled back: ' + $recovery)\n"
        "        return (Get-BenchmarkProtectionState -Excel $Excel)")


def test_179_a_flow_harness_that_tests_a_copy_is_rejected() -> None:
    """A HARNESS THAT TESTED A REIMPLEMENTATION WOULD PROVE NOTHING. It lifts the
    real functions out of the shipping file by AST; a definition of its own would
    make every count above a statement about the harness."""
    original = conformance.FLOW_HARNESS.read_text(encoding="utf-8")
    # AFTER THE LIFT, NOT BEFORE IT. A definition above the loop is simply
    # overwritten by the real one it lifts, which would make this mutation a
    # no-op that proves the opposite of what it claims.
    anchor = "# --- THE CALLER'S OWN try/finally, LIFTED FROM THE RUNNER"
    assert anchor in original
    damaged = original.replace(
        anchor,
        "function Open-BenchmarkFixtureWindow { param($Excel, $Manifest) return 'ok' }\n"
        + anchor, 1)
    assert damaged != original, "the flow-harness anchor moved"
    scratch = conformance.PCCM_ROOT / "tests" / "__flow_harness_tmp.ps1"
    with scratch.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    saved = conformance.FLOW_HARNESS
    refused = []
    try:
        conformance.FLOW_HARNESS = scratch
        conformance._MEMO.pop("flow", None)
        for name in FLOW_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        conformance.FLOW_HARNESS = saved
        conformance._MEMO.pop("flow", None)
        scratch.unlink()
    assert refused, "a reimplemented function in the flow harness survived"
    assert any(name.startswith("test_17") for name in refused), refused


# ===========================================================================
# O. THE TWO SHAPE DEFECTS
# ===========================================================================
# The shape controls read the runner ON DISK through a PowerShell subprocess, so
# they are excluded from the generic battery above and mutated here instead - by
# writing a damaged runner to disk and re-running the harness against it. A CLR
# type and a problem COUNT cannot be mutated in memory.
SHAPE_CONTROLS = (
    "test_180_every_planned_run_selects_exactly_one_integer_iteration_count",
    "test_181_the_sample_validator_returns_a_flat_list_and_stays_strict",
    "test_182_a_re_wrapped_list_still_breaks_and_the_guard_refuses_it",
    # Source-side, so they run here too.
    "test_183_the_two_place_array_contract_is_kept_at_both_of_its_sites",
    "test_184_no_local_reuses_a_typed_script_parameter_name",
    "test_186_the_median_still_needs_three_valid_warm_samples",
    "test_187_the_fixture_window_is_byte_identical_to_the_run_that_proved_it",
)


def _shape_mutation(expected: str, before: str, after: str) -> None:
    """Damage the runner ON DISK and re-run the executed shape harness on it."""
    with conformance.RUNNER.open(encoding="utf-8", newline="") as handle:
        source = handle.read()
    damaged = source.replace(before.replace("\n", "\r\n"), after.replace("\n", "\r\n"), 1)
    if damaged == source:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the runner")

    scratch = conformance.BOOTSTRAP / "__shape_mutation_tmp.ps1"
    with scratch.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    try:
        rows = conformance._shape_rows(scratch)
    finally:
        scratch.unlink()

    restore = _install({"shape": rows, "runner": damaged.replace("\r\n", "\n")})
    refused = []
    try:
        conformance._MEMO.pop("code", None)
        for name in SHAPE_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        restore()
        conformance._MEMO.pop("code", None)
    assert refused, "the mutation survived every shape check"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_180_re_wrapping_the_problem_list_is_rejected() -> None:
    """THE EXACT DEFECT THE FIRST TIMED RUN DIED ON. `@()` applied to a helper that
    already returns `,@(...)` collects ONE object into a new one-element array, so
    the count is 1 whatever the sample found: every execution INVALID, rendered as
    `System.Object[]`."""
    _shape_mutation(
        "test_18",
        "    $problems = Test-BenchmarkSample -Operation $Operation -Evidence $evidence `\n"
        "        -RequestedIterations $RequestedIterations",
        "    $problems = @(Test-BenchmarkSample -Operation $Operation -Evidence $evidence `\n"
        "        -RequestedIterations $RequestedIterations)")


def test_181_dropping_the_shape_guard_is_rejected() -> None:
    """THE TWO-PLACE CONTRACT NEEDS A CHECK AT THE CALL. Without it the next edit
    that half-keeps it reappears as every sample being invalid for an unreadable
    reason, which is a week of confusion rather than an abort."""
    _shape_mutation(
        "test_18",
        "    Assert-BenchmarkProblemList -Problems $problems -Where ([string]$Operation.key)\n",
        "")


def test_182_flattening_a_malformed_list_instead_of_refusing_is_rejected() -> None:
    """REPAIRING THE SHAPE HIDES THE DEFECT. A guard that flattened would let a
    re-wrapped list through, and the run would go on measuring with a validator
    nobody could trust."""
    _shape_mutation(
        "test_18",
        "        if ($problem -is [System.Array]) {\n"
        "            throw ('the sample validator returned a NESTED list for ' + $Where +",
        "        if ($problem -is [System.Array]) {\n"
        "            $problem = ($problem -join '; ')\n"
        "            $unreachable = ('the sample validator returned a NESTED list for ' + $Where +")


def test_183_dropping_the_comma_from_the_helper_is_rejected() -> None:
    """THE COMMA IS WHY AN EMPTY RESULT SURVIVES. `return @($problems)` on an empty
    list emits NOTHING to the pipeline, so the caller's direct assignment gets
    $null - and a $null problem list is not an empty one."""
    _shape_mutation(
        "test_18",
        "    return ,@($problems)\n}",
        "    return @($problems)\n}")


def test_184_restoring_the_parameter_name_collision_is_rejected() -> None:
    """THE EXACT DEFECT THAT ABORTED THE SIMULATION SETUP. `$iterations` IS the
    script parameter `[int[]]$Iterations` - names are case-insensitive - so every
    assignment is coerced back to `[int[]]`, and `[double]` of a one-element array
    raises while `[string]` of it still looks right."""
    _shape_mutation(
        "test_180",
        "        $runIterations = $null\n"
        "        if ($null -ne $run.iterations) { $runIterations = [int]$run.iterations }",
        "        $iterations = $null\n"
        "        if ($null -ne $run.iterations) { $iterations = [int]$run.iterations }\n"
        "        $runIterations = $iterations")


def test_185_selecting_an_element_out_of_the_iteration_count_is_rejected() -> None:
    """THE PATCH THAT WOULD HAVE HIDDEN IT. `[double]$iterations[0]` makes the
    symptom go away and leaves the shape wrong, so the value would still be an
    array everywhere else - and on a plan that ever carried a list it would
    silently measure only the first count."""
    _shape_mutation(
        "test_180",
        "        $runIterations = $null\n"
        "        if ($null -ne $run.iterations) { $runIterations = [int]$run.iterations }",
        "        $iterations = $null\n"
        "        if ($null -ne $run.iterations) { $iterations = [int]$run.iterations }\n"
        "        $runIterations = $iterations[0]")


def test_186_taking_the_whole_iteration_array_into_the_control_is_rejected() -> None:
    """THE FULL DECLARED LIST WHERE ONE COUNT BELONGS. `$declaredIterations` is
    every count the scenario declares; writing it to the control would ask the
    workbook for three numbers at once."""
    _shape_mutation(
        "test_180",
        "        if ($null -ne $run.iterations) { $runIterations = [int]$run.iterations }",
        "        if ($null -ne $run.iterations) { $runIterations = $declaredIterations }")


def test_187_counting_invalid_samples_toward_the_median_is_rejected() -> None:
    """A MEDIAN OF TWO GOOD RUNS AND A REFUSAL is not a measurement of anything,
    and the shape defect meant this gate was never reached on Windows."""
    _runner_mutation(
        "test_186",
        "        $validWarm = @($warm | Where-Object { $_.Valid })",
        "        $validWarm = @($warm)")


def test_188_computing_a_median_from_fewer_than_three_warm_samples_is_rejected() -> None:
    """THE COUNT COMES FROM THE PLAN, and the median exists only when every warm
    sample was valid."""
    _runner_mutation(
        "test_186",
        "        if ($validWarm.Count -eq [int]$run.warm_runs) {",
        "        if ($validWarm.Count -ge 1) {")


def test_189_touching_the_runtime_proven_fixture_window_is_rejected() -> None:
    """THE WINDOW IS RUNTIME PROVEN AND IS NOT REOPENED. These two defects are in
    the run loop, downstream of it; a shape correction that edited the window
    would put a proven architecture back in doubt."""
    _runner_mutation(
        "test_187",
        "    $null = Assert-BenchmarkProtectionApplied -Excel $Excel -Manifest $Manifest `\n"
        "        -Stage 'before the fixture maintenance window was opened'",
        "    $null = Get-BenchmarkProtectionState -Excel $Excel")


def test_190_re_wrapping_the_onedrive_roots_is_rejected() -> None:
    """THE SAME PAIRING AT ITS OTHER SITE. Double-wrapped, every OneDrive root
    became one nested array: a workbook under OneDrive was recorded as LOCAL and
    the roots field held a list containing a list."""
    _runner_mutation(
        "test_183",
        "    $roots = Get-BenchmarkOneDriveRoots",
        "    $roots = @(Get-BenchmarkOneDriveRoots)")


# The equivalence gate and the evidence record are neither the runner nor the
# shim, so they need their own mutation paths. Both damage a file the conformance
# module reads, run the named controls against the damaged text, and restore.
def _equivalence_mutation(expected: str, before: str, after: str) -> None:
    original = conformance._equiv_harness()
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the gate")
    restore = _install({"equiv_harness": damaged})
    try:
        refused = _run_battery()
    finally:
        restore()
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def _record_mutation(expected: str, before: str, after: str) -> None:
    """Damage the Windows evidence record on disk, run the battery, restore.

    ON DISK, because the record is read through `Path.read_text` at assertion
    time rather than through the memo - there is nothing to install.
    """
    path = conformance.RUN_EVIDENCE
    original = path.read_text(encoding="utf-8")
    damaged = original.replace(before, after, 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the record")
    path.write_text(damaged, encoding="utf-8")
    try:
        refused = _run_battery()
    finally:
        path.write_text(original, encoding="utf-8")
    assert refused, "the mutation survived the whole conformance battery"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


# ===========================================================================
# P. THE BULK FIXTURE
# ===========================================================================
# The heaviest mutations in this file, because the bulk builder is the one place
# where a fixture could quietly become a second business engine.
def test_190_publishing_a_calculation_result_from_the_fixture_is_rejected() -> None:
    """THE WORST THING A FIXTURE COULD DO. A builder that wrote into _Calc would
    hand the benchmark a workbook with the answers already in it, and every
    Calculate median after it would be measuring a no-op."""
    _runner_mutation(
        "test_191",
        "    # --- H. THE FIXTURE ENDS COHERENT ---------------------------------------\n"
        "    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `",
        "    Set-BenchmarkRangeBlock -Workbook $Workbook -SheetName '_Calc' `\n"
        "        -TableName 'tblCalcAnnual' -FirstRow 1 -FirstColumn 1 -Block $block `\n"
        "        -Description 'tblCalcAnnual'\n"
        "    $null = Assert-Phase5StructurallyCoherent -Excel $Excel `")


def test_191_publishing_a_simulation_result_from_the_fixture_is_rejected() -> None:
    """_SimData IS PRODUCTION'S. A fixture that wrote a simulation row would make
    the Simulation median a measurement of nothing."""
    _runner_mutation(
        "test_191",
        "    $inflationGrid = $null\n",
        "    Set-TableCell -Workbook $Workbook -SheetName '_SimData' -TableName 'tblSimResults' `\n"
        "        -RowIndex 1 -ColumnIndex 1 -Value ([double]1)\n"
        "    $inflationGrid = $null\n")


def test_192_fabricating_a_state_label_from_the_fixture_is_rejected() -> None:
    """CURRENT IS A CLAIM ABOUT WORK THAT WAS DONE. Writing the label without the
    work is the one shortcut that would make every correctness gate downstream
    pass on a lie."""
    _runner_mutation(
        "test_191",
        "    # --- F. ONE REAL PCCM_ApplyTimeline -------------------------------------",
        "    Set-NamedValue -Workbook $Workbook -DefinedName 'nmStructuralState' `\n"
        "        -Value ([double]1)\n"
        "    # --- F. ONE REAL PCCM_ApplyTimeline -------------------------------------")


def test_193_invoking_a_timed_endpoint_from_the_fixture_is_rejected() -> None:
    """THE FIXTURE MAY INVOKE ApplyTimeline AND NOTHING ELSE. A Calculate inside
    setup would warm exactly the thing the cold sample exists to measure."""
    _runner_mutation(
        "test_191",
        "    $applied = Invoke-Phase5ProductionOperation -Excel $Excel `\n"
        "        -Operation 'PCCM_ApplyTimeline' -Stage 'the bulk fixture structural baseline'",
        "    $applied = Invoke-Phase5ProductionOperation -Excel $Excel `\n"
        "        -Operation 'PCCM_ApplyTimeline' -Stage 'the bulk fixture structural baseline'\n"
        "    $null = Invoke-Phase5ProductionOperation -Excel $Excel `\n"
        "        -Operation 'PCCM_Calculate' -Stage 'warming the calculation'")


def test_194_reproducing_the_year_columns_in_the_fixture_is_rejected() -> None:
    """THE GRIDS' SHAPE IS PRODUCTION'S. A builder that added year columns itself
    would be reimplementing modProfiling.SetYearColumns, and the benchmark would
    stop testing that production can do it."""
    # RE-ANCHORED: the weight builder now returns a record, because `return $block`
    # emitted the matrix and the pipeline enumerated it into a rank-1 array.
    _runner_mutation(
        "test_192",
        "        $prepared = New-BenchmarkWeightBlock -Workbook $Workbook -Grid $grid `",
        "        $null = $Workbook.Worksheets.Item($grid.sheet).ListObjects.Item(\n"
        "            $grid.table_name).ListColumns.Add()\n"
        "        $prepared = New-BenchmarkWeightBlock -Workbook $Workbook -Grid $grid `")


def test_195_hard_coding_the_identifier_format_is_rejected() -> None:
    """THE PREFIX AND PAD COME FROM THE MANIFEST'S COUNTER PROJECTION. A literal
    would keep passing after the contract changed either one, and the register
    would be written with identifiers production would never issue."""
    _runner_mutation(
        "test_193",
        "    $prefix = [string]$Counter.prefix\n"
        "    $pad = [int]$Counter.pad_width",
        "    $prefix = 'CL-'\n"
        "    $pad = 3")


def test_196_not_comparing_the_issued_identifier_with_the_model_is_rejected() -> None:
    """TWO AUTHORITIES ON THE SAME IDENTITY. The model names CL-001.. and
    production issues in sequence; if they ever disagreed the fixture would be
    describing a different workbook, silently."""
    _runner_mutation(
        "test_193",
        "        if ($issued -cne $declared) {",
        "        if ($false) {")


def test_197_leaving_the_counter_at_zero_is_rejected() -> None:
    """THE COUNTER IS THE MODEL'S RECORD OF EVERY IDENTIFIER EVER ISSUED. Left at
    zero it disagrees with every row present, and modDrivers.HighestIssued exists
    to catch exactly that."""
    _runner_mutation(
        "test_193",
        "        CounterValue = [double]$drivers.Count",
        "        CounterValue = [double]0")


def test_198_writing_the_register_cell_by_cell_is_rejected() -> None:
    """THE WHOLE POINT. A per-cell loop puts the O(N x cols) COM cost back, and on
    LARGE that is the difference the four-hour abort measured."""
    _runner_mutation(
        "test_194",
        "    $rows = [int]$Block.GetLength(0)\n"
        "    $columns = [int]$Block.GetLength(1)",
        "    $rows = [int]$Block.GetLength(0)\n"
        "    $columns = [int]$Block.GetLength(1)\n"
        "    for ($r = 1; $r -le $rows; $r++) { $null = $r }")


def test_199_accepting_a_jagged_array_for_a_block_write_is_rejected() -> None:
    """A JAGGED ARRAY IS SILENTLY NOT THE VARIANT SHAPE EXCEL ACCEPTS. Dropping
    the rank check would turn a wrong-shaped block into a runtime surprise on the
    machine rather than a refusal here."""
    _runner_mutation(
        "test_194",
        "    if ($Block.Rank -ne 2) {",
        "    if ($false) {")


def test_200_deleting_register_rows_in_the_fixture_is_rejected() -> None:
    """THE DESTRUCTIVE PATH THIS PROJECT ALREADY REMOVED ONCE. Shrinking a register
    would delete the contract's reserved, validated rows - and the builder runs
    against a fresh workbook where it can never be needed."""
    # A REAL DELETE, not the old refusal - restoring that is test_250's subject
    # now. This is the destructive path the control exists to forbid.
    _runner_mutation(
        "test_195",
        "    if ($current -ge $MinimumRows) { return $current }",
        "    if ($current -gt $MinimumRows) {\n"
        "        $lo = $Workbook.Worksheets.Item($SheetName).ListObjects.Item($TableName)\n"
        "        for ($i = $current; $i -gt $MinimumRows; $i--) { $lo.ListRows.Item($i).Delete() }\n"
        "    }\n"
        "    if ($current -ge $MinimumRows) { return $current }")


def test_201_not_proving_the_register_grew_is_rejected() -> None:
    """A GROW THAT SILENTLY DID NOTHING leaves the one block write landing outside
    the table, and Excel would not complain."""
    _runner_mutation(
        "test_195",
        "    if ($after -ne $MinimumRows) {",
        "    if ($false) {")


def test_202_making_the_bulk_path_the_default_is_rejected() -> None:
    """THE ACCEPTED SMALL AND MEDIUM BASELINES WERE BUILT BY THE ENDPOINT PATH. A
    default flip would silently change what they are comparable with, and nothing
    in the artifact would have said so."""
    _runner_mutation(
        "test_196",
        "    [string]$FixtureMode = 'Endpoints'",
        "    [string]$FixtureMode = 'Bulk'")


def test_203_dropping_the_fixture_mode_from_the_artifact_is_rejected() -> None:
    """TWO METHODS THAT REACH THE SAME STATE ARE STILL TWO METHODS. A baseline that
    does not say which built it cannot be compared with one that does."""
    _runner_mutation(
        "test_196",
        "$report.Add('fixture_mode', [string]$FixtureMode)\n",
        "")


def test_204_building_over_a_populated_register_is_rejected() -> None:
    """IT WRITES IDENTIFIERS FROM SEQUENCE 1. Over existing rows that means
    duplicates and a counter that disagrees with the highest identifier present."""
    _runner_mutation(
        "test_198",
        "        if ($existing.Count -ne 0) {",
        "        if ($false) {")


def test_205_not_reading_the_register_back_before_the_sync_is_rejected() -> None:
    """PRODUCTION IS ABOUT TO SYNCHRONISE BOTH GRIDS FROM THAT REGISTER. If the
    block landed wrong, ApplyTimeline would faithfully key the grids to the wrong
    rows and everything downstream would agree with itself."""
    _runner_mutation(
        "test_198",
        "            if ([string]$ids[$i] -cne [string]$expected[$i]) {",
        "            if ($false) {")


def test_206_assuming_the_profiling_row_order_is_rejected() -> None:
    """modProfiling.SyncRows REBUILDS THE GRID FROM THE REGISTER, and nothing binds
    its physical order to the order the register was written in. Assuming it would
    put each driver's weights on some other driver's row."""
    _runner_mutation(
        "test_199",
        "        $driver = $byId[$rows[$r]]",
        "        $driver = @($Drivers)[$r]")


def test_207_writing_weights_over_a_non_contiguous_grid_is_rejected() -> None:
    """ONE RECTANGLE CANNOT STRADDLE A GAP. A blank key inside the keyed rows means
    the block would land on rows that belong to nobody."""
    _runner_mutation(
        "test_199",
        "        if ([string]::IsNullOrWhiteSpace([string]$body[$r][0])) {",
        "        if ($false) {")


def test_208_a_gate_that_only_builds_one_way_is_rejected() -> None:
    """THE COMPARISON IS THE GATE. A gate that built only the optimised path would
    prove the optimised path runs, which is not the question."""
    _equivalence_mutation(
        "test_200",
        "foreach ($mode in @('Endpoints', 'Bulk')) {",
        "foreach ($mode in @('Bulk')) {")


def test_209_a_gate_that_does_not_compare_the_calculation_fingerprint_is_rejected() -> None:
    """PRODUCTION'S OWN VERDICT ON THE TWO WORKBOOKS. Field-for-field equality of
    the inputs is necessary and not sufficient: the fingerprint is what says the
    model calculates to the same thing."""
    _equivalence_mutation(
        "test_200",
        "    $calcFingerprint = [string]$excel.Run('PCCM_CalculationFingerprint')",
        "    $calcFingerprint = 'not-compared'")


def test_210_a_gate_that_judges_its_own_evidence_is_rejected() -> None:
    """THE EVIDENCE AND THE VERDICT ARE TWO AUTHORITIES. A gate that exited
    non-zero on its own comparison would be deciding the thing the control exists
    to decide, and a vacuous comparison would then look like a pass."""
    _equivalence_mutation(
        "test_200",
        "    if ($reference.CalcFingerprint -ceq $optimised.CalcFingerprint) {",
        "    if ($false) { exit 1 }\n"
        "    if ($reference.CalcFingerprint -ceq $optimised.CalcFingerprint) {")


def test_211_claiming_the_gate_passed_while_it_is_unrun_is_rejected() -> None:
    """THE RECORD MAY SAY OUTSTANDING OR IT MAY SAY ALL-MATCHED. It may not say
    neither, and it may not report a difference and still read as passing."""
    _record_mutation(
        "test_201",
        "**NOT YET RUN.** It is a Windows verification",
        "**PASSED.** It is a Windows verification")


def test_212_softening_the_aborted_large_record_is_rejected() -> None:
    """>4 HOURS IN SETUP IS EVIDENCE ABOUT THE FIXTURE METHOD. A record that
    dropped the sentence refusing the wrong reading would let the next reader
    conclude that a Large calculation takes four hours."""
    _record_mutation(
        "test_202",
        "It **does NOT mean** any of the following",
        "It could mean any of the following")


def test_213_dropping_the_cost_derivation_from_the_record_is_rejected() -> None:
    """A NUMBER WITHOUT ITS DERIVATION IS A CLAIM. Windows has to be able to check
    the reduction, which means seeing where each count came from."""
    _record_mutation(
        "test_203",
        "| `modProfiling.SyncRows` | every existing weight into a Dictionary",
        "| the profiling sync | every existing weight into a Dictionary")


# ===========================================================================
# Q. THE EQUIVALENCE GATE'S STARTING BUNDLE
# ===========================================================================
# The five the authorisation named, plus the ones the corrected vocabulary now
# makes mutable. Each damages the gate on disk and re-runs the executed bundle
# harness against it, because which files arrive where is behaviour.
BUNDLE_CONTROLS = (
    "test_210_the_bundle_contract_is_derived_from_the_bootstrap",
    "test_211_both_bundles_receive_every_required_artifact",
    "test_212_the_two_bundles_are_isolated_and_proved_identical",
    "test_213_a_missing_required_artifact_refuses_before_excel",
    "test_214_no_stale_stage_b_workbook_travels_into_a_bundle",
    "test_215_a_pass_that_did_not_complete_cannot_be_reported_as_PASS",
    "test_216_differ_is_reserved_for_a_comparison_that_actually_ran",
    "test_217_calc_and_calcequiv_cannot_be_emitted_without_two_calculations",
    "test_218_nothing_is_built_when_the_starting_states_disagree",
    "test_219_the_equivalence_field_set_is_unchanged",
    # Reads the gate's snapshot too, so a gate mutation must exercise it: it is
    # the control that says the reserved suffix stays in the comparison.
    "test_257_the_snapshot_compares_the_reserved_suffix_not_just_the_first_n_rows",
)


def _gate_mutation(expected: str, before: str, after: str) -> None:
    """Damage the equivalence gate ON DISK and re-run the bundle harness on it.

    ON DISK, because the bundle harness lifts the gate's functions by AST through
    a subprocess - an in-memory copy is invisible to it, which is the same reason
    the flow and shape harnesses are mutated this way.
    """
    path = conformance.EQUIV_HARNESS
    with path.open(encoding="utf-8", newline="") as handle:
        original = handle.read()
    damaged = original.replace(before.replace("\n", "\r\n"), after.replace("\n", "\r\n"), 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the gate")
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    refused = []
    try:
        conformance._MEMO.pop("equiv_harness", None)
        conformance._MEMO.pop("bundle", None)
        for name in BUNDLE_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(original)
        conformance._MEMO.pop("equiv_harness", None)
        conformance._MEMO.pop("bundle", None)
    assert refused, "the mutation survived every bundle check"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_230_omitting_the_manifest_from_the_bundle_is_rejected() -> None:
    """THE EXACT DEFECT RUN 1 DIED ON. `stage_b_manifest.json` is the first thing
    build_stage_b.ps1 reads from the supplied -BuildDir, and without it both
    passes fail before Excel with a message telling the operator to run the Stage
    A they had just run."""
    _gate_mutation(
        "test_21",
        "    return @(\n"
        "        [pscustomobject]@{ Name = 'stage_b_manifest.json'; Kind = 'file' }\n",
        "    return @(\n")


def test_231_omitting_the_stage_a_workbook_from_the_bundle_is_rejected() -> None:
    """THE OTHER FILE THE BOOTSTRAP RESOLVES AGAINST THE BuildDir. Without it the
    bootstrap throws at line 100 instead of 94 - a different line, the same dead
    run."""
    _gate_mutation(
        "test_21",
        "        [pscustomobject]@{ Name = [string]$Manifest.stage_a_filename; Kind = 'file' }\n",
        "")


def test_232_omitting_the_generated_vba_from_the_bundle_is_rejected() -> None:
    """THE GENERATED MODULES ARE BUILD-DIRECTORY-RELATIVE ON PURPOSE. Resolving
    them against the repository instead is the defect build_stage_b.ps1's own
    comment records: the harness would import modConstants from the real build
    while testing a disposable one."""
    _gate_mutation(
        "test_21",
        "        [pscustomobject]@{ Name = $generatedLeaf; Kind = 'directory' }\n",
        "")


def test_233_sharing_one_mutable_workdir_between_the_passes_is_rejected() -> None:
    """TWO PASSES IN ONE DIRECTORY IS ONE PASS. The Bulk bootstrap would overwrite
    the Endpoints workbook, and the comparison would be of one workbook against
    itself - which would pass, and mean nothing."""
    _gate_mutation(
        "test_212",
        "    $root = Join-Path $WorkDir ('pccm-equivalence-' + $Mode.ToLower() + '-' + $Stamp)\n"
        "    if (Test-Path -LiteralPath $root) {\n"
        "        throw ('the disposable bundle directory ' + $root + ' already exists')\n"
        "    }",
        "    $root = Join-Path $WorkDir ('pccm-equivalence-shared-' + $Stamp)")


def test_234_removing_the_bundle_hash_equality_check_is_rejected() -> None:
    """THE TWO PASSES MUST START FROM THE SAME BYTES. Without the comparison a
    difference in the starting state would be invisible, and every EQUIV line
    afterwards would be about two different workbooks."""
    _gate_mutation(
        "test_212",
        "        if ([string]$Left.Digests[$key] -cne [string]$Right.Digests[$key]) {",
        "        if ($false) {")


def test_235_not_hashing_the_bundle_at_all_is_rejected() -> None:
    """A COMPARISON OVER AN EMPTY DIGEST MAP IS VACUOUS, and would report
    identical every time."""
    _gate_mutation(
        "test_212",
        "            $digests.Add($relative, [string](Get-FileHash -LiteralPath $file.FullName `\n"
        "                -Algorithm SHA256).Hash)",
        "            $null = $relative")


def test_236_restoring_the_PASS_RAISED_vocabulary_is_rejected() -> None:
    """RUN 1 PRINTED `PASS|Endpoints|RAISED` FOR A PASS THAT NEVER BUILT ANYTHING.
    A reader scanning for PASS would have counted two."""
    _gate_mutation(
        "test_215",
        "            Write-Output ('FAIL|' + $mode + '|' + $stage + '|' + $detail)",
        "            Write-Output ('PASS|' + $mode + '|RAISED|' + $detail)")


def test_237_reporting_a_setup_failure_as_differ_is_rejected() -> None:
    """`differ` MEANS THE TWO FIXTURES ARE NOT EQUIVALENT. Run 1 printed it after a
    setup failure, which is a claim it had no evidence for."""
    _gate_mutation(
        "test_216",
        "    Write-Output ('EQUIV|<not evaluated>|invalid|comparison was not executed: ' + $why)",
        "    Write-Output ('EQUIV|<no comparison>|differ|' + $why)")


def test_238_emitting_calcequiv_without_two_calculations_is_rejected() -> None:
    """NO PLACEHOLDER VERDICT ON A CALCULATION THAT NEVER RAN. Moving CALCEQUIV
    outside the both-completed branch would have printed one after run 1."""
    _gate_mutation(
        "test_217",
        "$completed = @($passes.Keys)\n",
        "$completed = @($passes.Keys)\n"
        "Write-Output 'CALCEQUIV|match|the production calculation fingerprint is identical'\n")


def test_239_completing_a_pass_without_a_fingerprint_is_rejected() -> None:
    """COMPLETED MEANS BOTH HALVES ARRIVED. A pass that snapshotted and never
    calculated would be compared, and CALCEQUIV would compare two empty strings
    and call them equal."""
    _gate_mutation(
        "test_215",
        "    if ([string]::IsNullOrWhiteSpace($calcFingerprint)) {",
        "    if ($false) {")


def test_240_building_the_passes_before_proving_identity_is_rejected() -> None:
    """REFUSING BEFORE EXCEL IS THE POINT. Running the passes regardless would burn
    two Stage-B bootstraps and two Excel sessions to produce a comparison that was
    never valid."""
    _gate_mutation(
        "test_218",
        "$passes = @{}\nif (-not $setupFailed) {\n",
        "$passes = @{}\nif ($true) {\n")


def test_241_carrying_a_stale_stage_b_workbook_into_a_bundle_is_rejected() -> None:
    """THE OUTPUT IS NOT AN INPUT. A stale .xlsm in the bundle would be opened by
    the pass instead of the one its own bootstrap just built."""
    _gate_mutation(
        "test_214",
        "        [pscustomobject]@{ Name = [string]$Manifest.stage_a_filename; Kind = 'file' }",
        "        [pscustomobject]@{ Name = [string]$Manifest.stage_a_filename; Kind = 'file' }\n"
        "        [pscustomobject]@{ Name = [string]$Manifest.stage_b_filename; Kind = 'file' }")


def test_242_narrowing_the_equivalence_field_set_is_rejected() -> None:
    """A BROKEN SETUP IS NO REASON TO COMPARE LESS. Dropping a family would make
    the gate pass on a fixture that differed in exactly that family."""
    _gate_mutation(
        "test_219",
        "    $state.Add('fingerprint.calculation_inputs', [string]$Excel.Run('PCCM_CurrentInputFingerprint'))",
        "    $null = 'calculation inputs not compared'")


def test_243_softening_the_run_1_record_is_rejected() -> None:
    """RUN 1 WAS INVALID, NOT A DIFFERENCE. A record that let it read as a semantic
    result would licence a Bulk baseline on evidence that does not exist."""
    _record_mutation(
        "test_220",
        "**No semantic equivalence claim may be made from it.**",
        "The comparison was inconclusive.")


# ===========================================================================
# R. RESERVED CAPACITY vs SEMANTIC COUNT
# ===========================================================================
# Each of these damages the runner ON DISK and re-runs the EXECUTED reserved-rows
# harness against it, because "how many rows and how many Adds" is a count.
RESERVED_CONTROLS = (
    "test_250_productions_own_rule_is_grow_only_when_capacity_runs_out",
    "test_251_reserved_capacity_is_kept_and_growth_happens_only_when_needed",
    "test_252_identifiers_stop_at_the_semantic_count_and_the_counter_matches",
    "test_253_columns_no_driver_fills_are_genuinely_blank",
    "test_254_the_grower_takes_a_floor_and_never_shrinks",
    "test_255_the_reserved_suffix_is_proved_blank_before_production_syncs",
    "test_256_profiling_geometry_is_still_productions_alone",
    "test_257_the_snapshot_compares_the_reserved_suffix_not_just_the_first_n_rows",
    "test_195_growing_a_register_is_bounded_proved_and_never_destructive",
)


def _reserved_mutation(expected: str, before: str, after: str) -> None:
    """Damage the runner ON DISK and re-run the executed reserved-rows harness."""
    path = conformance.RUNNER
    with path.open(encoding="utf-8", newline="") as handle:
        original = handle.read()
    damaged = original.replace(before.replace("\n", "\r\n"), after.replace("\n", "\r\n"), 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:60]!r} is no longer in the runner")
    with path.open("w", encoding="utf-8", newline="") as handle:
        handle.write(damaged)
    refused = []
    try:
        conformance._MEMO.pop("runner", None)
        conformance._MEMO.pop("code", None)
        conformance._MEMO.pop("reserved", None)
        for name in RESERVED_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(original)
        conformance._MEMO.pop("runner", None)
        conformance._MEMO.pop("code", None)
        conformance._MEMO.pop("reserved", None)
    assert refused, "the mutation survived every reserved-row check"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_250_shrinking_the_register_to_the_driver_count_is_rejected() -> None:
    """THE DEFECT EQUIVALENCE RUN 2 DIED ON, in its other form. Deleting down to
    twelve rows would delete rows production would have kept blank, and the two
    fixtures would differ physically - which the snapshot compares and rightly
    reports."""
    _reserved_mutation(
        "test_25",
        "    if ($current -ge $MinimumRows) { return $current }",
        "    if ($current -gt $MinimumRows) {\n"
        "        throw ($TableName + ' holds ' + [string]$current + ' rows, not ' +\n"
        "               [string]$MinimumRows)\n"
        "    }\n"
        "    if ($current -eq $MinimumRows) { return $current }")


def test_251_growing_past_the_semantic_count_is_rejected() -> None:
    """EXACTLY THE SHORTFALL. Overshooting would leave rows production never
    created, and the physical counts would diverge from the reference."""
    _reserved_mutation(
        "test_251",
        "        for ($i = $current; $i -lt $MinimumRows; $i++) {",
        "        for ($i = $current; $i -le $MinimumRows; $i++) {")


def test_252_growing_when_capacity_already_suffices_is_rejected() -> None:
    """TWELVE DRIVERS INTO TWENTY-FIVE RESERVED ROWS MUST MAKE ZERO ADDS. Growing
    anyway is a structural mutation nobody needed, and it would put thirteen extra
    rows on the register."""
    _reserved_mutation(
        "test_251",
        "    if ($current -ge $MinimumRows) { return $current }",
        "    if ($current -ge $MinimumRows) { $MinimumRows = $current + 1 }")


def test_253_issuing_an_identifier_past_the_semantic_count_is_rejected() -> None:
    """CL-013 IN A TWELVE-DRIVER SMALL FIXTURE. Production issues identifiers in
    sequence and stops; a thirteenth would disagree with the counter and with the
    model."""
    _reserved_mutation(
        "test_252",
        "    $block = New-Object 'object[,]' $drivers.Count, $columns.Count\n"
        "    $ids = @()\n"
        "    for ($index = 0; $index -lt $drivers.Count; $index++) {",
        "    $block = New-Object 'object[,]' $drivers.Count, $columns.Count\n"
        "    $ids = @(Get-BenchmarkPermanentId -Counter $Counter -Sequence ($drivers.Count + 1))\n"
        "    for ($index = 0; $index -lt $drivers.Count; $index++) {")


def test_254_leaving_the_counter_short_of_the_semantic_count_is_rejected() -> None:
    """THE COUNTER IS PRODUCTION'S RECORD OF EVERY IDENTIFIER EVER ISSUED. Short of
    N it disagrees with the rows present, and modDrivers.HighestIssued exists to
    catch that."""
    _reserved_mutation(
        "test_252",
        "        CounterValue = [double]$drivers.Count",
        "        CounterValue = [double]($drivers.Count - 1)")


def test_255_populating_a_column_no_driver_fills_is_rejected() -> None:
    """`category` AND `uom` ARE BLANK IN THE ENDPOINT-BUILT REGISTER. Writing ''
    there makes them populated, and the two bodies then differ in every row."""
    _reserved_mutation(
        "test_253",
        "            else { $block[$index, $c] = $null }",
        "            else { $block[$index, $c] = '' }")


def test_256_populating_a_reserved_row_is_rejected() -> None:
    """A VALUE IN A ROW PRODUCTION NEVER KEYED IS AN ORPHAN, and AddDriver refuses
    to mutate over one - so a fixture that left one would poison every later
    production command."""
    _reserved_mutation(
        "test_255",
        "        for ($row = $expected.Count; $row -lt $body.Count; $row++) {",
        "        for ($row = $body.Count; $row -lt $body.Count; $row++) {")


def test_257_dropping_the_reserved_suffix_check_is_rejected() -> None:
    """PRODUCTION IS ABOUT TO SYNCHRONISE BOTH GRIDS FROM THAT REGISTER. An orphan
    below the semantic rows would be carried into the grids, and everything
    downstream would agree with itself."""
    _reserved_mutation(
        "test_255",
        "                if ([string]$value -ne '') {",
        "                if ($false) {")


def test_258_resizing_a_profiling_grid_from_the_builder_is_rejected() -> None:
    """THE GRIDS ARE PRODUCTION'S. SyncRows owns their row count, and a builder
    that resized one would be reimplementing it."""
    # RE-ANCHORED for the same reason as test_194.
    _reserved_mutation(
        "test_256",
        "        $prepared = New-BenchmarkWeightBlock -Workbook $Workbook -Grid $grid `",
        "        $null = Set-BenchmarkRegisterRowCount -Workbook $Workbook `\n"
        "            -SheetName $grid.sheet -TableName $grid.table_name `\n"
        "            -MinimumRows @($pair.drivers).Count\n"
        "        $prepared = New-BenchmarkWeightBlock -Workbook $Workbook -Grid $grid `")


def test_259_counting_blank_grid_rows_as_keyed_rows_is_rejected() -> None:
    """A GRID WITH A BLANK RESERVED SUFFIX MUST STILL MATCH THE DRIVER COUNT.
    Counting the blanks would make a twelve-driver fixture look like twenty-five
    and the weight block would refuse for the wrong reason."""
    _reserved_mutation(
        "test_256",
        "        if ([string]::IsNullOrWhiteSpace($key)) { continue }",
        "        if ($false) { continue }")


def test_260_weakening_the_snapshot_to_ignore_reserved_rows_is_rejected() -> None:
    """IF ENDPOINTS HAS 25 PHYSICAL ROWS AND BULK HAS 12, THAT MUST BE A REAL
    DIFFER. Narrowing the body comparison to the first N rows would hide exactly
    the defect run 2 surfaced."""
    _gate_mutation(
        "test_257",
        "        $lines = @()\n"
        "        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $register.sheet `\n"
        "                -TableName $register.table_name)) {\n"
        "            $lines += ((@($row) -join '|'))\n"
        "        }",
        "        $lines = @()\n"
        "        foreach ($row in @(Get-TableBody -Workbook $Workbook -SheetName $register.sheet `\n"
        "                -TableName $register.table_name)) {\n"
        "            if ([string]$row[0] -ne '') { $lines += ((@($row) -join '|')) }\n"
        "        }")


def test_261_softening_the_run_2_record_is_rejected() -> None:
    """RUN 2 WAS INVALID, NOT A DIFFERENCE. A record that let it read as a semantic
    result would licence a Bulk baseline on evidence that does not exist."""
    _record_mutation(
        "test_258",
        "**must not be read as a fixture DIFFER**",
        "may be read as a fixture difference")


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


# ===========================================================================
# S. THE RECTANGULAR FIXTURE BLOCKS
# ===========================================================================
# THE RANK CONTROLS RUN A HARNESS AGAINST THE FILE ON DISK, so an in-memory mutation
# would never reach them. These write the runner, re-run the executed rank proof, and
# restore it in a finally - the same shape the reserved-row battery uses.
RANK_CONTROLS = (
    "test_261_every_fixture_block_is_a_rank_two_array_of_the_right_shape",
    "test_262_the_weight_block_record_reports_the_geometry_it_built",
    "test_263_a_matrix_is_never_the_thing_a_function_emits",
    "test_260_production_is_byte_identical_and_the_timed_path_did_not_move",
)


def _rank_mutation(expected: str, before: str, after: str) -> None:
    """Damage the runner ON DISK and re-run the EXECUTED rank proof against it."""
    path = conformance.RUNNER
    with path.open(encoding="utf-8", newline="") as handle:
        original = handle.read()
    damaged = original.replace(before.replace("\n", "\r\n"), after.replace("\n", "\r\n"), 1)
    if damaged == original:
        raise RuntimeError(
            f"the mutation changed nothing: {before[:70]!r} is no longer in the runner")
    refused = []
    try:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(damaged)
        conformance._MEMO.pop("runner", None)
        conformance._MEMO.pop("code", None)
        conformance._MEMO_ANY.pop("rank", None)
        for name in RANK_CONTROLS:
            try:
                getattr(conformance, name)()
            except BaseException:  # noqa: BLE001 - any refusal counts
                refused.append(name)
    finally:
        with path.open("w", encoding="utf-8", newline="") as handle:
            handle.write(original)
        conformance._MEMO.pop("runner", None)
        conformance._MEMO.pop("code", None)
        conformance._MEMO_ANY.pop("rank", None)
    assert refused, "the mutation survived every rank check"
    assert any(name.startswith(expected) for name in refused), (expected, refused)


def test_262_flattening_the_profiling_block_is_refused() -> None:
    """THE DEFECT WINDOWS FOUND, PLANTED. `return $block` on a rank-2 object[,] is
    EMITTED, and the pipeline enumerates a multidimensional array into its elements -
    so the caller's assignment gets an Object[] and Excel is handed a rank-1 array."""
    _rank_mutation(
        "test_26",
        "    return [pscustomobject]@{\n"
        "        Block   = $block\n"
        "        Rows    = $rows.Count\n"
        "        Columns = $Years\n"
        "        Keys    = $rows\n"
        "    }",
        "    return $block")


def test_263_returning_the_block_through_an_enumerating_form_is_refused() -> None:
    """SAME DEFECT IN ANOTHER COSTUME. Write-Output enumerates exactly as return
    does; the matrix must not be the thing the function emits by ANY route."""
    _rank_mutation(
        "test_26",
        "    return [pscustomobject]@{\n"
        "        Block   = $block\n"
        "        Rows    = $rows.Count\n"
        "        Columns = $Years\n"
        "        Keys    = $rows\n"
        "    }",
        "    Write-Output $block")


def test_264_transposing_the_block_dimensions_is_refused() -> None:
    """ROWS x YEARS, NOT YEARS x ROWS. A transposed block is still a perfectly
    rectangular array, so only its DIMENSIONS give it away - and it would write every
    driver's weights down the wrong axis."""
    _rank_mutation(
        "test_26",
        "    $block = New-Object 'object[,]' $rows.Count, $Years",
        "    $block = New-Object 'object[,]' $Years, $rows.Count")


def test_265_disabling_the_rank_guard_is_refused() -> None:
    """THE GUARD IS WHAT CAUGHT THIS ON WINDOWS, before a flattened array reached
    Excel and wrote something shaped like an answer."""
    _rank_mutation(
        "test_26",
        "    if ($Block.Rank -ne 2) {",
        "    if ($false) {")


def test_266_weakening_the_destination_dimension_check_is_refused() -> None:
    """Set-BenchmarkRangeBlock RESIZES THE ANCHOR TO THE BLOCK'S OWN DIMENSIONS, so a
    block of the wrong shape writes a perfectly consistent rectangle in the wrong
    place. The caller asserting what it ASKED for is what stops that."""
    _rank_mutation(
        "test_262",
        "        if ([int]$prepared.Rows -ne @($pair.drivers).Count) {",
        "        if ($false) {")


def test_267_weakening_the_column_dimension_check_is_refused() -> None:
    """THE OTHER AXIS. A block with the wrong number of year columns would write
    weights into the wrong years without changing its rank."""
    _rank_mutation(
        "test_262",
        "        if ([int]$prepared.Columns -ne $years) {",
        "        if ($false) {")


def test_268_weakening_the_register_geometry_check_is_refused() -> None:
    """THE SAME PROOF ON THE BLOCK THAT NEVER COLLAPSED. It travels as a record
    property and always has - the caller asserting it is what keeps that true rather
    than lucky."""
    _rank_mutation(
        "test_262",
        "        if ([int]$prepared.Block.GetLength(0) -ne @($pair.drivers).Count) {",
        "        if ($false) {")


def test_269_flattening_the_register_block_is_refused() -> None:
    """THE REGISTER BUILDER IS THE PATTERN THE FIX COPIED. Emitting its matrix would
    break the one path Windows has always written correctly."""
    _rank_mutation(
        "test_26",
        "    return [pscustomobject]@{\n"
        "        Block = $block\n"
        "        Ids = $ids",
        "    $null = $ids\n"
        "    return $block\n"
        "    $unreachable = [pscustomobject]@{\n"
        "        Block = $block\n"
        "        Ids = $ids")


if __name__ == "__main__":
    raise SystemExit(pytest.main([__file__, "-q"]))
