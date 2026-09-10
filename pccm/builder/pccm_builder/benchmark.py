"""P10-4A: the performance benchmark plan.

WHAT THIS FILE IS. The single authority for WHAT is benchmarked - the three
model sizes, the iteration matrix, the operations, the timing method, and the
policy by which a later run is compared with the baseline this one creates. It
is emitted as `build/phase10_benchmark_plan.json`, and the Windows runner READS
that file rather than declaring a matrix of its own.

WHY IT IS NOT DECLARED IN POWERSHELL. A matrix that lived in the runner could
not be checked on Linux without parsing PowerShell for its meaning, and the two
things that must never disagree - what the contract settled and what the machine
executes - would be a regex apart. Emitted as data, the plan is verified exactly
here and executed literally there.

WHAT THIS FILE MEASURES NOTHING WITH. It runs no benchmark, opens no workbook
and contains no timing code. It is a list of what to do, in what order, and what
evidence to keep.

THE ONE THING IT DELIBERATELY DOES NOT CARRY IS A PASS MARK. This is the first
delivery-oriented baseline, so there is no absolute threshold to compare
against and inventing one before the target machine has been observed would
make the first run fail or pass for a reason nobody could defend. The plan
carries the RATIOS that apply to later comparable runs, and nothing else.
"""
from __future__ import annotations

import json
import math
from pathlib import Path
from typing import Any

__all__ = [
    "BENCHMARK_SCHEMA_VERSION",
    "HARNESS_VERSION",
    "BASELINE_ID",
    "build_benchmark_plan",
    "cost_line_count",
    "emit_benchmark_plan",
]

# The shape of the emitted plan and of the result artifact. Bumped only when a
# consumer would have to change; the runner refuses a plan it does not know.
BENCHMARK_SCHEMA_VERSION = 1
HARNESS_VERSION = "1.0.0"

# WHAT A LATER RUN COMPARES ITSELF AGAINST. A result artifact carries this id so
# two runs can only be compared when they measured the same declared plan.
BASELINE_ID = "PCCM-P10-BASELINE-1"

# ---------------------------------------------------------------------------
# THE MODEL SIZES - the settled Phase-10 matrix, not a new one
# ---------------------------------------------------------------------------
# `docs/phase10_step1_contract.md` section 10: "Small = 20 drivers / 10 years .
# Medium = 100 / 25 . Large = 300 / 40." A control reads that sentence out of
# the record and compares it with these numbers, so the two cannot drift.
SCENARIOS: tuple[dict[str, Any], ...] = (
    {"id": "PERF-SMALL", "title": "Small", "drivers": 20, "years": 10},
    {"id": "PERF-MEDIUM", "title": "Medium", "drivers": 100, "years": 25},
    {"id": "PERF-LARGE", "title": "Large", "drivers": 300, "years": 40},
)

# THE COST/RISK SPLIT IS NOT INVENTED HERE. `bootstrap/windows/
# phase7_timing_scenarios.ps1` already settled it for the accepted Phase-7
# measurement - "60% Cost Lines, the remainder Risks: 20 -> 12/8, 100 -> 60/40,
# 300 -> 180/120" - and the same rule is used, so the two harnesses build
# comparably shaped models and the Phase-7 numbers stay readable as context.
COST_LINE_SHARE = 0.6

# ---------------------------------------------------------------------------
# THE ITERATION MATRIX - a practical matrix, not a Cartesian product
# ---------------------------------------------------------------------------
ITERATION_MATRIX: dict[str, tuple[int, ...]] = {
    "PERF-SMALL": (10_000, 50_000, 100_000),
    "PERF-MEDIUM": (10_000, 50_000, 100_000),
    "PERF-LARGE": (10_000, 50_000),
}

# AND THE COMBINATION THAT IS FORBIDDEN RATHER THAN MERELY ABSENT. The contract
# caps Large at 50,000 for Simulation, Sensitivity and annual replay. Recording
# it as a refusal with a reason is the difference between "we chose not to" and
# "the evidence is missing".
FORBIDDEN_COMBINATIONS: tuple[dict[str, Any], ...] = (
    {
        "scenario": "PERF-LARGE",
        "iterations": 100_000,
        "operations": ("simulation", "sensitivity", "annual"),
        "reason": ("not required / impractical: the accepted Phase-10 contract caps "
                   "Large at 50,000 iterations for Simulation, Sensitivity and annual "
                   "replay"),
    },
)

# ---------------------------------------------------------------------------
# THE OPERATIONS
# ---------------------------------------------------------------------------
# `kind` is the load-bearing field. A "command" is a real production endpoint,
# invoked by name through Application.Run and by nothing else; "recalculation"
# is ordinary Excel work with no endpoint at all, measured to establish what the
# Calculate figure sits on top of.
OPERATIONS: tuple[dict[str, Any], ...] = (
    {
        "key": "calculate",
        "label": "Calculate",
        "kind": "command",
        "endpoint": "PCCM_Calculate",
        "iteration_dependent": False,
        "evidence": ("automation_result", "calculation_status", "calculation_fingerprint"),
    },
    {
        "key": "recalculation",
        "label": "Workbook recalculation",
        "kind": "recalculation",
        "endpoint": None,
        "mechanism": "Application.CalculateFull",
        "iteration_dependent": False,
        "evidence": ("calculation_status",),
    },
    {
        "key": "simulation",
        "label": "Run Simulation",
        "kind": "command",
        "endpoint": "PCCM_RunSimulation",
        "iteration_dependent": True,
        "evidence": ("automation_result", "simulation_status", "published_iterations",
                     "active_bank", "run_id"),
    },
    {
        "key": "sensitivity",
        "label": "Run Sensitivity",
        "kind": "command",
        "endpoint": "PCCM_RunSensitivity",
        "iteration_dependent": True,
        "evidence": ("automation_result", "simulation_status", "published_iterations",
                     "sensitivity_record_count", "sensitivity_ranked",
                     "sensitivity_zero_variance"),
    },
    {
        "key": "annual",
        "label": "Run Annual Cash Flow",
        "kind": "command",
        "endpoint": "PCCM_RunAnnualStochastic",
        "iteration_dependent": True,
        "evidence": ("automation_result", "annual_distribution_state",
                     "annual_profile_state", "annual_profile_px", "annual_year_count"),
    },
)

# ---------------------------------------------------------------------------
# THE TIMING METHOD
# ---------------------------------------------------------------------------
# COLD AND WARM ARE DEFINED HERE AND NOWHERE ELSE, because "cold" is the word
# most easily used to mean whatever the number happened to be.
#
#   COLD  the FIRST execution of this operation in this Excel session against
#         this scenario and iteration count. It pays whatever first-touch cost
#         exists - the first entry into that code path, the first allocation of
#         its arrays, the first touch of the sheets it writes. It is measured
#         exactly as a warm run is: the clock surrounds one call and nothing
#         else. It is reported SEPARATELY and never averaged into anything.
#
#   WARM  each of the three executions immediately following the cold one, with
#         no other operation between them and no reopen of the workbook.
#
# WHAT COLD IS NOT: it is not the time to start Excel, not the time to open the
# workbook, and not the time to build the scenario. Those are setup, they are
# timed separately, and they are never part of an operation's elapsed time.
TIMING = {
    "cold_runs": 1,
    "warm_runs": 3,
    "comparison_statistic": "median_of_warm_runs",
    "aggregation": ("the median of the three warm runs. The mean is NOT used: it "
                    "would average away exactly the outlier a scheduler or a "
                    "background process introduces, and that outlier is "
                    "information about the machine."),
    "timing_source": "System.Diagnostics.Stopwatch",
    "timing_source_note": ("high-resolution and monotonic; it is not affected by a "
                           "wall-clock adjustment during the run"),
    "cold_definition": ("the first execution of the operation in this Excel session "
                        "against this scenario and iteration count"),
    "warm_definition": ("each of the three executions immediately after the cold one, "
                        "with no intervening operation and no workbook reopen"),
    "excluded_from_elapsed": ("Excel startup", "workbook open", "Stage-B bootstrap",
                              "scenario fixture construction", "evidence reads",
                              "report writing"),
}

# ---------------------------------------------------------------------------
# BASELINE FIRST
# ---------------------------------------------------------------------------
REGRESSION_POLICY = {
    "baseline_first": True,
    "baseline_id": BASELINE_ID,
    "absolute_threshold_seconds": None,
    "absolute_threshold_note": ("NONE. No absolute pass/fail threshold is contracted "
                                "before a baseline exists on the real target machine. "
                                "The first run RECORDS; it does not judge, and a slow "
                                "first measurement is a fact rather than a failure."),
    "investigate_ratio": 1.5,
    "blocking_ratio": 2.0,
    "applies_to": ("a later comparable run: the same plan, the same scenario, the "
                   "same operation, the same iteration count, compared warm median "
                   "against warm median"),
    "rules": (
        "at or below 1.5x the accepted baseline: no automatic regression finding",
        "above 1.5x: investigate",
        "above 2.0x: blocking regression unless explained and explicitly accepted",
    ),
}

# ---------------------------------------------------------------------------
# THE PHASE-7 NUMBERS ARE CONTEXT
# ---------------------------------------------------------------------------
# They are the operator's report of a different measurement, on a machine whose
# record was never settled, against code that has since changed. Carried so a
# reader can see roughly where the earlier evidence sat, and marked so nobody
# can mistake them for a bar to clear.
HISTORICAL_CONTEXT = {
    "status": "historical context only - NOT a threshold and NOT a baseline",
    "source": "docs/phase7_closure.md, the operator's report",
    "not_a_threshold_because": (
        "the code and the workbook have changed since",
        "the timing methodology may differ from this plan's cold/warm definition",
        "the machine and environment record was not settled when they were taken",
    ),
    "observations": (
        {"operation": "sensitivity", "drivers": 20, "iterations": 10_000,
         "seconds": "6.4-7.6"},
        {"operation": "sensitivity", "drivers": 100, "iterations": 10_000,
         "seconds": "31.4"},
        {"operation": "sensitivity", "drivers": 300, "iterations": 10_000,
         "seconds": "105.4"},
    ),
}

# ---------------------------------------------------------------------------
# WHAT THE MACHINE MUST RECORD ABOUT ITSELF
# ---------------------------------------------------------------------------
# A timing without this is a number with no interpretation. Every field is
# required to be present in the emitted result - an unavailable one is recorded
# as unavailable, never omitted, so a missing field always means the harness did
# not ask rather than the machine did not say.
ENVIRONMENT_FIELDS: tuple[str, ...] = (
    "captured_at_utc",
    "host_name",
    "powershell_version",
    "windows_edition",
    "windows_version",
    "windows_build",
    "cpu_model",
    "logical_processors",
    "installed_ram_gb",
    "excel_version",
    "excel_build",
    "excel_bitness",
    "excel_calculation_mode",
    "other_workbooks_open",
    "workbook_path",
    "workbook_location_type",
    "workbook_location_note",
    "repository_path",
    "repository_location_type",
    "onedrive_roots",
    "git_branch",
    "git_commit",
    "git_worktree_clean",
    "model_version",
    "builder_version",
    "build_phase",
    "harness_version",
    "schema_version",
)

# The location words. "onedrive" and "synced" are recorded as observations about
# the PATH; no claim is made here or in the runner about what either does to a
# measurement.
LOCATION_TYPES: tuple[str, ...] = ("local", "onedrive", "synced", "unknown")

# ---------------------------------------------------------------------------
# CORRECTNESS GATES
# ---------------------------------------------------------------------------
# A FAST FAILURE IS NOT A FAST OPERATION. Every timed execution carries the
# evidence that it did the work, and a sample that fails any gate is marked
# invalid and excluded from every statistic rather than quietly reported.
CORRECTNESS_GATES = {
    "every_execution": (
        "the automation result announces success for a command operation",
        "no refusal or error text was announced",
    ),
    "iteration_dependent": (
        "the published iteration count equals the requested iteration count",
    ),
    "invalid_sample_policy": ("a failed, refused or unverified execution is NOT a "
                              "performance sample. It is recorded with its refusal "
                              "text, the scenario is marked invalid, and no median "
                              "is computed from it."),
    "median_requires": "three valid warm samples",
}


def cost_line_count(drivers: int) -> int:
    """The accepted 60/40 split: 20 -> 12/8, 100 -> 60/40, 300 -> 180/120."""
    return int(math.ceil(float(drivers) * COST_LINE_SHARE))


def _scenario(entry: dict[str, Any]) -> dict[str, Any]:
    drivers = int(entry["drivers"])
    cost_lines = cost_line_count(drivers)
    risks = drivers - cost_lines
    if risks < 1:
        raise ValueError(f"{entry['id']} has no Risks; the split is degenerate")
    return {
        "id": str(entry["id"]),
        "title": str(entry["title"]),
        "drivers": drivers,
        "cost_lines": cost_lines,
        "risks": risks,
        "years": int(entry["years"]),
        "iterations": list(ITERATION_MATRIX[str(entry["id"])]),
    }


def _runs(scenario: dict[str, Any]) -> list[dict[str, Any]]:
    """The execution matrix for one scenario, in the order it must be run.

    ORDER IS PART OF THE PLAN. Calculate establishes the deterministic basis
    the rest depends on; a simulation must be CURRENT before Sensitivity or the
    annual replay will do anything at all. Running them in any other order would
    measure a refusal.
    """
    out: list[dict[str, Any]] = []
    for operation in OPERATIONS:
        if operation["iteration_dependent"]:
            continue
        # ITERATION-INDEPENDENT, SO MEASURED ONCE PER SCENARIO. Repeating
        # Calculate at 10k, 50k and 100k would produce three measurements of the
        # same work and make the matrix look richer than the evidence is.
        out.append({
            "scenario": scenario["id"],
            "operation": str(operation["key"]),
            "iterations": None,
            "cold_runs": TIMING["cold_runs"],
            "warm_runs": TIMING["warm_runs"],
        })
    for iterations in scenario["iterations"]:
        for operation in OPERATIONS:
            if not operation["iteration_dependent"]:
                continue
            out.append({
                "scenario": scenario["id"],
                "operation": str(operation["key"]),
                "iterations": int(iterations),
                "cold_runs": TIMING["cold_runs"],
                "warm_runs": TIMING["warm_runs"],
            })
    return out


def build_benchmark_plan() -> dict[str, Any]:
    """The whole plan, expanded, with every rule it was built under."""
    scenarios = [_scenario(entry) for entry in SCENARIOS]
    runs: list[dict[str, Any]] = []
    for scenario in scenarios:
        runs.extend(_runs(scenario))

    forbidden = [
        {"scenario": entry["scenario"], "iterations": entry["iterations"],
         "operations": list(entry["operations"]), "reason": entry["reason"]}
        for entry in FORBIDDEN_COMBINATIONS
    ]

    plan = {
        "schema_version": BENCHMARK_SCHEMA_VERSION,
        "harness_version": HARNESS_VERSION,
        "baseline_id": BASELINE_ID,
        "purpose": ("the first delivery-oriented performance baseline. It records; "
                    "it does not judge."),
        "cost_line_share": COST_LINE_SHARE,
        "scenarios": scenarios,
        "operations": [
            {key: (list(value) if isinstance(value, tuple) else value)
             for key, value in operation.items()}
            for operation in OPERATIONS
        ],
        "runs": runs,
        "forbidden": forbidden,
        "timing": {key: (list(value) if isinstance(value, tuple) else value)
                   for key, value in TIMING.items()},
        "correctness_gates": {key: (list(value) if isinstance(value, tuple) else value)
                              for key, value in CORRECTNESS_GATES.items()},
        "regression_policy": {key: (list(value) if isinstance(value, tuple) else value)
                              for key, value in REGRESSION_POLICY.items()},
        "historical_context": {
            "status": HISTORICAL_CONTEXT["status"],
            "source": HISTORICAL_CONTEXT["source"],
            "not_a_threshold_because": list(HISTORICAL_CONTEXT["not_a_threshold_because"]),
            "observations": [dict(entry) for entry in HISTORICAL_CONTEXT["observations"]],
        },
        "environment_fields": list(ENVIRONMENT_FIELDS),
        "location_types": list(LOCATION_TYPES),
    }
    _assert_plan_is_coherent(plan)
    return plan


def _assert_plan_is_coherent(plan: dict[str, Any]) -> None:
    """The plan proves its own rules before anyone is asked to execute it."""
    forbidden = {(entry["scenario"], int(entry["iterations"]))
                 for entry in plan["forbidden"]}
    for run in plan["runs"]:
        if run["iterations"] is None:
            continue
        key = (run["scenario"], int(run["iterations"]))
        if key in forbidden:
            raise ValueError(f"the plan expands a forbidden combination: {key}")

    independent = {str(operation["key"]) for operation in OPERATIONS
                   if not operation["iteration_dependent"]}
    for scenario in plan["scenarios"]:
        for key in independent:
            matching = [run for run in plan["runs"]
                        if run["scenario"] == scenario["id"] and run["operation"] == key]
            if len(matching) != 1:
                raise ValueError(
                    f"{key} is expanded {len(matching)} times for {scenario['id']}; an "
                    "iteration-independent operation is measured once per scenario")
            if matching[0]["iterations"] is not None:
                raise ValueError(f"{key} carries an iteration count")


def emit_benchmark_plan(path: Path) -> dict[str, Any]:
    plan = build_benchmark_plan()
    path.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    return plan
