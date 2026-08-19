"""The Phase-5 acceptance corpus, as plain data.

`build/phase5_cases.json` is TEST DATA. It is what a later Windows/VBA harness
reads so it can reproduce, on real Excel, the evidence this repository has already
proven in Python. It is not a calculation engine, it is not a second definition of
the mathematics, and nothing in the workbook ever reads it.

--------------------------------------------------------------------------------
THE EVIDENCE CHAIN, AND WHY THIS FILE DOES NOT BREAK IT
--------------------------------------------------------------------------------
    hand-derived literal  ->  Python oracle  ->  phase5_cases.json  ->  later VBA

This module occupies the third arrow: it defines each case as INPUT DATA and asks
the accepted oracle for the expected output. That is legitimate only because the
Step-3 test suite independently re-derives every emitted number - from the plan's
own section 23 literals, or by exact `Fraction` arithmetic written separately in
the test - and fails if the two disagree. A test that only compared this file's
output with the oracle that produced it would prove nothing at all.

--------------------------------------------------------------------------------
WHAT IS NOT CLAIMED
--------------------------------------------------------------------------------
Plan cases 32, 33, 34 and 37 are workbook-state and rollback behaviours. They
cannot be evidenced by a pure function, so they are emitted as `runtime_only` with
no expected numbers. Python does not pretend to prove them.
"""

from __future__ import annotations

from typing import Any

from . import calc_fingerprint as fp
from .calc_loader import CalcContract
from .calc_numeric import MAX_DOUBLE, CalculationRefusal
from .calc_oracle import (
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    FxRow,
    RiskDriver,
    Tolerances,
    calculate,
)

SCHEMA_VERSION = 1

# The shared fixture of plan section 23: "Triangular Min 80 / ML 100 / Max 150,
# Quantity 10, SAR, one profile at 100%" unless a case overrides it.
_BASE_COST: dict[str, Any] = {
    "permanent_id": "CL-001",
    "distribution": "Triangular",
    "currency": "SAR",
    "inflation_profile": "Standard",
    "min_value": 80,
    "most_likely": 100,
    "max_value": 150,
    "profile_weights": [1.0],
    "quantity": 10,
}
_BASE_RISK: dict[str, Any] = {
    "permanent_id": "R-001",
    "distribution": "Triangular",
    "currency": "SAR",
    "inflation_profile": "Standard",
    "min_value": 100,
    "most_likely": 200,
    "max_value": 450,
    "profile_weights": [1.0],
    "probability": 0.3,
}


def _cost(**overrides: Any) -> dict[str, Any]:
    return {**_BASE_COST, **overrides}


def _risk(**overrides: Any) -> dict[str, Any]:
    return {**_BASE_RISK, **overrides}


def _model(
    base_year: int = 2026,
    start_year: int = 2026,
    duration: int = 1,
    discount_rate: Any = 0.10,
    fx: list[dict[str, Any]] | None = None,
    inflation: dict[str, dict[str, Any]] | None = None,
    cost_lines: list[dict[str, Any]] | None = None,
    risks: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    return {
        "timeline": {
            "base_year": base_year, "start_year": start_year, "duration": duration
        },
        "discount_rate": discount_rate,
        "fx": fx if fx is not None else [{"currency": "SAR", "rate": 1}],
        "inflation": inflation if inflation is not None else {"Standard": {}},
        "cost_lines": cost_lines if cost_lines is not None else [],
        "risks": risks if risks is not None else [],
    }


def _three_year(rate: Any = 0.05) -> dict[str, dict[str, Any]]:
    return {"Standard": {"2027": rate, "2028": rate, "2029": rate}}


_PROFILE_3 = [0.2, 0.5, 0.3]

_SUBNORMAL = 5e-324

# What `kind` tells a later harness to assert:
#   analytical   - the model calculates; compare every emitted number
#   refusal      - the model must be refused; compare the refusal class
#   statistics   - a numerical-helper case with no model
#   fingerprint  - the evidence lives in the `fingerprint` section
#   runtime_only - workbook state or rollback; Python proves nothing here
ANALYTICAL = "analytical"
REFUSAL = "refusal"
STATISTICS = "statistics"
FINGERPRINT = "fingerprint"
RUNTIME_ONLY = "runtime_only"


# ---------------------------------------------------------------------------
# The case matrix of plan section 23
# ---------------------------------------------------------------------------
CASES: tuple[dict[str, Any], ...] = (
    {"id": 1, "kind": ANALYTICAL,
     "title": "SAR, no inflation, one project year",
     "model": _model(cost_lines=[_cost()])},
    {"id": 2, "kind": ANALYTICAL,
     "title": "foreign currency",
     "model": _model(
         fx=[{"currency": "SAR", "rate": 1}, {"currency": "USD", "rate": 3.75}],
         cost_lines=[_cost(currency="USD", quantity=4)])},
    {"id": 3, "kind": ANALYTICAL,
     "title": "multi-year profiling with compounded inflation",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 4, "kind": ANALYTICAL,
     "title": "present value across multiple years",
     "model": _model(base_year=2026, start_year=2027, duration=3, discount_rate=0.10,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 5, "kind": ANALYTICAL,
     "title": "Triangular deterministic basis versus mean",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 6, "kind": ANALYTICAL,
     "title": "Beta-PERT deterministic basis versus mean",
     "model": _model(cost_lines=[_cost(distribution="Beta-PERT")])},
    {"id": 7, "kind": ANALYTICAL,
     "title": "Uniform midpoint equals mean",
     "model": _model(cost_lines=[_cost(distribution="Uniform", most_likely=None)])},
    {"id": 8, "kind": ANALYTICAL,
     "title": "risk expected value with probability below one",
     "model": _model(risks=[_risk()])},
    {"id": 9, "kind": ANALYTICAL,
     "title": "multi-year risk profile",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     risks=[_risk(profile_weights=_PROFILE_3)])},
    {"id": 10, "kind": ANALYTICAL,
     "title": "Base Year equals Start Year",
     "model": _model(base_year=2027, start_year=2027, duration=2,
                     inflation={"Standard": {"2028": 0.05}},
                     cost_lines=[_cost(profile_weights=[0.5, 0.5])])},
    {"id": 11, "kind": ANALYTICAL,
     "title": "Base Year earlier than Start Year",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 12, "kind": ANALYTICAL,
     "title": "zero inflation",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(0.0),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 13, "kind": ANALYTICAL,
     "title": "negative but valid inflation",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(-0.02),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 14, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "blank required inflation rate",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation={"Standard": {"2027": 0.05, "2028": None, "2029": 0.05}},
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 15, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "profile does not sum to one hundred percent",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=[0.2, 0.5, 0.2])])},
    {"id": 16, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "Quantity of zero",
     "model": _model(cost_lines=[_cost(quantity=0)])},
    {"id": 17, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "negative Quantity",
     "model": _model(cost_lines=[_cost(quantity=-5)])},
    {"id": 18, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "discount rate of minus one hundred percent",
     "model": _model(discount_rate=-1.0, cost_lines=[_cost()])},
    {"id": 19, "kind": ANALYTICAL,
     "title": "discount rate negative but above minus one hundred percent",
     "model": _model(base_year=2026, start_year=2027, duration=3, discount_rate=-0.05,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 20, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "inflation rate of minus one hundred percent",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation={"Standard": {"2027": -1.0, "2028": 0.05, "2029": 0.05}},
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 21, "kind": ANALYTICAL,
     "title": "inflation rate negative but above minus one hundred percent",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(-0.02),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 22, "kind": ANALYTICAL,
     "title": "Uniform with a populated Most Likely, which is ignored",
     "model": _model(cost_lines=[_cost(distribution="Uniform", most_likely=999)])},
    {"id": 23, "kind": REFUSAL, "refusal": "ModelInputRefusal",
     "title": "profile summing to one hundred percent but containing a blank",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(),
                     cost_lines=[_cost(profile_weights=[0.5, None, 0.5])])},
    {"id": 24, "kind": REFUSAL, "refusal": "NumericalRangeRefusal",
     "title": "controlled refusal on Double overflow",
     "model": _model(base_year=2026, start_year=2027, duration=3,
                     inflation=_three_year(1e300),
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 25, "kind": ANALYTICAL,
     "title": "unreferenced incomplete FX row does not block",
     "model": _model(
         fx=[{"currency": "SAR", "rate": 1}, {"currency": "EUR", "rate": None}],
         cost_lines=[_cost()])},
    {"id": 26, "kind": FINGERPRINT,
     "title": "fingerprint reference vector",
     "reference": "fingerprint.reference"},
    {"id": 27, "kind": FINGERPRINT,
     "title": "delimiter-hostile field content",
     "reference": "fingerprint.collision_probes"},
    {"id": 28, "kind": STATISTICS,
     "title": "naive overflow with a representable result",
     "statistics": [
         {"statistic": "triangular_mean", "points": [1e308, 1e308, 1e308]},
         {"statistic": "beta_pert_mean", "points": [1e308, 1e308, 1e308]},
         {"statistic": "midpoint", "points": [1.5e308, 1.5e308]},
     ]},
    {"id": 29, "kind": REFUSAL, "refusal": "NumericalRangeRefusal",
     "title": "discount factor underflow",
     "model": _model(base_year=2026, start_year=2026, duration=40,
                     discount_rate=1e10,
                     inflation={"Standard": {str(y): 0.0 for y in range(2027, 2066)}},
                     cost_lines=[_cost(profile_weights=[1.0] + [0.0] * 39)])},
    {"id": 30, "kind": ANALYTICAL,
     "title": "cancellation-heavy reconciliation",
     "model": _model(cost_lines=[
         _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
               min_value=1e16, max_value=1e16, quantity=1, profile_weights=[1.0]),
         _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
               min_value=-1e16, max_value=-1e16, quantity=1, profile_weights=[1.0]),
         _cost(permanent_id="CL-003", distribution="Uniform", most_likely=None,
               min_value=1.0, max_value=1.0, quantity=1, profile_weights=[1.0]),
     ])},
    {"id": 31, "kind": ANALYTICAL,
     "title": "Base-Year factor row",
     "model": _model(base_year=2026, start_year=2028, duration=3,
                     inflation={"Standard": {"2027": 0.05, "2028": 0.05,
                                             "2029": 0.05, "2030": 0.05}},
                     cost_lines=[_cost(profile_weights=_PROFILE_3)])},
    {"id": 32, "kind": RUNTIME_ONLY,
     "title": "derived status reverts to CURRENT after an input is restored",
     "why": "workbook state across attempts; no pure function can evidence it"},
    {"id": 33, "kind": RUNTIME_ONLY,
     "title": "mid-write failure and full logical rollback",
     "why": "requires a real write and a real injected failure"},
    {"id": 34, "kind": RUNTIME_ONLY,
     "title": "invalid input with no Calculate attempted",
     "why": "the two state axes move independently only across real attempts"},
    {"id": 35, "kind": FINGERPRINT,
     "title": "locale separator injection",
     "reference": "fingerprint.decimal_separator"},
    {"id": 36, "kind": FINGERPRINT,
     "title": "reduction beyond Long",
     "reference": "fingerprint.reduction_vectors"},
    {"id": 37, "kind": RUNTIME_ONLY,
     "title": "failure at the commit boundary",
     "why": "requires a real commit-boundary write to fail"},
)


# ---------------------------------------------------------------------------
# The C1 / C2 regression corpus
# ---------------------------------------------------------------------------
# The plan's case matrix predates the edges implementation found. These are the
# FIXED, load-bearing vectors that drove Errata C1 and C2 - a finite, reviewable
# corpus, not a substitute for the Python property sweeps, which stay in the test
# suite where they belong.
REGRESSION_VECTORS: dict[str, tuple[dict[str, Any], ...]] = {
    "signed_sum": (
        {"name": "canonical order is preserved when it succeeds",
         "terms": [1e16, 1.0, -1e16], "expected": 0.0,
         "why": "reordering would give 1.0; tier 1 owns this result"},
        {"name": "headline cancellation",
         "terms": [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE], "expected": MAX_DOUBLE},
        {"name": "cancellation leaving a unit residual",
         "terms": [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, 1.0],
         "expected": 1.0},
        {"name": "cancellation leaving a subnormal residual",
         "terms": [MAX_DOUBLE, MAX_DOUBLE, -MAX_DOUBLE, -MAX_DOUBLE, _SUBNORMAL],
         "expected": _SUBNORMAL},
        {"name": "rounding residual of the cancellation is the answer",
         "terms": [6e307, -8e307, -1.7e308, 6e307, 7e307, 6e307, -1e292],
         "expected": -1e292,
         "why": "a rounded pair-cancellation answers -1.99792015476736e292"},
        {"name": "genuinely outside Double range",
         "terms": [-8e307, -7e307, -1.78e308, 5e307, -1e292, 1e308,
                   -MAX_DOUBLE, 1.78e308],
         "refusal": "NumericalRangeRefusal",
         "why": "exceeds MAX_DOUBLE by about half an ulp and still rounds to it"},
        {"name": "no cancellation and no representable total",
         "terms": [MAX_DOUBLE, MAX_DOUBLE], "refusal": "NumericalRangeRefusal"},
    ),
    "product": (
        {"name": "representable product survives a bad multiplication order",
         "factors": [1e308, 10.0, 0.01], "expected": 1e307},
        {"name": "representable product below the smallest normal",
         "factors": [1e100, 0.5, 1e150, _SUBNORMAL, 1e-250], "expected": _SUBNORMAL},
        {"name": "genuinely outside Double range by under one ulp",
         "factors": [1e50, MAX_DOUBLE, 1e-150, 1e100],
         "refusal": "NumericalRangeRefusal"},
        {"name": "no ordering rescues this product",
         "factors": [1e308, 10.0], "refusal": "NumericalRangeRefusal"},
    ),
    "convex_statistics": (
        {"name": "degenerate Triangular at MAX_DOUBLE", "statistic": "triangular_mean",
         "points": [MAX_DOUBLE, MAX_DOUBLE, MAX_DOUBLE], "expected": MAX_DOUBLE},
        {"name": "degenerate Beta-PERT at MAX_DOUBLE", "statistic": "beta_pert_mean",
         "points": [MAX_DOUBLE, MAX_DOUBLE, MAX_DOUBLE], "expected": MAX_DOUBLE},
        {"name": "degenerate Uniform at the smallest subnormal", "statistic": "midpoint",
         "points": [_SUBNORMAL, _SUBNORMAL], "expected": _SUBNORMAL},
        {"name": "subnormal Uniform midpoint", "statistic": "midpoint",
         "points": [_SUBNORMAL, 1e-323], "expected": 1e-323},
        {"name": "mathematically exact zero", "statistic": "midpoint",
         "points": [-MAX_DOUBLE, MAX_DOUBLE], "expected": 0.0},
        {"name": "non-zero statistic that collapses to zero", "statistic": "midpoint",
         "points": [-20 * _SUBNORMAL, 19 * _SUBNORMAL],
         "refusal": "NumericalRangeRefusal",
         "why": "the exact midpoint is -0.5 * 5e-324 and has no usable Double"},
    ),
    "materialization": (
        {"name": "Reproducer A: Knom refused on a pre-FX intermediate",
         "model": _model(
             base_year=2025, start_year=2026, duration=2, discount_rate=0.0,
             fx=[{"currency": "SAR", "rate": 1}, {"currency": "X", "rate": 0.5}],
             inflation={"P": {"2026": MAX_DOUBLE, "2027": 0.0}},
             cost_lines=[_cost(distribution="Uniform", most_likely=None,
                               min_value=1.0, max_value=1.0, quantity=1,
                               currency="X", inflation_profile="P",
                               profile_weights=[2.0, -1.0])]),
         "why": "2 * MAX_DOUBLE is formed before FX; it is not a published value"},
        {"name": "Reproducer B: annual row refused on a per-driver contribution",
         "model": _model(
             base_year=2025, start_year=2026, duration=3, discount_rate=0.0,
             inflation={"P": {"2026": MAX_DOUBLE, "2027": 0.0, "2028": -0.5}},
             cost_lines=[
                 _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                       min_value=2.0, max_value=2.0, quantity=1,
                       inflation_profile="P", profile_weights=[1.0, -1.0, 1.0]),
                 _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                       min_value=-2.0, max_value=-2.0, quantity=1,
                       inflation_profile="P", profile_weights=[1.0, -1.0, 1.0]),
             ]),
         "why": "each contribution is +/-2 * MAX_DOUBLE; every annual column is 0"},
        {"name": "Reproducer C: annual row refused on a contribution that underflows",
         "model": _model(
             base_year=2025, start_year=2026, duration=2, discount_rate=0.0,
             inflation={"P": {"2026": -0.5, "2027": 1.0}},
             cost_lines=[
                 _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                       min_value=1.0, max_value=1.0, quantity=1,
                       inflation_profile="P", profile_weights=[_SUBNORMAL, 1.0]),
                 _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                       min_value=1.0, max_value=1.0, quantity=1,
                       inflation_profile="P", profile_weights=[_SUBNORMAL, 1.0]),
             ]),
         "why": "each contribution is 0.5 * 5e-324; the annual row is exactly 5e-324"},
        {"name": "a published driver audit amount outside range still refuses",
         "model": _model(
             base_year=2025, start_year=2026, duration=1, discount_rate=0.0,
             inflation={"P": {"2026": MAX_DOUBLE}},
             cost_lines=[
                 _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                       min_value=2.0, max_value=2.0, quantity=1,
                       inflation_profile="P", profile_weights=[1.0]),
                 _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                       min_value=-2.0, max_value=-2.0, quantity=1,
                       inflation_profile="P", profile_weights=[1.0]),
             ]),
         "why": "tblCalcDrivers must publish CL-001's own 2 * MAX_DOUBLE row"},
    ),
    "conditioning": (
        {"name": "headline contribution conditioning survives cancellation",
         "model": _model(cost_lines=[
             _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                   min_value=1e16, max_value=1e16, quantity=1, profile_weights=[1.0]),
             _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                   min_value=-1e16, max_value=-1e16, quantity=1, profile_weights=[1.0]),
         ]),
         "why": "Erratum C1: conditioning on the already-cancelled headline fails"},
        {"name": "within-year annual conditioning survives cancellation",
         "model": _model(
             base_year=2026, start_year=2027, duration=2,
             inflation={"Standard": {"2027": 0.0, "2028": 0.0}},
             cost_lines=[
                 _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                       min_value=1e16, max_value=1e16, quantity=1,
                       profile_weights=[0.5, 0.5]),
                 _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                       min_value=-1e16, max_value=-1e16, quantity=1,
                       profile_weights=[0.5, 0.5]),
             ]),
         "why": "Erratum C1: the annual aggregate is zero while its terms are 1e16 apart"},
    ),
    "row_order": (
        {"name": "canonical permanent-ID order fixes the result",
         "model": _model(cost_lines=[
             _cost(permanent_id="CL-001", distribution="Uniform", most_likely=None,
                   min_value=1e16, max_value=1e16, quantity=1, profile_weights=[1.0]),
             _cost(permanent_id="CL-002", distribution="Uniform", most_likely=None,
                   min_value=1.0, max_value=1.0, quantity=1, profile_weights=[1.0]),
             _cost(permanent_id="CL-003", distribution="Uniform", most_likely=None,
                   min_value=-1e16, max_value=-1e16, quantity=1, profile_weights=[1.0]),
         ]),
         "why": "the same drivers in any row order must give the same totals"},
    ),
}


# ---------------------------------------------------------------------------
# model materialisation and evaluation
# ---------------------------------------------------------------------------
def to_model(payload: dict[str, Any]) -> CalculationModel:
    """The JSON-shaped fixture as the oracle's own value type."""
    timeline = payload["timeline"]
    return CalculationModel(
        timeline=AppliedTimeline(
            timeline["base_year"], timeline["start_year"], timeline["duration"]
        ),
        discount_rate=payload["discount_rate"],
        fx_rows=tuple(FxRow(row["currency"], row["rate"]) for row in payload["fx"]),
        inflation_rates={
            profile: {int(year): rate for year, rate in rates.items()}
            for profile, rates in payload["inflation"].items()
        },
        cost_drivers=tuple(
            CostDriver(
                entry["permanent_id"], entry["distribution"], entry["currency"],
                entry["inflation_profile"], entry["min_value"], entry["most_likely"],
                entry["max_value"], tuple(entry["profile_weights"]),
                quantity=entry["quantity"],
            )
            for entry in payload["cost_lines"]
        ),
        risk_drivers=tuple(
            RiskDriver(
                entry["permanent_id"], entry["distribution"], entry["currency"],
                entry["inflation_profile"], entry["min_value"], entry["most_likely"],
                entry["max_value"], tuple(entry["profile_weights"]),
                probability=entry["probability"],
            )
            for entry in payload["risks"]
        ),
    )


def tolerances_from(calc: CalcContract) -> Tolerances:
    """The contract's tolerance numbers as the oracle's own value type.

    One conversion point, so the oracle never reads YAML and the contract never
    imports the oracle - the same pattern `oracle_limits` uses for Phase 4.
    """
    return Tolerances(
        profiling_sum_absolute=calc.tolerances.profiling_sum_absolute,
        identity_absolute_floor=calc.tolerances.identity_absolute_floor,
        identity_relative_coefficient=calc.tolerances.identity_relative_coefficient,
        conditioning_scale_floor=calc.tolerances.conditioning_scale_floor,
    )


def evaluate(payload: dict[str, Any], tolerances: Tolerances) -> dict[str, Any]:
    """Every published value of one calculated model, as plain JSON data.

    Exactly the values a later harness must compare on real Excel: the resolved
    factor tables, each driver's audit row, the six annual columns and the ten
    headline totals. Nothing derived, nothing rounded, nothing formatted.
    """
    model = to_model(payload)
    result = calculate(model, tolerances)
    return {
        "resolved_fx": dict(sorted(result.resolved_fx.items())),
        # ADDED IN GATE-B B1 CORRECTION ROUND 1, and PROJECTED, not computed.
        #
        # Independent review found two published `_Calc` columns with no emitted
        # expectation behind them - tblCalcYears.Calendar Year and
        # tblCalcFX.Referenced By - so the Windows harness could only have
        # asserted them by deriving `start_year + index - 1` and counting driver
        # references in PowerShell. Both facts are already the oracle's:
        # `AppliedTimeline.project_years()` owns the calendar year and the model's
        # own driver list owns the reference count. No oracle algorithm changed;
        # these blocks only surface what `calculate` already had in hand.
        "calc_years": [
            {"project_index": index, "calendar_year": year,
             "discount_factor": result.discount_factors[index]}
            for index, year in model.timeline.project_years()
        ],
        "resolved_fx_rows": [
            {"currency": currency, "fx_to_sar": rate,
             "referenced_by": sum(
                 1 for driver in (*model.cost_drivers, *model.risk_drivers)
                 if driver.currency == currency
             )}
            for currency, rate in sorted(result.resolved_fx.items())
        ],
        # The applied-timeline text the successful commit records in C16. The
        # FORMAT is modCalcReport.AppliedTimelineText's; this is a checked copy,
        # pinned to that procedure by the Gate-B harness source suite so the two
        # cannot drift.
        "applied_timeline": (
            f"{model.timeline.base_year}/{model.timeline.start_year}"
            f"/{model.timeline.duration}"
        ),
        "inflation_factors": [
            {"profile": row.profile, "calendar_year": row.calendar_year,
             "annual_rate": row.annual_rate, "cumulative_factor": row.cumulative_factor}
            for row in result.inflation_factors
        ],
        "discount_factors": {
            str(index): value for index, value in sorted(result.discount_factors.items())
        },
        "drivers": [
            {
                "permanent_id": driver.permanent_id,
                "driver_kind": driver.driver_kind.value,
                "distribution": driver.distribution,
                "central_basis": driver.central_basis,
                "currency": driver.currency,
                "fx_to_sar": driver.fx_to_sar,
                "inflation_profile": driver.inflation_profile,
                "quantity": driver.quantity,
                "probability": driver.probability,
                "central_value": driver.central_value,
                "mean_value": driver.mean_value,
                "knom": driver.knom,
                "kpv": driver.kpv,
                "deterministic_nominal": driver.deterministic_nominal,
                "deterministic_pv": driver.deterministic_pv,
                "mean_basis_nominal": driver.mean_basis_nominal,
                "mean_basis_pv": driver.mean_basis_pv,
                "uncertainty_mean_shift_nominal": driver.uncertainty_mean_shift_nominal,
                "uncertainty_mean_shift_pv": driver.uncertainty_mean_shift_pv,
                "expected_risk_nominal": driver.expected_risk_nominal,
                "expected_risk_pv": driver.expected_risk_pv,
                "weights": list(driver.weights),
            }
            for driver in result.drivers
        ],
        "annual": [
            {
                "project_index": row.project_index,
                "calendar_year": row.calendar_year,
                "base_cost_nominal": row.base_cost_nominal,
                "expected_risk_nominal": row.expected_risk_nominal,
                "total_nominal": row.total_nominal,
                "base_cost_pv": row.base_cost_pv,
                "expected_risk_pv": row.expected_risk_pv,
                "total_pv": row.total_pv,
            }
            for row in result.annual
        ],
        "totals": {
            "a_nom": result.totals.a_nom, "a_pv": result.totals.a_pv,
            "b_nom": result.totals.b_nom, "b_pv": result.totals.b_pv,
            "c_nom": result.totals.c_nom, "c_pv": result.totals.c_pv,
            "d_nom": result.totals.d_nom, "d_pv": result.totals.d_pv,
            "e_nom": result.totals.e_nom, "e_pv": result.totals.e_pv,
        },
    }


def refusal_of(payload: dict[str, Any], tolerances: Tolerances) -> str:
    """The refusal class a model must produce, or a loud failure if it calculates."""
    try:
        calculate(to_model(payload), tolerances)
    except CalculationRefusal as error:
        return type(error).__name__
    raise RuntimeError("a case declared as a refusal calculated successfully")


def statistic(name: str, points: list[float]) -> float:
    from .calc_numeric import beta_pert_mean, midpoint, triangular_mean

    return {
        "triangular_mean": triangular_mean,
        "beta_pert_mean": beta_pert_mean,
        "midpoint": midpoint,
    }[name](*points)


# ---------------------------------------------------------------------------
# the fingerprint evidence
# ---------------------------------------------------------------------------
UNIT_SEPARATOR = "\u001f"
NUL = "\u0000"
LINE_FEED = "\u000a"

COLLISION_PROBE_INPUTS: tuple[tuple[str, ...], ...] = (
    ("A:B", "C"),
    ("A", "B:C"),
    ("AB", "C"),
    ("A", "B", "C"),
    ("A" + UNIT_SEPARATOR + "B", "C"),
    ("A" + NUL + "B", "C"),
    ("A" + LINE_FEED + "B", "C"),
    ("A", UNIT_SEPARATOR, "C"),
)

# The ten locked canonical numeric encodings of plan section 11.3. Every one is
# carried into the corpus so the later Gate-B diagnostic can exercise the encoder
# DIRECTLY against JSON-owned expectations, rather than hardcoding literals in
# PowerShell or VBA. The label is textual so `-0` stays distinguishable from `0`
# after a JSON round trip.
NUMERIC_ENCODING_VECTORS: tuple[tuple[str, float], ...] = (
    ("0", 0.0),
    ("-0", -0.0),
    ("1", 1.0),
    ("-1", -1.0),
    ("0.1", 0.1),
    ("1e-20", 1e-20),
    ("1e+20", 1e20),
    ("0.1 + 0.2", 0.1 + 0.2),
    ("MAX_DOUBLE", 1.7976931348623157e308),
    ("minimum subnormal", 5e-324),
)

# One more for the separator proof only: the hostile value that drove the Step-1
# positional-normalisation correction. It is not one of the ten locked encodings.
SEPARATOR_EXTRA_VECTORS: tuple[tuple[str, float], ...] = (
    ("-9.87e-5", -9.87e-5),
)

# The two locked UTF-16 behaviours a real VBA implementation has to get right:
# `AscW` returns a SIGNED 16-bit Integer, so every unit above U+7FFF comes back
# negative and must be normalised; and a non-BMP character contributes TWO units,
# so the length prefix counts code units and never code points.
UTF16_VECTORS: tuple[tuple[str, str, str], ...] = (
    ("bmp_above_7fff", "\u9ad8",
     "a single BMP code unit above U+7FFF, where VBA AscW returns a negative Integer"),
    ("non_bmp", "\U0001f600",
     "one code point, two UTF-16 code units: the surrogate pair VBA walks"),
    ("mixed_length_prefix", "A\U0001f600",
     "two code points but three code units, so the length prefix cannot use Len()"),
)

REDUCTION_INPUTS: tuple[tuple[str, int, int], ...] = (
    ("FP_MOD_1", 2147483646, 65535),
    ("FP_MOD_2", 2147483628, 65535),
    ("FP_MOD_1", 1234567890, 41),
    ("FP_MOD_2", 1234567890, 41),
)


def _signed_ascw(unit: int) -> int:
    """The `Integer` VBA `AscW` returns for a UTF-16 code unit.

    `calc_fingerprint.normalise_code_unit` owns the inverse and stays the
    authority: the round trip is asserted here so a drift between the two would
    fail the build rather than ship a wrong expectation.
    """
    signed = unit - 65536 if unit > 32767 else unit
    if fp.normalise_code_unit(signed) != unit:
        raise RuntimeError(f"signed AscW projection disagrees with the authority: {unit}")
    return signed


def reference_stream(fingerprint_version: int) -> str:
    """Golden case 1 as a canonical stream, built by the fingerprint authority.

    Not restated anywhere: `calc_fingerprint.py` owns the encoding, and the Step-3
    test holds its own copy of the expected stream, its 366-code-unit length and
    its digest to prove the emitted value did not drift.
    """
    header = [
        fp.number_field(2026), fp.number_field(2026),
        fp.number_field(1), fp.number_field(0.10),
    ]
    cost = fp.DriverRecord(
        "CL-001",
        (
            fp.text_field("Triangular"), fp.number_field(10), fp.number_field(80),
            fp.number_field(150), fp.number_field(100), fp.number_field(1),
            fp.number_field(1), fp.number_field(1),
        ),
    )
    return fp.build_canonical_stream(
        version=fingerprint_version, header_fields=header, cost_records=[cost]
    )


def fingerprint_section(calc: CalcContract) -> dict[str, Any]:
    """Cases 26, 27, 35 and 36, projected from the fingerprint authority.

    Every value here is COMPUTED by `calc_fingerprint.py`, which owns the hash
    mathematics outright. None is restated in YAML or here.
    """
    moduli = {"FP_MOD_1": fp.FP_MOD_1, "FP_MOD_2": fp.FP_MOD_2}
    stream = reference_stream(calc.fingerprint_version)
    return {
        "constants": {
            "FP_BASE": fp.FP_BASE,
            "FP_MOD_1": fp.FP_MOD_1,
            "FP_MOD_2": fp.FP_MOD_2,
            "FP_INIT_1": fp.FP_INIT_1,
            "FP_INIT_2": fp.FP_INIT_2,
            "FP_VERSION": calc.fingerprint_version,
            "STREAM_TAG": fp.STREAM_TAG,
            "SECTION_ORDER": list(fp.SECTION_ORDER),
        },
        "reference": {
            "case": 26,
            "stream": stream,
            "code_units": fp.utf16_length(stream),
            "digest": fp.fingerprint(stream),
        },
        "collision_probes": [
            {"case": 27, "values": list(values), "digest": fp.fingerprint_probe(values)}
            for values in COLLISION_PROBE_INPUTS
        ],
        "numeric_encodings": {
            "case": 26,
            "vectors": [
                {"label": label, "value": value,
                 "expected": fp.canonical_number(value, ".")}
                for label, value in NUMERIC_ENCODING_VECTORS
            ],
        },
        "utf16_vectors": {
            "case": 26,
            "vectors": [
                {
                    "key": key,
                    "text": text,
                    "code_point_count": len(text),
                    "utf16_length": fp.utf16_length(text),
                    "code_units": list(fp.utf16_code_units(text)),
                    "signed_ascw": [_signed_ascw(unit)
                                    for unit in fp.utf16_code_units(text)],
                    "canonical_text_field": fp.encode_fields([fp.text_field(text)]),
                }
                for key, text, _ in UTF16_VECTORS
            ],
        },
        "decimal_separator": {
            "case": 35,
            "vectors": [
                {"label": label, "value": value,
                 "expected": fp.canonical_number(value, "."),
                 "point": fp.canonical_number(value, "."),
                 "comma": fp.canonical_number(value, ",")}
                for label, value in NUMERIC_ENCODING_VECTORS + SEPARATOR_EXTRA_VECTORS
            ],
        },
        "reduction_vectors": [
            {"case": 36, "modulus_name": name, "modulus": moduli[name], "h": h, "u": u,
             "x": h * fp.FP_BASE + u,
             "remainder": fp.reduce_exact(h, u, moduli[name]),
             "double_only_remainder": fp.reduce_double_only(h, u, moduli[name])}
            for name, h, u in REDUCTION_INPUTS
        ],
    }


# ---------------------------------------------------------------------------
# ===========================================================================
# GATE-B EVIDENCE CORPUS
#
# ADDED IN GATE-B B1 CORRECTION ROUND 2, and deliberately SEPARATE from the 37
# plan cases, which keep their meanings exactly.
#
# The 37 plan cases are the analytical/refusal PLAN corpus. Nine of them are
# prerequisite refusals, and those nine do not exhaust plan section 18: several
# locked predicates - Base Year after Start Year, STRUCTURE CHANGE PENDING, a
# duplicated referenced currency, a non-numeric Probability, an unknown
# Distribution - had no real-Windows scenario at all.
#
# Most of those boundaries CANNOT be expressed as valid analytical models: "abc"
# is not a Discount Rate and a blank is not a Quantity, and the typed oracle is
# right to refuse them at its own boundary. So this corpus describes WORKBOOK
# MUTATIONS instead: what to change, in which cell, and what the refusal must
# say. The oracle is not weakened and no analytical expectation is invented for
# a refused model.
#
# The DETAIL TOKENS are the discriminator. "some error occurred" is not evidence
# that the intended predicate fired, and freezing whole sentences would make a
# harmless wording edit a false failure. Each token below is a fragment of the
# accepted production message that names the PREDICATE - taken from
# modCalcResolve.bas and modCalcCheck.bas.
# ===========================================================================
GATE_B_PREREQUISITE_SCHEMA_VERSION = 1

# `mutation.kind` vocabulary, consumed by the Windows harness:
#   entered_structure   write an ENTERED structural input and do NOT re-apply
#   named_number        write a number into a Setup defined name
#   named_text          write text into a Setup defined name
#   named_blank         clear a Setup defined name
#   register_cell       write into a driver register cell (number / text / null)
#   fx_row              add, retype or clear a tblFXRates row
#   fx_remove           delete the tblFXRates row for a currency
#   inflation_cell      write into the inflation grid at (profile, calendar year)
#   profiling_cell      write into a profiling grid at (driver, project year)
GATE_B_PREREQUISITES: list[dict[str, Any]] = [
    # --- timeline and the structural handoff -----------------------------
    {"id": "PQ-01", "section": "18.T1", "predicate": "base_year_after_start_year",
     "title": "Base Year later than Start Year",
     "base_plan_case": 3,
     "mutation": {"kind": "entered_structure", "target": "base_year", "value": 2030,
                  "apply_timeline": True},
     "detail_tokens": ["Base Year", "Start Year"]},
    {"id": "PQ-02", "section": "18.T2", "predicate": "structure_change_pending",
     "title": "an entered structural value changed and the timeline was not re-applied",
     "base_plan_case": 3,
     "mutation": {"kind": "entered_structure", "target": "duration", "value": 4,
                  "apply_timeline": False},
     "detail_tokens": ["STRUCTURE CHANGE PENDING", "structural prerequisite"]},
    # --- the discount input's type ----------------------------------------
    {"id": "PQ-03", "section": "18.D1", "predicate": "discount_rate_blank",
     "title": "Discount Rate blank",
     "base_plan_case": 3,
     "mutation": {"kind": "named_blank", "target": "discount_rate"},
     "detail_tokens": ["Discount Rate", "blank"]},
    {"id": "PQ-04", "section": "18.D2", "predicate": "discount_rate_non_numeric",
     "title": "Discount Rate non-numeric",
     "base_plan_case": 3,
     "mutation": {"kind": "named_text", "target": "discount_rate", "value": "abc"},
     "detail_tokens": ["Discount Rate", "not numeric"]},
    # --- FX and the reporting-currency invariant ---------------------------
    {"id": "PQ-05", "section": "18.F1", "predicate": "referenced_currency_missing",
     "title": "a referenced foreign currency has no tblFXRates row",
     "base_plan_case": 2,
     "mutation": {"kind": "fx_remove", "currency": "USD"},
     "detail_tokens": ["FX", "USD", "rows"]},
    {"id": "PQ-06", "section": "18.F2", "predicate": "referenced_currency_duplicated",
     "title": "a referenced foreign currency appears twice",
     "base_plan_case": 2,
     "mutation": {"kind": "fx_row", "currency": "USD", "rate": 3.75, "append": True},
     "detail_tokens": ["FX", "USD", "rows"]},
    {"id": "PQ-07", "section": "18.F3", "predicate": "referenced_rate_not_positive",
     "title": "a referenced FX rate is not strictly positive",
     "base_plan_case": 2,
     "mutation": {"kind": "fx_row", "currency": "USD", "rate": 0},
     "detail_tokens": ["FX", "USD", "strictly positive"]},
    {"id": "PQ-08", "section": "18.F4", "predicate": "referenced_rate_blank",
     "title": "a referenced FX rate is blank",
     "base_plan_case": 2,
     "mutation": {"kind": "fx_row", "currency": "USD", "rate": None},
     "detail_tokens": ["FX rate for referenced currency", "USD", "blank"]},
    {"id": "PQ-09", "section": "18.F5", "predicate": "referenced_rate_non_numeric",
     "title": "a referenced FX rate is non-numeric",
     "base_plan_case": 2,
     "mutation": {"kind": "fx_row", "currency": "USD", "rate": "n/a"},
     "detail_tokens": ["FX rate for referenced currency", "USD", "not numeric"]},
    {"id": "PQ-10", "section": "18.F6", "predicate": "reporting_currency_missing",
     "title": "the reporting currency has no row",
     "base_plan_case": 3,
     "mutation": {"kind": "fx_remove", "currency": "SAR"},
     "detail_tokens": ["the reporting currency", "SAR", "exactly once"]},
    {"id": "PQ-11", "section": "18.F7", "predicate": "reporting_currency_duplicated",
     "title": "the reporting currency appears twice",
     "base_plan_case": 3,
     "mutation": {"kind": "fx_row", "currency": "SAR", "rate": 1, "append": True},
     "detail_tokens": ["the reporting currency", "SAR", "exactly once"]},
    {"id": "PQ-12", "section": "18.F8", "predicate": "reporting_currency_rate_not_one",
     "title": "the reporting currency rate is not exactly 1",
     "base_plan_case": 3,
     "mutation": {"kind": "fx_row", "currency": "SAR", "rate": 2},
     "detail_tokens": ["the reporting currency", "SAR", "must resolve to"]},
    # --- inflation ---------------------------------------------------------
    {"id": "PQ-13", "section": "18.I1", "predicate": "referenced_profile_missing",
     "title": "a referenced inflation profile is absent from the grid",
     "base_plan_case": 3,
     "mutation": {"kind": "inflation_profile_rename", "profile": "Standard",
                  "value": "Renamed"},
     "detail_tokens": ["inflation: profile", "Standard", "not present"]},
    {"id": "PQ-14", "section": "18.I2", "predicate": "referenced_rate_non_numeric",
     "title": "a required referenced inflation rate is non-numeric",
     "base_plan_case": 3,
     "mutation": {"kind": "inflation_cell", "profile": "Standard", "calendar_year": 2027,
                  "value": "n/a"},
     "detail_tokens": ["inflation profile", "2027", "not numeric"]},
    # --- profiling ---------------------------------------------------------
    {"id": "PQ-15", "section": "18.P1", "predicate": "profiling_cell_non_numeric",
     "title": "a required profiling cell is non-numeric",
     "base_plan_case": 3,
     "mutation": {"kind": "profiling_cell", "grid": "cost_profiling",
                  "permanent_id": "CL-001", "project_year": 2, "value": "n/a"},
     "detail_tokens": ["profiling for driver", "CL-001", "not numeric"]},
    # --- distribution ------------------------------------------------------
    {"id": "PQ-16", "section": "18.X1", "predicate": "distribution_missing",
     "title": "Distribution is blank",
     "base_plan_case": 3,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "distribution", "value": None},
     "detail_tokens": ["Distribution", "blank"]},
    {"id": "PQ-17", "section": "18.X2", "predicate": "distribution_unknown",
     "title": "Distribution is not one of the three accepted kinds",
     "base_plan_case": 3,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "distribution", "value": "Lognormal"},
     "detail_tokens": ["Distribution", "not an accepted distribution"]},
    # --- three-point ordering ----------------------------------------------
    {"id": "PQ-18", "section": "18.O1", "predicate": "triangular_ordering",
     "title": "Triangular Min <= Most Likely <= Max is violated",
     "base_plan_case": 3,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "unit_cost_most_likely",
                  "value": 500},
     "detail_tokens": ["Triangular", "Min <= Most Likely <= Max"]},
    {"id": "PQ-19", "section": "18.O2", "predicate": "beta_pert_ordering",
     "title": "Beta-PERT Min <= Most Likely <= Max is violated",
     "base_plan_case": 6,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "unit_cost_most_likely",
                  "value": 500},
     "detail_tokens": ["Beta-PERT", "Min <= Most Likely <= Max"]},
    {"id": "PQ-20", "section": "18.O3", "predicate": "uniform_ordering",
     "title": "Uniform Min <= Max is violated",
     "base_plan_case": 7,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "unit_cost_min", "value": 900},
     "detail_tokens": ["Uniform", "Min <= Max"]},
    # --- Quantity ----------------------------------------------------------
    {"id": "PQ-21", "section": "18.Q1", "predicate": "quantity_missing",
     "title": "Quantity is blank",
     "base_plan_case": 3,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "quantity", "value": None},
     "detail_tokens": ["Quantity", "blank"]},
    {"id": "PQ-22", "section": "18.Q2", "predicate": "quantity_non_numeric",
     "title": "Quantity is non-numeric",
     "base_plan_case": 3,
     "mutation": {"kind": "register_cell", "register": "cost_lines",
                  "permanent_id": "CL-001", "column": "quantity", "value": "n/a"},
     "detail_tokens": ["Quantity", "not numeric"]},
    # --- Probability -------------------------------------------------------
    {"id": "PQ-23", "section": "18.R1", "predicate": "probability_missing",
     "title": "Probability is blank",
     "base_plan_case": 9,
     "mutation": {"kind": "register_cell", "register": "risk_register",
                  "permanent_id": "R-001", "column": "probability", "value": None},
     "detail_tokens": ["Probability", "blank"]},
    {"id": "PQ-24", "section": "18.R2", "predicate": "probability_non_numeric",
     "title": "Probability is non-numeric",
     "base_plan_case": 9,
     "mutation": {"kind": "register_cell", "register": "risk_register",
                  "permanent_id": "R-001", "column": "probability", "value": "n/a"},
     "detail_tokens": ["Probability", "not numeric"]},
    {"id": "PQ-25", "section": "18.R3", "predicate": "probability_below_zero",
     "title": "Probability below zero",
     "base_plan_case": 9,
     "mutation": {"kind": "register_cell", "register": "risk_register",
                  "permanent_id": "R-001", "column": "probability", "value": -0.5},
     "detail_tokens": ["Probability", "fraction in [0, 1]"]},
    {"id": "PQ-26", "section": "18.R4", "predicate": "probability_above_one",
     "title": "Probability above one",
     "base_plan_case": 9,
     "mutation": {"kind": "register_cell", "register": "risk_register",
                  "permanent_id": "R-001", "column": "probability", "value": 1.5},
     "detail_tokens": ["Probability", "fraction in [0, 1]"]},
]

# THE COMPLEMENT. Referenced-only resolution means an assumption nobody uses
# cannot block a valid model, and that is a locked semantic in its own right: a
# harness that only proved refusals would accept a model that refused too much.
GATE_B_NO_BLOCK: list[dict[str, Any]] = [
    {"id": "PN-01", "section": "18.N1", "predicate": "unreferenced_fx_duplicated",
     "title": "an UNREFERENCED foreign currency appearing twice does not block",
     "base_plan_case": 3,
     "mutation": {"kind": "fx_row", "currency": "ZZZ", "rate": 3.75, "append": True,
                  "repeat": 2}},
    {"id": "PN-02", "section": "18.N2", "predicate": "unreferenced_fx_blank_rate",
     "title": "an UNREFERENCED foreign currency with a blank rate does not block",
     "base_plan_case": 3,
     "mutation": {"kind": "fx_row", "currency": "ZZZ", "rate": None, "append": True}},
    {"id": "PN-03", "section": "18.N3", "predicate": "unreferenced_profile_incomplete",
     "title": "an UNREFERENCED inflation profile with a missing rate does not block",
     "base_plan_case": 3,
     "mutation": {"kind": "inflation_profile_add", "profile": "Unused",
                  "calendar_year": 2027, "value": None}},
]

# The multi-driver fixture the driver-audit reconstruction needs. Several Cost
# Lines AND several Risks, so A, B, C and D are all non-trivially represented and
# the cost/risk partition is not vacuous. Built from the accepted model helpers
# and evaluated through the accepted oracle - no expected value is invented.
GATE_B_AUDIT_MODEL = _model(
    base_year=2026, start_year=2027, duration=3,
    inflation={"Standard": {"2027": 0.05, "2028": 0.05, "2029": 0.05},
               "Flat": {"2027": 0.0, "2028": 0.0, "2029": 0.0}},
    fx=[{"currency": "SAR", "rate": 1}, {"currency": "USD", "rate": 3.75}],
    cost_lines=[
        _cost(permanent_id="CL-001", profile_weights=_PROFILE_3),
        _cost(permanent_id="CL-002", distribution="Beta-PERT", currency="USD",
              quantity=4, min_value=50, most_likely=70, max_value=120,
              profile_weights=[0.5, 0.25, 0.25]),
        _cost(permanent_id="CL-003", distribution="Uniform", most_likely=None,
              inflation_profile="Flat", quantity=2, min_value=10, max_value=30,
              profile_weights=[0.0, 0.0, 1.0]),
    ],
    risks=[
        _risk(permanent_id="R-001", profile_weights=_PROFILE_3, probability=0.3),
        _risk(permanent_id="R-002", distribution="Uniform", most_likely=None,
              currency="USD", min_value=20, max_value=60, probability=0.75,
              profile_weights=[0.4, 0.4, 0.2]),
    ],
)

# The audit-column relationships the reconstruction check must use. Ordinals are
# the accepted tblCalcDrivers schema's, and they are emitted so PowerShell reads
# them rather than restating "column 18" in its own source.
GATE_B_AUDIT_RELATIONSHIPS: list[dict[str, Any]] = [
    {"headline": "a_nom", "driver_column": "deterministic_nominal", "kind": "Cost Line"},
    {"headline": "a_pv", "driver_column": "deterministic_pv", "kind": "Cost Line"},
    {"headline": "b_nom", "driver_column": "uncertainty_mean_shift_nominal", "kind": "Cost Line"},
    {"headline": "b_pv", "driver_column": "uncertainty_mean_shift_pv", "kind": "Cost Line"},
    {"headline": "c_nom", "driver_column": "mean_basis_nominal", "kind": "Cost Line"},
    {"headline": "c_pv", "driver_column": "mean_basis_pv", "kind": "Cost Line"},
    {"headline": "d_nom", "driver_column": "expected_risk_nominal", "kind": "Risk"},
    {"headline": "d_pv", "driver_column": "expected_risk_pv", "kind": "Risk"},
]


# The detail discriminators for the NINE refusal plan cases. They stay in the
# Gate-B corpus rather than being added to the plan cases themselves: the 37 keep
# their meanings exactly, and PowerShell still holds no list of its own. Each
# token is a fragment of the accepted production message that names the
# PREDICATE - from modCalcResolve.bas, modCalcCheck.bas and modCalcFactors.bas.
GATE_B_PLAN_REFUSAL_TOKENS: dict[int, list[str]] = {
    14: ["inflation profile", "blank"],
    15: ["profiling weights sum to"],
    16: ["Quantity", "strictly positive"],
    17: ["Quantity", "strictly positive"],
    18: ["discount rate", "1 + r <= 0"],
    20: ["inflation profile", "1 + rate <= 0"],
    23: ["profiling for driver", "blank"],
    24: ["inflation factor"],
    29: ["discount factors"],
}


def gate_b_section(tolerances: Tolerances) -> dict[str, Any]:
    """The Gate-B-only evidence corpus. Separate from `plan_cases` by design."""
    audit_expected = evaluate(GATE_B_AUDIT_MODEL, tolerances)
    cost_ids = {line["permanent_id"] for line in GATE_B_AUDIT_MODEL["cost_lines"]}
    risk_ids = {risk["permanent_id"] for risk in GATE_B_AUDIT_MODEL["risks"]}
    if not cost_ids or not risk_ids:
        raise RuntimeError("the audit fixture must carry both Cost Lines and Risks")
    if len(cost_ids) < 2:
        raise RuntimeError("the audit fixture must carry more than one Cost Line")
    refusals = {entry["id"] for entry in CASES if entry["kind"] == REFUSAL}
    if refusals != set(GATE_B_PLAN_REFUSAL_TOKENS):
        raise RuntimeError(
            "every refusal plan case needs a Gate-B detail discriminator; "
            f"missing {sorted(refusals - set(GATE_B_PLAN_REFUSAL_TOKENS))}, "
            f"extra {sorted(set(GATE_B_PLAN_REFUSAL_TOKENS) - refusals)}"
        )
    return {
        "schema_version": GATE_B_PREREQUISITE_SCHEMA_VERSION,
        "purpose": (
            "Gate-B-only runtime evidence: the plan section 18 prerequisite matrix as "
            "WORKBOOK MUTATIONS, the referenced-only no-block complement, and the "
            "multi-driver audit-reconstruction fixture. Separate from plan_cases, "
            "which keep their meanings."
        ),
        "prerequisite_cases": [
            {**entry, "expected_attempt": "REFUSED", "expected_status": "INVALID",
             "snapshot_unchanged": True}
            for entry in GATE_B_PREREQUISITES
        ],
        "no_block_cases": [
            {**entry, "expected_attempt": "SUCCESS", "expected_status": "CURRENT",
             "detail_tokens": [], "snapshot_unchanged": False}
            for entry in GATE_B_NO_BLOCK
        ],
        "plan_refusal_tokens": {
            str(case_id): list(tokens)
            for case_id, tokens in sorted(GATE_B_PLAN_REFUSAL_TOKENS.items())
        },
        "audit_reconstruction": {
            "title": "driver-audit A/B/C/D reconstruction over a multi-driver model",
            "model": GATE_B_AUDIT_MODEL,
            "expected": audit_expected,
            "relationships": GATE_B_AUDIT_RELATIONSHIPS,
        },
    }


def build_cases(calc: CalcContract, model_version: str) -> dict[str, Any]:
    """The whole `phase5_cases.json` document, deterministically ordered."""
    tolerances = tolerances_from(calc)
    cases: list[dict[str, Any]] = []
    for entry in CASES:
        record: dict[str, Any] = {
            "id": entry["id"], "kind": entry["kind"], "title": entry["title"]
        }
        if entry["kind"] == ANALYTICAL:
            record["model"] = entry["model"]
            record["expected"] = evaluate(entry["model"], tolerances)
        elif entry["kind"] == REFUSAL:
            record["model"] = entry["model"]
            observed = refusal_of(entry["model"], tolerances)
            if observed != entry["refusal"]:
                raise RuntimeError(
                    f"case {entry['id']} expects {entry['refusal']}, oracle gave {observed}"
                )
            record["expected_refusal"] = observed
        elif entry["kind"] == STATISTICS:
            record["statistics"] = [
                {**item, "expected": statistic(item["statistic"], item["points"])}
                for item in entry["statistics"]
            ]
        elif entry["kind"] == FINGERPRINT:
            record["reference"] = entry["reference"]
        else:
            record["why"] = entry["why"]
        cases.append(record)

    return {
        "schema_version": SCHEMA_VERSION,
        "purpose": (
            "Expected-value corpus for the later Windows/VBA acceptance harness. "
            "Test data only: nothing here is read by the workbook, and nothing here "
            "defines the mathematics."
        ),
        "provenance": {
            "model_version": model_version,
            "calc_contract_version": calc.version,
            "fingerprint_version": calc.fingerprint_version,
            "oracle": "builder/pccm_builder/calc_oracle.py",
            "numerical_kernel": "builder/pccm_builder/calc_numeric.py",
            "fingerprint_authority": "builder/pccm_builder/calc_fingerprint.py",
        },
        "tolerances": {
            "profiling_sum_absolute": calc.tolerances.profiling_sum_absolute,
            "identity_absolute_floor": calc.tolerances.identity_absolute_floor,
            "identity_relative_coefficient": calc.tolerances.identity_relative_coefficient,
            "conditioning_scale_floor": calc.tolerances.conditioning_scale_floor,
        },
        "fingerprint": fingerprint_section(calc),
        "plan_cases": cases,
        "gate_b": gate_b_section(tolerances),
        "regression_vectors": _regression_section(tolerances),
    }


def _regression_section(tolerances: Tolerances) -> dict[str, Any]:
    """The C1/C2 corpus, with every model-shaped vector actually evaluated."""
    section: dict[str, Any] = {}
    for key in sorted(REGRESSION_VECTORS):
        rendered: list[dict[str, Any]] = []
        for entry in REGRESSION_VECTORS[key]:
            record = dict(entry)
            if "model" in record:
                try:
                    record["expected"] = evaluate(record["model"], tolerances)
                except CalculationRefusal as error:
                    record["expected_refusal"] = type(error).__name__
            elif "statistic" in record and "refusal" in record:
                record["expected_refusal"] = record.pop("refusal")
            elif "refusal" in record:
                record["expected_refusal"] = record.pop("refusal")
            rendered.append(record)
        section[key] = rendered
    return section
