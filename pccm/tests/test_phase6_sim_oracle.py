#!/usr/bin/env python3
"""PCCM Phase 6 Step-4 conformance tests for the pure Python simulation oracle.

`builder/pccm_builder/sim_oracle.py` turns one accepted resolved model, an
iteration count and an effective seed into per-iteration nominal and PV totals, a
result digest and the summary statistics that follow. These tests prove the
contribution rules, the D6-18b severity discipline, canonical accumulation order,
replay, row-order invariance, every retained Step-0 digest vector, the statistics
and the reporting-only role of the Selected Confidence Level.

WHAT IS PROVED HERE IS SEMANTICS, NOT SPEED. Elapsed runtime is not a Phase-6
gate - the design-target performance authority is the later Windows Gate B - so
the fixtures are representative rather than large, and no test runs a
300-driver 100,000-iteration simulation to make a point about structure.

NO WORKBOOK, NO EXCEL, NO COM, NO `_SimData`, NO VBA. Nothing here writes
anything anywhere.

Runs standalone or under pytest.
"""

from __future__ import annotations

import json
import math
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    RngReference,
    load_calc_contract,
    load_contract,
    load_sim_contract,
)
from pccm_builder.calc_cases import tolerances_from  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    CalculationRefusal,
    NumericalRangeRefusal,
    safe_signed_sum,
)
from pccm_builder.calc_oracle import (  # noqa: E402
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    FxRow,
    RiskDriver,
)
from pccm_builder.sim_oracle import (  # noqa: E402
    Contingency,
    DeterministicBase,
    SimOracleError,
    SimulationSummary,
    business_minimum_iterations,
    contingency_at,
    deterministic_base_of,
    effective_seed_from_nonce,
    prepare_simulation,
    resolve_percentile_ladder,
    result_digest,
    result_digest_stream,
    rng_reference_signature,
    run_simulation,
    validate_iterations,
    validate_result_digest_contract,
)
from pccm_builder.sim_rng import COST_KIND, RISK_KIND  # noqa: E402
from pccm_builder.sim_stats import (  # noqa: E402
    MeasureStatistics,
    SimStatsError,
    describe,
    percentile_type7,
    sample_mean,
    sample_standard_deviation,
)

SPEC = PCCM_ROOT / "spec"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0"

_CACHE: dict[str, object] = {}


def _inputs():
    if "inputs" not in _CACHE:
        _CACHE["inputs"] = load_contract(SPEC / "input_contract.yaml")
    return _CACHE["inputs"]


def _sim():
    if "sim" not in _CACHE:
        _CACHE["sim"] = load_sim_contract(SPEC / "sim_contract.yaml")
    return _CACHE["sim"]


def _tolerances():
    if "tol" not in _CACHE:
        _CACHE["tol"] = tolerances_from(load_calc_contract(SPEC / "calc_contract.yaml"))
    return _CACHE["tol"]


def _ref() -> RngReference:
    if "ref" not in _CACHE:
        _CACHE["ref"] = RngReference.from_contracts(_sim(), _inputs())
    return _CACHE["ref"]


# ---------------------------------------------------------------------------
# fixtures
#
# Every fixture below uses a single applied project year with weight 1, no
# inflation and FX 1, so `Knom` is exactly 1 and `Kpv` is exactly 1. That is
# deliberate: a contribution is then exactly `sample * Quantity`, and a test can
# state the arithmetic it expects instead of comparing against a second
# implementation of the escalation path.
# ---------------------------------------------------------------------------
def _cost(
    permanent_id: str,
    distribution: str = "Triangular",
    minimum: object = 80.0,
    most_likely: object = 100.0,
    maximum: object = 150.0,
    quantity: object = 1.0,
) -> CostDriver:
    return CostDriver(
        permanent_id, distribution, "SAR", "Standard",
        minimum, most_likely, maximum, (1.0,), quantity=quantity,
    )


def _risk(
    permanent_id: str,
    distribution: str = "Triangular",
    minimum: object = 100.0,
    most_likely: object = 200.0,
    maximum: object = 400.0,
    probability: object = 0.5,
) -> RiskDriver:
    return RiskDriver(
        permanent_id, distribution, "SAR", "Standard",
        minimum, most_likely, maximum, (1.0,), probability=probability,
    )


def _model(costs=(), risks=(), discount_rate: object = 0.10) -> CalculationModel:
    return CalculationModel(
        timeline=AppliedTimeline(2026, 2026, 1),
        discount_rate=discount_rate,
        fx_rows=(FxRow("SAR", 1),),
        inflation_rates={"Standard": {}},
        cost_drivers=tuple(costs),
        risk_drivers=tuple(risks),
    )


def _prepare(model: CalculationModel, seed: int = 12345, iterations: int = 1000):
    return prepare_simulation(
        _ref(), _sim(), _inputs(), model, _tolerances(),
        effective_seed=seed, iterations=iterations,
    )


def _run(model: CalculationModel, seed: int = 12345, iterations: int = 1000):
    prepared, result = _prepare(model, seed, iterations)
    return run_simulation(_ref(), prepared), prepared, result


def _diag(run, permanent_id: str, role: str):
    for record in run.diagnostics:
        if record.permanent_id == permanent_id and record.role == role:
            return record
    raise AssertionError(f"no {role} diagnostics for {permanent_id!r}")


# ===========================================================================
# A - one degenerate Cost Line
# ===========================================================================
def test_01_a_degenerate_cost_line_is_constant_and_consumes_nothing() -> None:
    """Zero uncertainty means zero draws, and every iteration is the same number."""
    model = _model(costs=[_cost("CL-001", "Uniform", 250.0, None, 250.0, quantity=2.0)])
    run, prepared, result = _run(model)

    assert set(run.total_nominal) == {500.0}, "a degenerate driver produced a spread"
    assert set(run.total_pv) == {500.0}
    stream = _diag(run, "CL-001", "value")
    assert stream.uniforms_consumed == 0, "a degenerate driver consumed a uniform"
    assert stream.final_state == stream.initial_state, "the stream advanced"
    assert run.summary.nominal.sample_standard_deviation == 0.0
    assert prepared.deterministic_base.nominal == 500.0
    assert result.totals.a_nom == 500.0


def test_02_a_uniform_ignores_most_likely_all_the_way_through_the_engine() -> None:
    """D1 survives preparation: a populated Most Likely reaches nothing."""
    plain = _model(costs=[_cost("CL-001", "Uniform", 250.0, None, 250.0, quantity=2.0)])
    noisy = _model(costs=[_cost("CL-001", "Uniform", 250.0, 9.9e99, 250.0, quantity=2.0)])
    left, prepared_left, _ = _run(plain)
    right, prepared_right, _ = _run(noisy)

    assert prepared_left.cost_drivers[0].most_likely is None
    assert prepared_right.cost_drivers[0].most_likely is None, (
        "the ignored Most Likely was carried into the prepared model"
    )
    assert left.total_nominal == right.total_nominal
    assert left.result_digest == right.result_digest


# ===========================================================================
# B - one stochastic Cost Line
# ===========================================================================
def test_03_a_stochastic_cost_line_spans_its_support() -> None:
    model = _model(costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=1.0)])
    run, _, _ = _run(model, iterations=2000)

    assert len(run.total_nominal) == 2000
    assert len(set(run.total_nominal)) > 1900, "the sampler produced a suspiciously flat run"
    assert min(run.total_nominal) >= 80.0 and max(run.total_nominal) <= 150.0
    assert _diag(run, "CL-001", "value").uniforms_consumed == 2000, (
        "a Triangular consumes exactly one uniform per iteration"
    )


# ===========================================================================
# C - Quantity applied exactly once
# ===========================================================================
def test_04_quantity_is_deterministic_and_applied_exactly_once() -> None:
    """The total is LINEAR in Quantity. Twice would be quadratic; omitted, flat."""
    unit = 250.0
    for quantity in (1.0, 3.0, 7.5):
        model = _model(
            costs=[_cost("CL-001", "Uniform", unit, None, unit, quantity=quantity)]
        )
        run, _, _ = _run(model)
        expected = unit * quantity
        assert run.total_nominal[0] == expected, (
            f"Quantity {quantity}: got {run.total_nominal[0]}, expected {expected}. "
            f"Applied twice would give {unit * quantity * quantity}."
        )
        assert run.total_pv[0] == expected


def test_05_quantity_sits_outside_the_distribution_not_inside_it() -> None:
    """A stochastic line scales exactly, sample for sample, when Quantity changes.

    If Quantity were inside the distribution the samples themselves would differ,
    the stream would consume differently, and the ratio would not be constant.
    """
    single = _model(costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=1.0)])
    triple = _model(costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=3.0)])
    one, _, _ = _run(single, iterations=1000)
    three, _, _ = _run(triple, iterations=1000)

    assert _diag(one, "CL-001", "value").final_state == _diag(three, "CL-001", "value").final_state
    for left, right in zip(one.total_nominal, three.total_nominal):
        assert right == left * 3.0


# ===========================================================================
# D, E, F - Risk occurrence
# ===========================================================================
def test_06_a_risk_with_probability_zero_never_occurs() -> None:
    model = _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.0)])
    run, _, _ = _run(model, iterations=2000)

    assert set(run.total_nominal) == {0.0}, "a p = 0 risk contributed"
    assert _diag(run, "R-001", "occurrence").uniforms_consumed == 2000, (
        "occurrence consumes one uniform per iteration whatever the probability"
    )
    assert _diag(run, "R-001", "severity").uniforms_consumed == 2000, (
        "D6-18b: severity is sampled even when the risk can never occur"
    )


def test_07_a_risk_with_probability_one_always_occurs() -> None:
    model = _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=1.0)])
    run, _, _ = _run(model, iterations=2000)

    assert all(total > 0.0 for total in run.total_nominal), "a p = 1 risk failed to occur"
    assert min(run.total_nominal) >= 100.0 and max(run.total_nominal) <= 400.0
    assert _diag(run, "R-001", "occurrence").uniforms_consumed == 2000


def test_08_an_intermediate_probability_occurs_at_about_its_rate() -> None:
    iterations = 5000
    model = _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.3)])
    run, _, _ = _run(model, iterations=iterations)

    occurred = sum(1 for total in run.total_nominal if total != 0.0)
    rate = occurred / iterations
    assert abs(rate - 0.3) < 0.02, f"occurrence rate {rate} is not near 0.3"
    assert 0 < occurred < iterations, "the run is degenerate in one direction"


# ===========================================================================
# G, H - D6-18b
# ===========================================================================
def test_09_severity_is_invoked_every_risk_iteration() -> None:
    """The severity stream advances by a full run's worth of draws, not by the
    number of occurrences."""
    iterations = 3000
    model = _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.05)])
    run, _, _ = _run(model, iterations=iterations)

    occurred = sum(1 for total in run.total_nominal if total != 0.0)
    severity = _diag(run, "R-001", "severity")
    assert 0 < occurred < iterations // 4, "the fixture must occur rarely to be meaningful"
    assert severity.uniforms_consumed == iterations, (
        f"severity consumed {severity.uniforms_consumed} uniforms for {iterations} iterations; "
        f"sampling only on occurrence would have consumed about {occurred}"
    )


def test_10_a_degenerate_risk_severity_consumes_nothing_but_is_still_invoked() -> None:
    """Degenerate severity: zero consumption, unchanged stream, and the risk still
    contributes its constant whenever it occurs."""
    model = _model(risks=[_risk("R-001", "Triangular", 90.0, 90.0, 90.0, probability=0.4)])
    run, _, _ = _run(model, iterations=2000)

    severity = _diag(run, "R-001", "severity")
    assert severity.uniforms_consumed == 0
    assert severity.final_state == severity.initial_state
    assert _diag(run, "R-001", "occurrence").uniforms_consumed == 2000
    assert set(run.total_nominal) == {0.0, 90.0}


# ===========================================================================
# I - probability-only comparability
# ===========================================================================
def test_11_changing_only_probability_leaves_the_severity_sequence_identical() -> None:
    """THE REASON D6-18B EXISTS.

    Two runs of the same model, the same seed and the same severity
    distribution, differing only in Probability. The occurrence decisions differ
    and the contributions differ - but the severity stream is drawn the same
    number of times, lands in the same final state, and produces the same
    iteration-indexed severity sequence. Sampling severity only on occurrence
    would make the two runs incomparable.
    """
    iterations = 2000
    runs = []
    for probability in (0.2, 0.8):
        model = _model(
            costs=[_cost("CL-001", "Uniform", 10.0, None, 10.0, quantity=1.0)],
            risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=probability)],
        )
        runs.append(_run(model, iterations=iterations)[0])
    low, high = runs

    severity_low = _diag(low, "R-001", "severity")
    severity_high = _diag(high, "R-001", "severity")
    assert severity_low.final_state == severity_high.final_state, (
        "the severity streams diverged, so the two runs cannot be compared"
    )
    assert severity_low.uniforms_consumed == severity_high.uniforms_consumed == iterations

    # The severity SEQUENCE itself, reconstructed independently of the engine
    # from the stream's own initial state, is the same for both runs.
    sequence = _replay_severity(severity_low.initial_state, iterations)
    assert _replay_severity(severity_high.initial_state, iterations) == sequence

    # And the runs really do differ where they are supposed to.
    low_occurred = sum(1 for total in low.total_nominal if total != 10.0)
    high_occurred = sum(1 for total in high.total_nominal if total != 10.0)
    assert low_occurred != high_occurred, "the control is vacuous: the probabilities did nothing"
    assert low.result_digest != high.result_digest


def _replay_severity(initial_state, iterations: int) -> tuple[float, ...]:
    """The severity sequence, drawn straight from the stream by the samplers."""
    from pccm_builder import sample_triangular

    state = initial_state
    out = []
    for _ in range(iterations):
        drawn = sample_triangular(_ref(), state, 100.0, 200.0, 400.0)
        out.append(drawn.value)
        state = drawn.state
    return tuple(out)


def test_12_occurrence_and_severity_never_share_a_stream() -> None:
    model = _model(risks=[_risk("R-001"), _risk("R-002"), _risk("R-003")])
    run, prepared, _ = _run(model)

    indices = [record.stream_index for record in run.diagnostics]
    assert len(indices) == len(set(indices)), f"components share a stream: {indices}"
    for driver in prepared.risk_drivers:
        assert driver.occurrence_stream_index != driver.value_stream_index
        assert driver.occurrence_initial_state != driver.value_initial_state


# ===========================================================================
# stream initialisation - section 6
# ===========================================================================
def test_13_streams_are_assigned_cost_first_then_each_risk_interleaved() -> None:
    """Cost Lines first; then each Risk in Permanent-ID order, occurrence before
    severity. Not three global blocks."""
    model = _model(
        costs=[_cost("CL-002"), _cost("CL-001")],
        risks=[_risk("R-002"), _risk("R-001")],
    )
    run, prepared, _ = _run(model)

    ordered = [
        (record.kind, record.permanent_id, record.role)
        for record in sorted(run.diagnostics, key=lambda item: item.stream_index)
    ]
    assert ordered == [
        (COST_KIND, "CL-001", "value"),
        (COST_KIND, "CL-002", "value"),
        (RISK_KIND, "R-001", "occurrence"),
        (RISK_KIND, "R-001", "severity"),
        (RISK_KIND, "R-002", "occurrence"),
        (RISK_KIND, "R-002", "severity"),
    ]
    assert [record.stream_index for record in run.diagnostics] == [0, 1, 2, 3, 4, 5]
    assert prepared.base_state == _ref().fixed_seed_to_state(prepared.effective_seed)


def test_14_the_engine_uses_the_accepted_step2_stream_states_unchanged() -> None:
    """Every initial state is exactly what the Step-2 reference assigns."""
    model = _model(costs=[_cost("CL-001"), _cost("CL-002")], risks=[_risk("R-001")])
    run, prepared, _ = _run(model)

    components = _ref().components_for(("CL-001", "CL-002"), ("R-001",))
    expected = {
        (component.kind, component.permanent_id, component.role): (index, state)
        for component, index, state in _ref().component_stream_states(
            _ref().fixed_seed_to_state(prepared.effective_seed), components
        )
    }
    for record in run.diagnostics:
        index, state = expected[(record.kind, record.permanent_id, record.role)]
        assert record.stream_index == index
        assert record.initial_state == state


# ===========================================================================
# J, K, L, M - whole simulations
# ===========================================================================
def _mixed_model(risk_distribution: str = "Triangular") -> CalculationModel:
    return _model(
        costs=[
            _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, quantity=1.0),
            _cost("CL-003", "Beta-PERT", 5.0, 7.0, 20.0, quantity=3.0),
        ],
        risks=[
            _risk("R-001", risk_distribution, 100.0, 200.0, 400.0, probability=0.3),
            _risk("R-002", "Uniform", 50.0, None, 90.0, probability=0.7),
        ],
    )


def test_15_a_mixed_three_cost_two_risk_model_runs_end_to_end() -> None:
    run, prepared, _ = _run(_mixed_model(), iterations=2000)

    assert len(run.total_nominal) == len(run.total_pv) == 2000
    assert len(run.diagnostics) == 3 + 2 * 2
    assert all(math.isfinite(total) for total in run.total_nominal)
    assert all(math.isfinite(total) for total in run.total_pv)
    assert len(run.result_digest) == 16
    assert run.summary.nominal.count == 2000 and run.summary.pv.count == 2000
    assert set(run.summary.nominal.percentiles) == set(prepared.ladder.ordered)


def test_16_a_model_with_no_beta_pert_runs_and_consumes_one_uniform_per_draw() -> None:
    """K: every component consumes exactly one uniform per iteration."""
    iterations = 1500
    model = _model(
        costs=[
            _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, quantity=1.0),
        ],
        risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.4)],
    )
    run, _, _ = _run(model, iterations=iterations)

    for record in run.diagnostics:
        assert record.uniforms_consumed == iterations, (
            f"{record.permanent_id} {record.role} consumed {record.uniforms_consumed}"
        )


def test_17_an_exact_friendly_no_beta_fixture_reproduces_hand_arithmetic() -> None:
    """L: every value is a power-of-two multiple, so the totals are exact.

    Cost lines are degenerate at exactly representable values with exactly
    representable quantities, and the risk is degenerate at a power of two, so
    each iteration total is one of exactly two numbers and both are stated here
    rather than compared against another implementation.
    """
    model = _model(
        costs=[
            _cost("CL-001", "Uniform", 256.0, None, 256.0, quantity=2.0),      # 512
            _cost("CL-002", "Triangular", 0.5, 0.5, 0.5, quantity=8.0),        # 4
            _cost("CL-003", "Beta-PERT", -32.0, -32.0, -32.0, quantity=0.25),  # -8
        ],
        risks=[_risk("R-001", "Uniform", 64.0, None, 64.0, probability=0.5)],
    )
    run, prepared, _ = _run(model, iterations=1000)

    assert set(run.total_nominal) == {508.0, 572.0}, sorted(set(run.total_nominal))
    assert set(run.total_pv) == {508.0, 572.0}
    assert prepared.deterministic_base.nominal == 508.0, "A excludes risks entirely"
    for record in run.diagnostics:
        expected = 1000 if record.role == "occurrence" else 0
        assert record.uniforms_consumed == expected, (record.permanent_id, record.role)


def test_18_a_beta_containing_model_runs_and_consumes_two_uniforms_per_attempt() -> None:
    """M: the Beta component's consumption is even and above one per iteration."""
    iterations = 2000
    run, _, _ = _run(_mixed_model("Beta-PERT"), iterations=iterations)

    for permanent_id in ("CL-003",):
        record = _diag(run, permanent_id, "value")
        assert record.uniforms_consumed % 2 == 0, "a Cheng attempt consumes two uniforms"
        assert record.uniforms_consumed >= 2 * iterations, "fewer attempts than iterations"
    severity = _diag(run, "R-001", "severity")
    assert severity.uniforms_consumed % 2 == 0
    assert severity.uniforms_consumed >= 2 * iterations


# ===========================================================================
# N, O, P - replay and seed scope
# ===========================================================================
def test_19_the_same_seed_replays_exactly() -> None:
    """No tolerance anywhere: the tuples and the digest are equal outright."""
    prepared, _ = _prepare(_mixed_model(), seed=777, iterations=1500)
    first = run_simulation(_ref(), prepared)
    second = run_simulation(_ref(), prepared)

    assert first.total_nominal == second.total_nominal
    assert first.total_pv == second.total_pv
    assert first.result_digest == second.result_digest
    assert [record.final_state for record in first.diagnostics] == [
        record.final_state for record in second.diagnostics
    ]
    # And again from a freshly prepared model, not just a re-run of one object.
    third, _, _ = _run(_mixed_model(), seed=777, iterations=1500)
    assert third.total_nominal == first.total_nominal
    assert third.result_digest == first.result_digest


def test_20_different_seeds_give_different_initial_streams_universally() -> None:
    """A: the UNIVERSAL half of the seed claim - about streams, not digests."""
    seen = {}
    for seed in (1, 2, 999, 12345, 2147483646):
        state = _ref().fixed_seed_to_state(seed)
        assert state.words not in seen.values(), f"seed {seed} reused a state"
        seen[seed] = state.words
    assert len(seen) == 5


def test_21_different_seeds_diverge_on_a_non_degenerate_fixture() -> None:
    """B: the FIXTURE-SCOPED half. Uncertainty must reach the retained total."""
    model = _mixed_model()
    digests = set()
    for seed in (1, 2, 999, 12345):
        run, _, _ = _run(model, seed=seed, iterations=1200)
        digests.add(run.result_digest)
    assert len(digests) == 4, "seeds collided on a fixture whose uncertainty is real"


def test_22_a_fully_degenerate_fixture_gives_the_same_digest_for_every_seed() -> None:
    """P: this is ACCEPTED behaviour, not a defect.

    A model with no uncertainty produces the same totals whatever the seed. The
    withdrawn universal claim "different seed -> different digest" would call this
    a failure; it is a correct answer about a model that has nothing to vary.
    """
    model = _model(
        costs=[_cost("CL-001", "Uniform", 250.0, 999.0, 250.0, quantity=2.0)],
        risks=[_risk("R-001", "Triangular", 7.0, 7.0, 7.0, probability=1.0)],
    )
    digests = set()
    totals = set()
    for seed in (1, 2, 999, 12345, 2147483646):
        run, _, _ = _run(model, seed=seed, iterations=1000)
        digests.add(run.result_digest)
        totals.add(run.total_nominal)
    assert len(digests) == 1 and len(totals) == 1
    assert next(iter(totals))[0] == 507.0


# ===========================================================================
# Q - row order invariance
# ===========================================================================
def test_23_physical_row_order_changes_nothing_at_all() -> None:
    """Both invariances at once: stream assignment AND accumulation order."""
    costs = [
        _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0),
        _cost("CL-002", "Uniform", 10.0, None, 20.0, quantity=1.0),
        _cost("CL-003", "Beta-PERT", 5.0, 7.0, 20.0, quantity=3.0),
    ]
    risks = [
        _risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.3),
        _risk("R-002", "Uniform", 50.0, None, 90.0, probability=0.7),
    ]
    orders = (
        (costs, risks),
        ([costs[2], costs[0], costs[1]], [risks[1], risks[0]]),
        ([costs[1], costs[2], costs[0]], list(reversed(risks))),
    )
    baseline = None
    for cost_order, risk_order in orders:
        run, prepared, _ = _run(_model(cost_order, risk_order), seed=4242, iterations=1200)
        identity = (
            [driver.permanent_id for driver in prepared.drivers],
            [(record.permanent_id, record.role, record.stream_index)
             for record in run.diagnostics],
            run.total_nominal,
            run.total_pv,
            run.result_digest,
        )
        if baseline is None:
            baseline = identity
            assert identity[0] == ["CL-001", "CL-002", "CL-003", "R-001", "R-002"]
        else:
            assert identity == baseline, "a physical row reorder changed the answer"


# ===========================================================================
# R - canonical accumulation order on a non-associative fixture
# ===========================================================================
_NON_ASSOCIATIVE = (1.0, 1.0e16, -1.0e16)
"""Three exactly representable contributions whose Double sum is order-dependent.

Left to right: `1 + 1e16` rounds to `1e16`, then `-1e16` gives `0.0`.
Reversed: `-1e16 + 1e16` is `0.0`, then `+1` gives `1.0`.

The difference is REQUIRED ONLY HERE. Reversing arbitrary contributions usually
changes nothing, and no test claims otherwise.
"""


def test_24_the_constructed_fixture_really_is_order_dependent() -> None:
    """The independent half: the oracle proves the two orders disagree."""
    canonical = safe_signed_sum(list(_NON_ASSOCIATIVE), "canonical")
    reversed_order = safe_signed_sum(list(reversed(_NON_ASSOCIATIVE)), "reversed")
    assert canonical == 0.0
    assert reversed_order == 1.0
    assert canonical != reversed_order


def test_25_the_engine_accumulates_in_canonical_order() -> None:
    """The engine half: with the fixture wired to canonical Permanent IDs, the
    retained total is the canonical sum and not the reversed one."""
    model = _model(
        costs=[
            _cost("CL-003", "Uniform", -1.0e16, None, -1.0e16, quantity=1.0),
            _cost("CL-001", "Uniform", 1.0, None, 1.0, quantity=1.0),
            _cost("CL-002", "Uniform", 1.0e16, None, 1.0e16, quantity=1.0),
        ]
    )
    run, prepared, _ = _run(model, iterations=1000)

    assert [driver.permanent_id for driver in prepared.drivers] == [
        "CL-001", "CL-002", "CL-003"
    ]
    assert set(run.total_nominal) == {0.0}, (
        f"the engine produced {sorted(set(run.total_nominal))}; the reversed order would "
        "have produced 1.0"
    )
    assert set(run.total_pv) == {0.0}


def test_26_costs_are_accumulated_before_risks() -> None:
    """The kind axis, on its own non-associative construction.

    Canonical order is `[cost 1e16, cost 1.0, risk -1e16]`: `1e16 + 1` ties to
    `1e16`, and `-1e16` closes it at `0.0`. Accumulating risks first gives
    `[-1e16, 1e16, 1.0]`, where the cancellation happens FIRST and the `1.0`
    survives. The two orders are `0.0` and `1.0`.

    Note the cost ordering matters here too, so the fixture is written with
    CL-001 as the large term - ascending Permanent ID puts it first, which is
    what makes the small term the one that is absorbed.
    """
    model = _model(
        costs=[
            _cost("CL-001", "Uniform", 1.0e16, None, 1.0e16, quantity=1.0),
            _cost("CL-002", "Uniform", 1.0, None, 1.0, quantity=1.0),
        ],
        risks=[_risk("R-001", "Uniform", -1.0e16, None, -1.0e16, probability=1.0)],
    )
    run, prepared, _ = _run(model, iterations=1000)

    canonical = safe_signed_sum([1.0e16, 1.0, -1.0e16], "cost first")
    risk_first = safe_signed_sum([-1.0e16, 1.0e16, 1.0], "risk first")
    assert canonical == 0.0
    assert risk_first == 1.0, "the control is vacuous"

    assert [driver.permanent_id for driver in prepared.drivers] == [
        "CL-001", "CL-002", "R-001"
    ]
    assert set(run.total_nominal) == {0.0}, (
        f"the engine produced {sorted(set(run.total_nominal))}; accumulating risks first "
        "would have produced 1.0"
    )


def test_27_nominal_and_pv_are_independent_accumulators() -> None:
    """PV is never derived by discounting the sampled nominal total."""
    model = _model(
        costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0)],
        risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.5)],
        discount_rate=0.10,
    )
    prepared, result = _prepare(
        _model(
            costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0)],
            risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.5)],
        ),
        iterations=1000,
    )
    run = run_simulation(_ref(), prepared)
    driver = prepared.cost_drivers[0]
    # With one applied project year at the base year the two factors coincide,
    # so the check that matters is structural: each total is formed from its own
    # factor, driver by driver, and the PV total is never the nominal total
    # multiplied by a discount factor after the fact.
    assert driver.knom == 1.0 and driver.kpv == 1.0
    assert run.total_pv == run.total_nominal
    assert result.discount_factors[1] == 1.0


def test_28_pv_uses_kpv_and_nominal_uses_knom() -> None:
    """A multi-year model where the two factors genuinely differ."""
    model = CalculationModel(
        timeline=AppliedTimeline(2026, 2027, 2),
        discount_rate=0.10,
        fx_rows=(FxRow("SAR", 1),),
        inflation_rates={"Standard": {2027: 0.0, 2028: 0.0}},
        cost_drivers=(
            CostDriver("CL-001", "Uniform", "SAR", "Standard",
                       100.0, None, 100.0, (0.5, 0.5), quantity=1.0),
        ),
    )
    run, prepared, _ = _run(model, iterations=1000)
    driver = prepared.cost_drivers[0]

    assert driver.knom != driver.kpv, "the fixture does not separate the two factors"
    assert run.total_nominal[0] == 100.0 * driver.knom
    assert run.total_pv[0] == 100.0 * driver.kpv
    assert run.total_pv[0] != run.total_nominal[0]


# ===========================================================================
# section 12 - retained output shape
# ===========================================================================
def test_29_retained_output_is_two_immutable_arrays_in_iteration_order() -> None:
    run, _, _ = _run(_mixed_model(), iterations=1200)

    assert isinstance(run.total_nominal, tuple) and isinstance(run.total_pv, tuple)
    assert len(run.total_nominal) == len(run.total_pv) == 1200
    # Statistics sorted copies; the originals are untouched and are NOT sorted.
    assert list(run.total_nominal) != sorted(run.total_nominal), (
        "the retained array came back sorted"
    )
    assert run.summary.nominal.minimum == min(run.total_nominal)
    assert run.summary.nominal.maximum == max(run.total_nominal)


def test_30_no_per_driver_per_iteration_matrix_is_retained() -> None:
    """The retained stochastic arrays are exactly two, whatever the driver count."""
    iterations = 1000
    small, _, _ = _run(_model(costs=[_cost("CL-001")]), iterations=iterations)
    large, _, _ = _run(_mixed_model(), iterations=iterations)

    def stochastic_arrays(run) -> list[tuple]:
        return [
            value for value in vars(run).values()
            if isinstance(value, tuple) and len(value) == iterations
        ]

    assert len(stochastic_arrays(small)) == 2
    assert len(stochastic_arrays(large)) == 2, "a per-driver sample matrix survived the run"
    assert len(large.diagnostics) == 7, "diagnostics are per component, not per iteration"
    for record in large.diagnostics:
        assert not hasattr(record, "samples")


# ===========================================================================
# S - the result digest
# ===========================================================================
def test_31_every_step0_digest_vector_reproduces_exactly() -> None:
    """All seven retained cases, on the canonical stream AND on the digest."""
    cases = json.loads(
        (EVIDENCE / "vectors" / "digest_vectors.json").read_text(encoding="utf-8")
    )["cases"]
    assert len(cases) == 7
    labels = {case["label"] for case in cases}
    assert labels == {
        "base", "reversed_iteration_order", "nominal_and_pv_swapped",
        "one_iteration_dropped", "one_ulp_perturbation", "version_2", "empty",
    }
    for case in cases:
        nominal = [float(value) for value in case["totals_nominal"]]
        pv = [float(value) for value in case["totals_pv"]]
        assert result_digest_stream(case["version"], nominal, pv) == case["stream"], case["label"]
        assert result_digest(case["version"], nominal, pv) == case["digest"], case["label"]


def test_32_the_digest_helper_is_standalone_and_frames_an_empty_sequence() -> None:
    """A real run can never be empty - the business minimum is at least 1000 - but
    the framing vector must still be reproducible without an engine."""
    empty = result_digest_stream(1, [], [])
    assert empty == 'S7:PCCM-RDI1:1S6:RESULTI1:0'
    assert result_digest(1, [], []) == "12ED977808313D71"


def test_33_the_engine_digest_equals_the_standalone_digest_of_its_arrays() -> None:
    run, prepared, _ = _run(_mixed_model(), iterations=1100)
    assert run.result_digest == result_digest(
        prepared.sim_method_version, run.total_nominal, run.total_pv
    )
    assert run.sim_method_version == _sim().sim_method_version == 1


def test_34_the_digest_grammar_matches_the_contract() -> None:
    validate_result_digest_contract(_sim())
    block = _sim().raw["result_digest"]
    assert block["stream_tag"] == "PCCM-RD"
    assert block["version_field_source"] == "sim_method_version"
    assert block["samples_sorted_for_digest"] is False


def test_35_a_mismatched_array_pair_is_refused() -> None:
    try:
        result_digest(1, [1.0, 2.0], [1.0])
    except SimOracleError as error:
        assert "disagree in length" in str(error)
        return
    raise AssertionError("mismatched retained arrays were accepted")


# ===========================================================================
# T - Type-7 percentile hand vectors
# ===========================================================================
def test_36_type_7_percentiles_match_hand_derived_vectors() -> None:
    """`h = (n-1)p`, `lo = floor(h)`, `hi = min(lo+1, n-1)`, `f = h - lo`,
    `Px = (1-f)x[lo] + f x[hi]`, worked through by hand for n = 1, 2, 3, 4, 10."""
    cases = (
        # n = 1: every percentile is the single observation
        ([5.0], 0.0, 5.0), ([5.0], 0.5, 5.0), ([5.0], 1.0, 5.0),
        # n = 2: h = p, so p is the interpolation weight outright
        ([10.0, 20.0], 0.0, 10.0),
        ([10.0, 20.0], 0.5, 15.0),
        ([10.0, 20.0], 1.0, 20.0),
        ([10.0, 20.0], 0.25, 12.5),
        # n = 3: h = 2p. p = 0.5 -> h = 1 exactly, an integral h.
        ([10.0, 20.0, 60.0], 0.5, 20.0),
        ([10.0, 20.0, 60.0], 0.25, 15.0),
        ([10.0, 20.0, 60.0], 0.75, 40.0),
        # n = 4: h = 3p. p = 1/3 -> h = 1 exactly.
        ([1.0, 2.0, 3.0, 4.0], 1.0 / 3.0, 2.0),
        ([1.0, 2.0, 3.0, 4.0], 0.5, 2.5),
        ([1.0, 2.0, 3.0, 4.0], 0.9, 3.7),
        ([1.0, 2.0, 3.0, 4.0], 0.0, 1.0),
        ([1.0, 2.0, 3.0, 4.0], 1.0, 4.0),
        # n = 10: h = 9p. p = 0.5 -> h = 4.5, midway between x[4] and x[5].
        (list(range(1, 11)), 0.5, 5.5),
        (list(range(1, 11)), 0.1, 1.9),
        (list(range(1, 11)), 0.9, 9.1),
        ([float(v) for v in range(1, 11)], 1.0 / 3.0, 4.0),
    )
    for values, p, expected in cases:
        got = percentile_type7([float(v) for v in values], p)
        assert got == expected or math.isclose(got, expected, rel_tol=0.0, abs_tol=1e-12), (
            values, p, got, expected
        )


def test_37_the_percentile_input_is_never_reordered() -> None:
    values = [9.0, 1.0, 5.0, 3.0]
    original = list(values)
    percentile_type7(values, 0.5)
    describe(values, (("P50", 0.5),))
    assert values == original, "the caller's sequence was sorted in place"


def test_38_type_7_interpolates_convexly_at_extreme_magnitudes() -> None:
    """The unsafe difference form overflows where the convex form is exact."""
    low, high = -1.7e308, 1.7e308
    for p in (0.25, 0.5, 0.75):
        got = percentile_type7([low, high], p)
        assert math.isfinite(got)
        assert low <= got <= high
        unsafe = low + p * (high - low)
        assert not math.isfinite(unsafe), "the control is vacuous"
    assert percentile_type7([low, high], 0.5) == 0.0
    assert percentile_type7([low, high], 0.0) == low
    assert percentile_type7([low, high], 1.0) == high


# ===========================================================================
# U - mean and sample standard deviation hand vectors
# ===========================================================================
def test_39_sample_mean_and_sd_match_hand_derived_values() -> None:
    """`SD` uses `n - 1`. The population divisor `n` is a different number and
    each case states both so the substitution cannot pass unnoticed."""
    cases = (
        ([2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0], 5.0, math.sqrt(32.0 / 7.0), 2.0),
        ([1.0, 2.0, 3.0, 4.0], 2.5, math.sqrt(5.0 / 3.0), math.sqrt(1.25)),
        ([10.0, 10.0, 10.0], 10.0, 0.0, 0.0),
        ([-3.0, 3.0], 0.0, math.sqrt(18.0 / 1.0), 3.0),
    )
    for values, mean, sample_sd, population_sd in cases:
        assert sample_mean(values) == mean, values
        got = sample_standard_deviation(values)
        assert math.isclose(got, sample_sd, rel_tol=1e-15), (values, got, sample_sd)
        if sample_sd != 0.0:
            assert not math.isclose(got, population_sd, rel_tol=1e-9), (
                f"{values}: the sample and population deviations are indistinguishable"
            )


def test_40_the_sample_mean_is_scale_safe() -> None:
    """A representable mean is produced where the running sum does not exist."""
    values = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]
    assert sample_mean(values) == 8.5e307

    naive = 0.0
    for value in values:
        naive += value
    assert math.isinf(naive), "the control is vacuous: the naive sum did not overflow"

    assert sample_mean([0.0, 0.0, 0.0]) == 0.0
    assert sample_mean([5e-324, 5e-324]) == 5e-324
    assert sample_mean([1.0]) == 1.0


def test_41_the_sample_sd_is_scale_safe_and_refuses_below_two_observations() -> None:
    values = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]
    got = sample_standard_deviation(values)
    assert math.isfinite(got) and got > 1.6e308

    mean = sample_mean(values)
    assert math.isinf(values[0] - mean), "an unguarded Welford delta would not have overflowed"
    assert math.isinf(sum(value * value for value in values)), (
        "a naive sum of squares would not have overflowed"
    )

    for short in ([], [1.0]):
        try:
            sample_standard_deviation(short)
        except SimStatsError as error:
            assert "n - 1" in str(error)
            continue
        raise AssertionError(f"{short!r} produced a sample standard deviation")


def test_42_the_run_statistics_are_the_helpers_applied_to_the_retained_arrays() -> None:
    run, _, _ = _run(_mixed_model(), iterations=1500)
    for measure, values in (("nominal", run.total_nominal), ("pv", run.total_pv)):
        stats = run.summary.measure(measure)
        assert stats.mean == sample_mean(values, measure)
        assert stats.sample_standard_deviation == sample_standard_deviation(values, measure)
        assert stats.minimum == min(values) and stats.maximum == max(values)
        for label, p in run.summary.ladder.points:
            assert stats.percentiles[label] == percentile_type7(values, p)


# ===========================================================================
# V - the reported percentile ladder
# ===========================================================================
def test_43_the_ladder_is_resolved_from_its_owners_and_is_complete() -> None:
    ladder = resolve_percentile_ladder(_sim(), _inputs())

    assert ladder.ordered == (
        "P10", "P50", "P55", "P60", "P65", "P70", "P75", "P80", "P85", "P90", "P95"
    )
    assert len(ladder.ordered) == 11
    assert ladder.fixed == ("P10",)
    assert ladder.selectable == (
        "P50", "P55", "P60", "P65", "P70", "P75", "P80", "P85", "P90", "P95"
    )
    assert "P10" not in ladder.selectable, "P10 must not be selectable"
    assert ladder.headline == ("P10", "P50", "P70", "P90")
    assert dict(ladder.points)["P10"] == 0.10
    assert dict(ladder.points)["P95"] == 0.95


def test_44_the_ladder_is_read_from_the_contracts_not_restated() -> None:
    """The selectable values are exactly the config table's rows."""
    table = next(
        entry for entry in _inputs().config_tables if entry.key == "confidence_levels"
    )
    assert resolve_percentile_ladder(_sim(), _inputs()).selectable == tuple(
        str(row[0]) for row in table.seed_rows
    )
    assert tuple(_sim().raw["statistics"]["fixed_nonselectable_percentiles"]) == ("P10",)


def test_45_every_ladder_value_is_stored_for_both_measures() -> None:
    run, prepared, _ = _run(_mixed_model(), iterations=1200)

    for measure in ("nominal", "pv"):
        stats = run.summary.measure(measure)
        assert tuple(sorted(stats.percentiles)) == tuple(sorted(prepared.ladder.ordered))
        ordered_values = [stats.percentiles[label] for label in prepared.ladder.ordered]
        assert ordered_values == sorted(ordered_values), (
            "the ladder is not monotonic, so the labels and values are misaligned"
        )
    for label in prepared.ladder.headline:
        assert label in run.summary.nominal.percentiles


# ===========================================================================
# W - the Selected Confidence Level is reporting only
# ===========================================================================
def test_46_selected_confidence_level_never_enters_execution() -> None:
    """It is not an argument to preparation or to the run - there is nowhere for
    it to enter - and every reported value is identical across every selection."""
    run, prepared, result = _run(_mixed_model(), iterations=1500)
    base = deterministic_base_of(result)

    reference_digest = run.result_digest
    reference_mean = run.summary.nominal.mean
    reference_ladder = dict(run.summary.nominal.percentiles)

    seen = {}
    for level in prepared.ladder.selectable:
        contingency = contingency_at(run.summary, level, base)
        assert run.result_digest == reference_digest
        assert run.summary.nominal.mean == reference_mean
        assert dict(run.summary.nominal.percentiles) == reference_ladder
        seen[level] = contingency.nominal
    assert len(set(seen.values())) > 1, "the control is vacuous: every level gave the same answer"

    import inspect

    for function in (prepare_simulation, run_simulation):
        parameters = set(inspect.signature(function).parameters)
        assert not any("confidence" in name for name in parameters), (
            f"{function.__name__} accepts a confidence-level argument: {parameters}"
        )


def test_47_p10_is_reported_but_cannot_be_selected() -> None:
    run, prepared, result = _run(_mixed_model(), iterations=1200)

    assert "P10" in run.summary.nominal.percentiles
    try:
        contingency_at(run.summary, "P10", deterministic_base_of(result))
    except SimOracleError as error:
        assert "not selectable" in str(error)
    else:
        raise AssertionError("P10 was accepted as a contingency selector")

    for bad in ("P42", "p50", "P100", "", "median"):
        try:
            contingency_at(run.summary, bad, deterministic_base_of(result))
        except SimOracleError:
            continue
        raise AssertionError(f"{bad!r} was accepted as a confidence level")


# ===========================================================================
# X - contingency
# ===========================================================================
def test_48_contingency_is_the_selected_percentile_minus_the_phase5_base_a() -> None:
    run, prepared, result = _run(_mixed_model(), iterations=1500)
    base = deterministic_base_of(result)

    assert base.nominal == result.totals.a_nom and base.pv == result.totals.a_pv
    assert base == prepared.deterministic_base

    for level in ("P50", "P80", "P95"):
        contingency = contingency_at(run.summary, level, base)
        assert contingency.nominal == run.summary.nominal.percentiles[level] - base.nominal
        assert contingency.pv == run.summary.pv.percentiles[level] - base.pv
        assert contingency.base_nominal == result.totals.a_nom

    # The forbidden baselines are all different numbers on this fixture.
    assert run.summary.nominal.mean != base.nominal
    assert result.totals.e_nom != base.nominal
    assert result.totals.a_nom + result.totals.d_nom != base.nominal


def test_49_a_negative_contingency_is_reported_not_clamped() -> None:
    """A selected percentile below A gives a negative contingency, which is a real
    statement about the model."""
    model = _model(
        costs=[_cost("CL-001", "Triangular", 10.0, 20.0, 30.0, quantity=1.0)]
    )
    run, _, result = _run(model, iterations=2000)
    base = DeterministicBase(1.0e6, 1.0e6)

    contingency = contingency_at(run.summary, "P50", base)
    assert contingency.nominal < 0.0
    assert contingency.nominal == run.summary.nominal.percentiles["P50"] - 1.0e6
    assert contingency.pv < 0.0


# ===========================================================================
# section 25 - analytical expectation cross-check
# ===========================================================================
def _agrees_with_expectation(run, expected: float, measure: str, allowance_sigma: float = 4.0):
    """A TEST-OWNED statistical allowance, in standard errors of the mean.

    This is evidence that the simulation TARGETS the accepted analytical
    expectation. It is not a runtime tolerance, not a digest tolerance and not a
    field of `sim_contract.yaml`; a finite Monte Carlo mean is never asserted
    equal to an expectation.
    """
    stats = run.summary.measure(measure)
    standard_error = stats.sample_standard_deviation / math.sqrt(stats.count)
    allowance = allowance_sigma * standard_error + 1e-9 * abs(expected)
    difference = abs(stats.mean - expected)
    assert difference <= allowance, (
        f"{measure}: simulation mean {stats.mean} differs from the analytical expectation "
        f"{expected} by {difference}, beyond the {allowance_sigma}-sigma test allowance "
        f"{allowance}"
    )


def test_50_a_stochastic_cost_line_targets_its_analytical_expectation() -> None:
    model = _model(costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0)])
    run, prepared, result = _run(model, seed=20260825, iterations=20000)

    assert result.totals.e_nom == 220.0, "the Triangular mean is (80 + 100 + 150) / 3"
    _agrees_with_expectation(run, prepared.analytical_expectation.nominal, "nominal")
    _agrees_with_expectation(run, prepared.analytical_expectation.pv, "pv")


def test_51_a_stochastic_risk_occurrence_targets_its_expected_value() -> None:
    model = _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.4)])
    run, prepared, result = _run(model, seed=20260825, iterations=20000)

    assert result.totals.e_nom == 0.4 * 250.0
    _agrees_with_expectation(run, prepared.analytical_expectation.nominal, "nominal")
    _agrees_with_expectation(run, prepared.analytical_expectation.pv, "pv")


def test_52_a_beta_pert_severity_targets_its_expected_value() -> None:
    model = _model(risks=[_risk("R-001", "Beta-PERT", 100.0, 200.0, 400.0, probability=0.5)])
    run, prepared, result = _run(model, seed=20260825, iterations=20000)

    assert math.isclose(result.totals.e_nom, 0.5 * (100.0 + 4 * 200.0 + 400.0) / 6.0)
    _agrees_with_expectation(run, prepared.analytical_expectation.nominal, "nominal")


def test_53_a_mixed_model_targets_its_analytical_expectation() -> None:
    run, prepared, result = _run(_mixed_model("Beta-PERT"), seed=20260825, iterations=20000)

    assert prepared.analytical_expectation.nominal == result.totals.e_nom
    _agrees_with_expectation(run, prepared.analytical_expectation.nominal, "nominal")
    _agrees_with_expectation(run, prepared.analytical_expectation.pv, "pv")
    assert prepared.analytical_expectation.nominal != prepared.deterministic_base.nominal


# ===========================================================================
# section 5 - iteration count pre-flight
# ===========================================================================
def test_54_the_business_minimum_is_read_from_input_contract() -> None:
    assert business_minimum_iterations(_inputs()) == 1000
    spec = _inputs().inputs["monte_carlo_iterations"]
    assert spec.validation["formula1"] == "1000"
    assert spec.validation["kind"] == "whole"


def test_55_the_technical_ceiling_is_read_from_sim_contract() -> None:
    assert _sim().max_iterations_representable == 1048543
    assert _sim().reserved_rows_h == 33


def test_56_an_inadmissible_iteration_count_is_refused() -> None:
    """Boundary validation, proved WITHOUT allocating a million-row simulation."""
    assert validate_iterations(_sim(), _inputs(), 1000) == 1000
    assert validate_iterations(_sim(), _inputs(), 1048543) == 1048543

    for bad, fragment in (
        (999, "business minimum"),
        (0, "business minimum"),
        (-1, "business minimum"),
        (1048544, "TECHNICAL ceiling"),
        (2_000_000, "TECHNICAL ceiling"),
        (1000.0, "whole integer"),
        ("1000", "whole integer"),
        (True, "whole integer"),
        (None, "whole integer"),
    ):
        try:
            validate_iterations(_sim(), _inputs(), bad)
        except SimOracleError as error:
            assert fragment in str(error), (bad, str(error))
            continue
        raise AssertionError(f"iteration count {bad!r} was accepted")


def test_57_the_refusal_precedes_stream_construction_and_any_draw() -> None:
    """Nothing is allocated, no stream is built and no uniform is drawn."""
    model = _mixed_model()
    for bad in (999, 1048544):
        try:
            _prepare(model, iterations=bad)
        except SimOracleError:
            continue
        raise AssertionError(f"preparation proceeded with {bad} iterations")

    # The order is structural, not incidental: a model that would itself be
    # refused still reports the ITERATION problem first.
    broken = _model(costs=[_cost("CL-001", "Triangular", 150.0, 100.0, 80.0)])
    try:
        _prepare(broken, iterations=999)
    except SimOracleError as error:
        assert "business minimum" in str(error)
        return
    raise AssertionError("a model with an inverted triple was prepared at 999 iterations")


# ===========================================================================
# section 4 - effective seed, not nonce lifecycle
# ===========================================================================
def test_58_the_engine_takes_an_effective_seed_and_persists_nothing() -> None:
    reference = _ref()
    assert effective_seed_from_nonce(reference, 0) == 1
    assert effective_seed_from_nonce(reference, 1) == reference.auto_multiplier
    # Pure: calling it twice with the same nonce gives the same seed, and no
    # counter moved anywhere.
    assert effective_seed_from_nonce(reference, 7) == effective_seed_from_nonce(reference, 7)

    # The engine's entry points take an effective seed, not a nonce and not a
    # workbook counter: there is nowhere for the transactional lifecycle to
    # enter. (That no persistence CALL exists anywhere in the module is proved
    # semantically by the Step-4 scope guards, which parse the source rather
    # than searching its prose.)
    import inspect

    parameters = set(inspect.signature(prepare_simulation).parameters)
    assert "effective_seed" in parameters
    assert not any("nonce" in name for name in parameters), parameters
    assert not any("nonce" in name for name in inspect.signature(run_simulation).parameters)


# ===========================================================================
# section 27, 28 - error semantics and no partial result
# ===========================================================================
def _fails_mid_run(quantity: float) -> CalculationModel:
    """A model Phase 5 accepts and the simulation cannot always evaluate.

    The stochastic line is symmetric about zero, so its DETERMINISTIC central
    value and its mean are both exactly zero and every Phase-5 total is zero -
    the model is perfectly calculable. Only a SAMPLE far enough from the centre
    overflows when multiplied by `quantity`, which is what puts the failure
    inside the iteration loop where Step 4 has to handle it.
    """
    return _model(
        costs=[
            _cost("CL-001", "Uniform", 5.0, None, 5.0, quantity=1.0),
            _cost("CL-002", "Uniform", -1.0e308, None, 1.0e308, quantity=quantity),
        ]
    )


def test_59_a_failing_iteration_returns_no_partial_result() -> None:
    """The engine gets partway and hands back NOTHING - not a short tuple.

    With this quantity roughly one iteration in ten is unrepresentable, so the
    run completes iteration 1 and fails at iteration 2. A design that returned
    what it had would have a one-element result to offer; this one raises.
    """
    prepared, result = _prepare(_fails_mid_run(2.0), iterations=1000)
    assert result.totals.a_nom == 5.0, "the model itself is perfectly calculable"

    try:
        run_simulation(_ref(), prepared)
    except NumericalRangeRefusal as error:
        message = str(error)
        assert message.startswith("iteration 2:"), message
        assert not hasattr(error, "total_nominal"), "a partial result rode out on the exception"
    else:
        raise AssertionError("an unrepresentable contribution was accepted")

    # Deterministic: the same prepared model fails at the same iteration again,
    # and still returns nothing.
    try:
        run_simulation(_ref(), prepared)
    except NumericalRangeRefusal as error:
        assert str(error).startswith("iteration 2:")
        return
    raise AssertionError("the second attempt succeeded where the first refused")


def test_60_a_refusal_names_the_iteration_the_driver_and_the_stage() -> None:
    prepared, _ = _prepare(_fails_mid_run(1.0e10), iterations=1000)
    try:
        run_simulation(_ref(), prepared)
    except CalculationRefusal as error:
        message = str(error)
        assert "iteration 1" in message
        assert "CL-002" in message, "the refusal does not name the driver"
        assert "Cost Line" in message, "the refusal does not name the driver kind"
        assert "contribution" in message, "the refusal does not name the numerical stage"
        assert isinstance(error, NumericalRangeRefusal)
        return
    raise AssertionError("an overflowing product was accepted")


def test_61_the_phase5_refusal_hierarchy_is_preserved_not_replaced() -> None:
    """A model the calculation refuses is refused by ITS class, not by
    `SimOracleError`."""
    from pccm_builder.calc_numeric import ModelInputRefusal

    inverted = _model(costs=[_cost("CL-001", "Triangular", 150.0, 100.0, 80.0)])
    try:
        _prepare(inverted, iterations=1000)
    except ModelInputRefusal:
        pass
    else:
        raise AssertionError("an inverted three-point triple was prepared")

    bad_probability = _model(risks=[_risk("R-001", probability=1.5)])
    try:
        _prepare(bad_probability, iterations=1000)
    except ModelInputRefusal:
        pass
    else:
        raise AssertionError("a probability outside [0, 1] was prepared")


def test_62_run_simulation_refuses_something_that_is_not_a_prepared_model() -> None:
    for bad in (None, 42, "model", object()):
        try:
            run_simulation(_ref(), bad)
        except SimOracleError:
            continue
        raise AssertionError(f"{bad!r} was accepted as a prepared model")


# ===========================================================================
# Y - the extreme-domain pipeline
# ===========================================================================
def test_63_large_positive_endpoints_near_double_max() -> None:
    """Family 1."""
    model = _model(costs=[_cost("CL-001", "Uniform", 1.0e307, None, 1.5e307, quantity=1.0)])
    run, _, _ = _run(model, iterations=1000)

    assert all(1.0e307 <= total <= 1.5e307 for total in run.total_nominal)
    assert math.isfinite(run.summary.nominal.mean)
    assert math.isfinite(run.summary.nominal.sample_standard_deviation)
    assert math.isfinite(run.summary.nominal.percentiles["P95"])


def test_64_large_negative_endpoints() -> None:
    """Family 2. No positivity rule exists and none is invented."""
    model = _model(costs=[_cost("CL-001", "Uniform", -1.5e307, None, -1.0e307, quantity=1.0)])
    run, _, _ = _run(model, iterations=1000)

    assert all(-1.5e307 <= total <= -1.0e307 for total in run.total_nominal)
    assert run.summary.nominal.mean < 0.0
    assert math.isfinite(run.summary.nominal.sample_standard_deviation)


def test_65_support_crossing_zero() -> None:
    """Family 3."""
    model = _model(costs=[_cost("CL-001", "Triangular", -1.0e307, 0.0, 1.0e307, quantity=1.0)])
    run, _, _ = _run(model, iterations=2000)

    assert min(run.total_nominal) < 0.0 < max(run.total_nominal)
    assert math.isfinite(run.summary.nominal.mean)
    assert math.isfinite(run.summary.nominal.percentiles["P10"])


def test_66_the_largest_authorised_opposite_sign_accumulation() -> None:
    """Family 4: the partial sum leaves Double range and the total does not.

    `1.5e308 + 1.5e308` is infinite, so a naive left-to-right accumulation
    refuses. The exact total is `1.5e308`, which exists, and the accepted
    `safe_signed_sum` rescue produces it.
    """
    model = _model(
        costs=[
            _cost("CL-001", "Uniform", 1.5e308, None, 1.5e308, quantity=1.0),
            _cost("CL-002", "Uniform", 1.5e308, None, 1.5e308, quantity=1.0),
            _cost("CL-003", "Uniform", -1.5e308, None, -1.5e308, quantity=1.0),
        ]
    )
    naive = 0.0
    for term in (1.5e308, 1.5e308, -1.5e308):
        naive += term
    assert math.isinf(naive), "the control is vacuous"

    run, _, _ = _run(model, iterations=1000)
    assert set(run.total_nominal) == {1.5e308}, (
        "the accumulation refused a representable total, or produced the wrong one"
    )

    # The retained totals are exact AND SO ARE THEIR STATISTICS. Every retained
    # total is the same Double, so the distribution has no dispersion, and the
    # constant-sample invariant returns the exact mean and exactly zero rather
    # than accumulating the value a thousand times and rediscovering
    # 1.4999999999999677e308 with a standard deviation of 3.2e294.
    assert run.summary.nominal.mean == 1.5e308
    assert run.summary.nominal.sample_standard_deviation == 0.0
    assert set(run.summary.nominal.percentiles.values()) == {1.5e308}


def test_67_tiny_and_subnormal_scale_values() -> None:
    """Family 5."""
    model = _model(
        costs=[_cost("CL-001", "Uniform", 5e-324, None, 1e-320, quantity=1.0)]
    )
    run, _, _ = _run(model, iterations=1000)

    assert all(0.0 < total <= 1e-320 for total in run.total_nominal)
    assert 0.0 < run.summary.nominal.mean <= 1e-320
    assert math.isfinite(run.summary.nominal.sample_standard_deviation)


def test_68_endpoint_mode_beta_reaches_bc_by_the_rule() -> None:
    """Family 7: `m = a` and `m = b` are BC, not special cases."""
    for minimum, mode, maximum in ((0.0, 0.0, 100.0), (0.0, 100.0, 100.0)):
        model = _model(
            costs=[_cost("CL-001", "Beta-PERT", minimum, mode, maximum, quantity=1.0)]
        )
        run, prepared, _ = _run(model, iterations=1000)
        shape = prepared.cost_drivers[0].beta_shape
        assert shape is not None and shape.dispatch == "BC", (minimum, mode, maximum)
        assert all(minimum <= total <= maximum for total in run.total_nominal)
        assert _diag(run, "CL-001", "value").uniforms_consumed % 2 == 0


def test_69_statistics_over_near_max_opposite_sign_totals() -> None:
    """Families 8 and 9, on the statistics helpers directly.

    A retained array can legally hold totals of opposite sign near Double
    maximum. Their mean and sample deviation exist; a naive sum, a naive sum of
    squares and an unguarded deviation do not.
    """
    totals = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]
    stats = describe(totals, (("P10", 0.10), ("P50", 0.50), ("P90", 0.90)))

    assert stats.mean == 8.5e307
    assert math.isfinite(stats.sample_standard_deviation)
    assert stats.minimum == -1.7e308 and stats.maximum == 1.7e308
    assert stats.percentiles["P50"] == 1.7e308
    assert math.isfinite(stats.percentiles["P10"])
    assert all(math.isfinite(value) for value in stats.percentiles.values())


def test_70_contingency_subtraction_at_the_extremes() -> None:
    """Family 10: representable differences are produced; an unrepresentable one
    is refused, naming the stage, rather than returned as infinity."""
    ladder = resolve_percentile_ladder(_sim(), _inputs())
    percentiles = {label: 1.0e308 for label in ladder.ordered}
    stats = MeasureStatistics(
        count=1000, mean=1.0e308, sample_standard_deviation=0.0,
        minimum=1.0e308, maximum=1.0e308, percentiles=percentiles,
    )
    summary = SimulationSummary(nominal=stats, pv=stats, ladder=ladder)

    fine = contingency_at(summary, "P90", DeterministicBase(-5.0e307, -5.0e307))
    assert fine.nominal == 1.5e308 and math.isfinite(fine.nominal)

    try:
        contingency_at(summary, "P90", DeterministicBase(-1.0e308, -1.0e308))
    except NumericalRangeRefusal as error:
        assert "contingency" in str(error)
        return
    raise AssertionError("an unrepresentable contingency was returned instead of refused")


def test_71_no_silent_non_finite_value_reaches_a_retained_array() -> None:
    run, _, _ = _run(_mixed_model("Beta-PERT"), iterations=2000)

    for values in (run.total_nominal, run.total_pv):
        assert all(math.isfinite(total) for total in values)
    stats = run.summary.nominal
    assert math.isfinite(stats.mean) and math.isfinite(stats.sample_standard_deviation)
    assert all(math.isfinite(value) for value in stats.percentiles.values())


# ===========================================================================
# Z - no evidence, runtime or workbook dependency
# ===========================================================================
def test_72_the_engine_runs_with_no_file_access_at_all() -> None:
    """Contracts and models are supplied before the loop; the run opens nothing.

    `open` is replaced for the duration of the call, so any file access - the
    Step-0 evidence included - fails loudly instead of passing unnoticed.
    """
    import builtins
    import io

    prepared, result = _prepare(_mixed_model(), iterations=1200)
    opened: list[str] = []

    def refuse(*args, **kwargs):
        opened.append(repr(args[:1]))
        raise AssertionError(f"the engine opened a file: {args[:1]}")

    real_open, real_io_open = builtins.open, io.open
    builtins.open, io.open = refuse, refuse
    try:
        run = run_simulation(_ref(), prepared)
        contingency_at(run.summary, "P80", deterministic_base_of(result))
    finally:
        builtins.open, io.open = real_open, real_io_open

    assert not opened
    assert len(run.total_nominal) == 1200


def test_73_the_engine_imports_no_evidence_and_no_workbook_machinery() -> None:
    import ast

    source = (PCCM_ROOT / "builder" / "pccm_builder" / "sim_oracle.py").read_text(
        encoding="utf-8"
    )
    tree = ast.parse(source)
    imported: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(alias.name for alias in node.names)
        elif isinstance(node, ast.ImportFrom):
            imported.add(node.module or "")
    for module in imported:
        assert "evidence" not in module, module
        assert "openpyxl" not in module, module
        assert "win32" not in module and "comtypes" not in module, module
        assert "workbook" not in module, module
        assert "numpy" not in module and "random" not in module, module


def test_74_the_package_re_exports_the_step4_surface() -> None:
    """The public API is reachable from `pccm_builder`, and is the same object."""
    import pccm_builder
    from pccm_builder import sim_oracle, sim_stats

    surface = {
        "prepare_simulation": sim_oracle.prepare_simulation,
        "run_simulation": sim_oracle.run_simulation,
        "validate_iterations": sim_oracle.validate_iterations,
        "business_minimum_iterations": sim_oracle.business_minimum_iterations,
        "resolve_percentile_ladder": sim_oracle.resolve_percentile_ladder,
        "effective_seed_from_nonce": sim_oracle.effective_seed_from_nonce,
        "result_digest": sim_oracle.result_digest,
        "result_digest_stream": sim_oracle.result_digest_stream,
        "rng_reference_signature": sim_oracle.rng_reference_signature,
        "validate_result_digest_contract": sim_oracle.validate_result_digest_contract,
        "contingency_at": sim_oracle.contingency_at,
        "deterministic_base_of": sim_oracle.deterministic_base_of,
        "SimOracleError": sim_oracle.SimOracleError,
        "SimulationResult": sim_oracle.SimulationResult,
        "SimulationSummary": sim_oracle.SimulationSummary,
        "PreparedSimulationModel": sim_oracle.PreparedSimulationModel,
        "PreparedSimulationDriver": sim_oracle.PreparedSimulationDriver,
        "PercentileLadder": sim_oracle.PercentileLadder,
        "DeterministicBase": sim_oracle.DeterministicBase,
        "AnalyticalExpectation": sim_oracle.AnalyticalExpectation,
        "ComponentDiagnostics": sim_oracle.ComponentDiagnostics,
        "Contingency": sim_oracle.Contingency,
        "sample_mean": sim_stats.sample_mean,
        "sample_standard_deviation": sim_stats.sample_standard_deviation,
        "percentile_type7": sim_stats.percentile_type7,
        "describe": sim_stats.describe,
        "MeasureStatistics": sim_stats.MeasureStatistics,
        "SimStatsError": sim_stats.SimStatsError,
    }
    for name, expected in surface.items():
        assert name in pccm_builder.__all__, f"{name} is not exported"
        assert getattr(pccm_builder, name) is expected, name

    assert sorted(pccm_builder.__all__) == sorted(set(pccm_builder.__all__))
    for name in pccm_builder.__all__:
        assert hasattr(pccm_builder, name), name


# ===========================================================================
# the constant-sample invariant
#
# Ordinary O(n * eps) accumulation drift is acceptable on a sample that varies.
# It is NOT acceptable to manufacture dispersion for a sample that does not:
# PCCM must never report stochastic spread for a distribution that has none.
# ===========================================================================
_CONSTANT_CASES = (
    ("ordinary", 0.1),
    ("just above one", 1.1),
    ("large", 1.0e100),
    ("near Double maximum", 1.5e308),
    ("subnormal", 5e-324),
    ("negative", -12345.678),
    ("zero", 0.0),
)


def test_75_a_constant_sample_has_an_exact_mean_and_exactly_zero_deviation() -> None:
    for label, value in _CONSTANT_CASES:
        values = [value] * 1000
        assert len(set(values)) == 1, label

        assert sample_mean(values) == value, (
            f"{label}: mean of 1000 copies of {value!r} came back as "
            f"{sample_mean(values)!r}"
        )
        assert sample_standard_deviation(values) == 0.0, (
            f"{label}: a constant sample acquired dispersion "
            f"{sample_standard_deviation(values)!r}"
        )


def test_76_every_statistic_of_a_constant_sample_is_that_value() -> None:
    ladder = resolve_percentile_ladder(_sim(), _inputs())
    for label, value in _CONSTANT_CASES:
        stats = describe([value] * 1000, ladder.points, label)

        assert stats.mean == value, label
        assert stats.sample_standard_deviation == 0.0, label
        assert stats.minimum == value and stats.maximum == value, label
        assert set(stats.percentiles.values()) == {value}, (
            f"{label}: the ladder is not flat: {sorted(set(stats.percentiles.values()))}"
        )
        for point in ladder.ordered:
            assert stats.percentiles[point] == value, (label, point)


def test_77_the_naive_paths_really_would_have_drifted() -> None:
    """The controls above are not vacuous - the arithmetic they replace drifts."""
    values = [1.5e308] * 1000
    scale = 8.98846567431158e307              # 2**1023
    scaled = [value / scale for value in values]
    naive_total = 0.0
    for value in scaled:
        naive_total += value
    drifted_mean = (naive_total / len(values)) * scale
    assert drifted_mean != 1.5e308, "the accumulation used to be exact after all"
    assert math.isclose(drifted_mean, 1.5e308, rel_tol=1e-12)

    ordinary = [0.1] * 1000
    naive = 0.0
    for value in ordinary:
        naive += value
    assert naive / 1000 != 0.1, "even an ordinary constant sample drifted"
    assert sample_mean(ordinary) == 0.1


def test_78_a_fully_degenerate_simulation_reports_no_dispersion() -> None:
    """End to end: every retained total identical, and every statistic exact.

    The model has no uncertainty anywhere - a degenerate Uniform with an ignored
    Most Likely, a degenerate Triangular, a degenerate Beta-PERT and a certain
    risk with degenerate severity - so its iteration total is one representable
    number repeated N times.
    """
    model = _model(
        costs=[
            _cost("CL-001", "Uniform", 256.0, 999.0, 256.0, quantity=2.0),   # 512
            _cost("CL-002", "Triangular", 0.5, 0.5, 0.5, quantity=8.0),      # 4
            _cost("CL-003", "Beta-PERT", -32.0, -32.0, -32.0, quantity=0.25),  # -8
        ],
        risks=[_risk("R-001", "Uniform", 64.0, None, 64.0, probability=1.0)],
    )
    run, prepared, _ = _run(model, iterations=1500)
    expected = 572.0

    assert set(run.total_nominal) == {expected}
    assert set(run.total_pv) == {expected}
    for measure in ("nominal", "pv"):
        stats = run.summary.measure(measure)
        assert stats.mean == expected, measure
        assert stats.sample_standard_deviation == 0.0, measure
        assert stats.minimum == expected and stats.maximum == expected, measure
        for label in prepared.ladder.ordered:
            assert stats.percentiles[label] == expected, (measure, label)

    # Only the occurrence stream drew anything; no severity or cost uniform was
    # consumed, so the run is degenerate in the sampler sense too.
    assert sorted(record.uniforms_consumed for record in run.diagnostics) == [
        0, 0, 0, 0, 1500
    ]


def test_79_a_near_maximum_constant_simulation_reports_no_dispersion() -> None:
    """The same invariant at the top of the Double range."""
    model = _model(
        costs=[_cost("CL-001", "Uniform", 1.5e308, None, 1.5e308, quantity=1.0)]
    )
    run, prepared, _ = _run(model, iterations=1200)

    assert set(run.total_nominal) == {1.5e308}
    stats = run.summary.nominal
    assert stats.mean == 1.5e308
    assert stats.sample_standard_deviation == 0.0
    assert stats.minimum == 1.5e308 and stats.maximum == 1.5e308
    assert set(stats.percentiles.values()) == {1.5e308}
    assert len(stats.percentiles) == len(prepared.ladder.ordered) == 11


def test_80_a_sample_that_varies_by_one_ulp_is_not_flattened() -> None:
    """The invariant tests the DOUBLE, not an approximate closeness."""
    base = 1.0
    nudged = math.nextafter(base, 2.0)
    values = [base] * 999 + [nudged]

    assert len(set(values)) == 2
    assert sample_standard_deviation(values) > 0.0, (
        "a real one-ulp dispersion was flattened to zero"
    )
    assert percentile_type7(values, 1.0) == nudged
    assert percentile_type7(values, 0.0) == base


# ===========================================================================
# describe() sorts once
# ===========================================================================
def test_81_describe_sorts_exactly_once_per_measure() -> None:
    """Eleven percentiles used to mean twelve sorts: one in `describe` and one
    inside `percentile_type7` for every label."""
    import pccm_builder.sim_stats as stats_module

    ladder = resolve_percentile_ladder(_sim(), _inputs())
    assert len(ladder.points) == 11

    values = [float((index * 7919) % 1000) for index in range(1000)]
    calls = []
    real_sorted = sorted

    def counting_sorted(*args, **kwargs):
        calls.append(1)
        return real_sorted(*args, **kwargs)

    stats_module.sorted = counting_sorted          # shadows the builtin lookup
    try:
        described = stats_module.describe(values, ladder.points, "sort count")
    finally:
        del stats_module.sorted

    assert len(calls) == 1, (
        f"describe sorted {len(calls)} times for an 11-point ladder; the accepted "
        "operation model is exactly one sort per measure"
    )

    # And the numbers are unchanged by the split.
    for label, p in ladder.points:
        assert described.percentiles[label] == percentile_type7(values, p)


def test_82_the_public_percentile_helper_still_sorts_its_own_copy() -> None:
    import pccm_builder.sim_stats as stats_module

    values = [9.0, 1.0, 5.0, 3.0]
    original = list(values)
    calls = []
    real_sorted = sorted

    def counting_sorted(*args, **kwargs):
        calls.append(1)
        return real_sorted(*args, **kwargs)

    stats_module.sorted = counting_sorted
    try:
        assert stats_module.percentile_type7(values, 0.5) == 4.0
    finally:
        del stats_module.sorted

    assert len(calls) == 1
    assert values == original, "the caller's sequence was reordered"


def test_83_the_sorted_helper_refuses_nothing_the_public_one_accepts() -> None:
    """Every hand vector of `test_36` is unchanged by the split."""
    test_36_type_7_percentiles_match_hand_derived_vectors()
    assert percentile_type7([1.0, 2.0, 3.0, 4.0], 0.9) == 3.7
    assert percentile_type7([10.0, 20.0, 60.0], 0.5) == 20.0
    assert percentile_type7([float(v) for v in range(1, 11)], 0.5) == 5.5


# ===========================================================================
# the RNG reference binding
# ===========================================================================
def test_84_the_accepted_reference_reproduces_every_pinned_number() -> None:
    """The binding is coherence enforcement only: nothing moved."""
    run, prepared, result = _run(_mixed_model(), seed=12345, iterations=2000)

    assert run.result_digest == "7F58EA884DAA8D65"
    assert run.total_nominal[0] == 641.731721357026
    assert run.total_nominal[1] == 304.5370973159493
    assert run.summary.nominal.mean == 376.3270995496381
    assert run.summary.nominal.sample_standard_deviation == 119.53338321454508
    assert run.summary.nominal.percentiles["P10"] == 253.0081369880356
    assert run.summary.nominal.percentiles["P50"] == 335.15094832897046
    assert run.summary.nominal.percentiles["P90"] == 574.8427214665053
    assert run.summary.nominal.percentiles["P95"] == 612.8092541358852
    assert contingency_at(
        run.summary, "P80", deterministic_base_of(result)
    ).nominal == 261.9975727286816
    assert prepared.rng_version == 1 and run.rng_version == 1


def test_85_the_binding_is_a_value_not_a_pointer() -> None:
    """`role_order` inside a frozen `RngReference` is a live dict; the signature
    must not be able to change underneath a prepared model."""
    reference = _ref()
    signature = rng_reference_signature(reference)

    assert isinstance(signature, tuple)
    assert signature == rng_reference_signature(reference), "the snapshot is unstable"
    assert signature == rng_reference_signature(
        RngReference.from_contracts(_sim(), _inputs())
    ), "two references derived from the same contracts disagree"

    def immutable(value) -> bool:
        if isinstance(value, tuple):
            return all(immutable(item) for item in value)
        return isinstance(value, (int, float, str, bool)) or value is None

    assert immutable(signature), "the signature holds something mutable"


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
