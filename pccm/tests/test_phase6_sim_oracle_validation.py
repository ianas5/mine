#!/usr/bin/env python3
"""PCCM Phase 6 Step-4 mutation controls and scope guards for the simulation oracle.

Every control here implements a LOCAL wrong variant and shows the accepted
engine disagrees with it. Nothing in `spec/` is touched: a control that had to
edit a contract to fail would be proving something about the contract instead of
about the implementation.

THE LOCAL ENGINE IS THE INSTRUMENT. `_local_run` is an independent
re-implementation of one iteration, parameterised by exactly one mutation. Run
with no mutation it must reproduce the accepted engine BIT FOR BIT - that
equality is asserted first, and it is what makes every other control
non-vacuous. If the unmutated local engine drifted, a mutation could "differ"
for the wrong reason.

WHERE A MUTATION IS NOT DETECTABLE BY LUCK, IT IS DETECTED BY CONSTRUCTION.
Substituting naive left-to-right addition for the accepted accumulation changes
nothing on an ordinary fixture - tier 1 of `safe_signed_sum` IS left-to-right
addition - and both halves of that fact are recorded rather than one.

THE INSTRUMENTS ARE VERSION-INDEPENDENT. That accumulation mutation is written
out as an explicit loop, never as Python's built-in `sum`, which applies
Neumaier compensation to float sequences from CPython 3.12 and is therefore no
longer a definition of naive addition. A control built on it would be measuring
the interpreter.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import dataclasses
import math
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    Draw,
    RngReference,
    bernoulli_occurs,
    load_calc_contract,
    load_contract,
    load_sim_contract,
    sample_prepared_beta,
    sample_triangular,
    sample_uniform,
)
from pccm_builder.calc_cases import tolerances_from  # noqa: E402
from pccm_builder.calc_fingerprint import (  # noqa: E402
    encode_section,
    fingerprint,
    integer_field,
    number_field,
    text_field,
)
from pccm_builder.calc_numeric import (  # noqa: E402
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
    DeterministicBase,
    SimOracleError,
    contingency_at,
    deterministic_base_of,
    prepare_simulation,
    result_digest,
    result_digest_stream,
    rng_reference_signature,
    run_simulation,
)
from pccm_builder.sim_sample import (  # noqa: E402
    FAMILY_TRIANGULAR,
    FAMILY_UNIFORM,
)
from pccm_builder.sim_stats import (  # noqa: E402
    MeasureStatistics,
    describe,
    percentile_type7,
    sample_mean,
    sample_standard_deviation,
)

SPEC = PCCM_ROOT / "spec"
BUILDER = PCCM_ROOT / "builder" / "pccm_builder"
SIM_ORACLE = BUILDER / "sim_oracle.py"
SIM_STATS = BUILDER / "sim_stats.py"

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
# ---------------------------------------------------------------------------
def _cost(permanent_id, distribution="Triangular", minimum=80.0, most_likely=100.0,
          maximum=150.0, quantity=1.0, weights=(1.0,)):
    return CostDriver(
        permanent_id, distribution, "SAR", "Standard",
        minimum, most_likely, maximum, weights, quantity=quantity,
    )


def _risk(permanent_id, distribution="Triangular", minimum=100.0, most_likely=200.0,
          maximum=400.0, probability=0.5, weights=(1.0,)):
    return RiskDriver(
        permanent_id, distribution, "SAR", "Standard",
        minimum, most_likely, maximum, weights, probability=probability,
    )


def _model(costs=(), risks=(), duration=1, inflation=None, start_year=2026):
    return CalculationModel(
        timeline=AppliedTimeline(2026, start_year, duration),
        discount_rate=0.10,
        fx_rows=(FxRow("SAR", 1),),
        inflation_rates={"Standard": inflation or {}},
        cost_drivers=tuple(costs),
        risk_drivers=tuple(risks),
    )


def _prepare(model, seed=12345, iterations=1000):
    return prepare_simulation(
        _ref(), _sim(), _inputs(), model, _tolerances(),
        effective_seed=seed, iterations=iterations,
    )


def _mixed():
    return _model(
        costs=[
            _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=2.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, quantity=1.0),
            _cost("CL-003", "Beta-PERT", 5.0, 7.0, 20.0, quantity=3.0),
        ],
        risks=[
            _risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.3),
            _risk("R-002", "Uniform", 50.0, None, 90.0, probability=0.7),
        ],
    )


# ---------------------------------------------------------------------------
# the local engine
# ---------------------------------------------------------------------------
def naive_left_to_right(values) -> float:
    """Ordinary unchecked left-to-right binary64 addition. Nothing else.

    WHY THIS IS SPELLED OUT RATHER THAN CALLING `sum`. Python's built-in `sum`
    is not a stable definition of naive accumulation: from CPython 3.12 it
    applies Neumaier compensation to float sequences, so `sum([0.1] * 20)` is
    `2.0` on 3.13 and `2.0000000000000004` on 3.11 - and the second is what a
    plain accumulation loop produces on both. A mutation control built on `sum`
    is therefore testing the host runtime rather than the mutation.

    The mutation being modelled is exactly "substitute `safe_signed_sum` with
    ordinary unchecked addition", so ordinary unchecked addition is written out.
    No `sum`, no `math.fsum`, no compensated or pairwise algorithm.
    """
    total = 0.0
    for value in values:
        total = total + value
    return total


def _sample(state, driver):
    if driver.distribution == FAMILY_UNIFORM:
        return sample_uniform(_ref(), state, driver.minimum, driver.maximum, driver.most_likely)
    if driver.distribution == FAMILY_TRIANGULAR:
        return sample_triangular(_ref(), state, driver.minimum, driver.most_likely, driver.maximum)
    return sample_prepared_beta(_ref(), state, driver.beta_shape)


def _local_run(prepared, *, mutation: str = "", row_order: tuple = ()):
    """One iteration of the engine, re-derived, with at most one thing wrong."""
    costs = list(prepared.cost_drivers)
    risks = list(prepared.risk_drivers)
    cost_state = {driver.permanent_id: driver.value_initial_state for driver in costs}
    severity_state = {driver.permanent_id: driver.value_initial_state for driver in risks}
    occurrence_state = {
        driver.permanent_id: driver.occurrence_initial_state for driver in risks
    }
    nominal: list[float] = []
    pv: list[float] = []

    for _ in range(prepared.iterations):
        entries: list[tuple] = []

        for driver in costs:
            drawn = _sample(cost_state[driver.permanent_id], driver)
            cost_state[driver.permanent_id] = drawn.state
            quantity = driver.quantity
            if mutation == "quantity_omitted":
                quantity = 1.0
            elif mutation == "quantity_twice":
                quantity = quantity * quantity
            entries.append(
                (driver.permanent_id, "cost",
                 drawn.value * quantity * driver.knom,
                 drawn.value * quantity * driver.kpv)
            )

        for driver in risks:
            if mutation == "merged_streams":
                occurrence = bernoulli_occurs(
                    _ref(), severity_state[driver.permanent_id], driver.probability
                )
                severity_state[driver.permanent_id] = occurrence.state
            else:
                occurrence = bernoulli_occurs(
                    _ref(), occurrence_state[driver.permanent_id], driver.probability
                )
                occurrence_state[driver.permanent_id] = occurrence.state

            severity = None
            if not (mutation == "severity_only_on_occurrence" and not occurrence.occurred):
                severity = _sample(severity_state[driver.permanent_id], driver)
                severity_state[driver.permanent_id] = severity.state

            if mutation == "probability_in_k":
                value = severity.value if severity is not None else 0.0
                contribution = (
                    value * driver.probability * driver.knom,
                    value * driver.probability * driver.kpv,
                )
            elif occurrence.occurred and severity is not None:
                quantity = 3.0 if mutation == "risk_quantity" else 1.0
                contribution = (
                    severity.value * quantity * driver.knom,
                    severity.value * quantity * driver.kpv,
                )
            else:
                contribution = (0.0, 0.0)
            entries.append((driver.permanent_id, "risk", contribution[0], contribution[1]))

        if mutation == "risks_first":
            entries = (
                [entry for entry in entries if entry[1] == "risk"]
                + [entry for entry in entries if entry[1] == "cost"]
            )
        elif mutation == "reversed_accumulation":
            entries = list(reversed(entries))
        elif mutation == "physical_row_order":
            entries.sort(key=lambda entry: row_order.index(entry[0]))

        nominal_terms = [entry[2] for entry in entries]
        pv_terms = [entry[3] for entry in entries]

        if mutation == "naive_accumulation":
            nominal.append(naive_left_to_right(nominal_terms))
            pv.append(naive_left_to_right(pv_terms))
        elif mutation == "pv_from_nominal":
            total = safe_signed_sum(nominal_terms, "nominal")
            first = (costs + risks)[0]
            nominal.append(total)
            pv.append(total * (first.kpv / first.knom))
        else:
            nominal.append(safe_signed_sum(nominal_terms, "nominal"))
            pv.append(safe_signed_sum(pv_terms, "pv"))

    return tuple(nominal), tuple(pv)


# ===========================================================================
# the instrument itself
# ===========================================================================
def test_01_the_unmutated_local_engine_reproduces_the_accepted_engine_exactly() -> None:
    """Every control below rests on this. No tolerance."""
    prepared, _ = _prepare(_mixed(), iterations=1200)
    run = run_simulation(_ref(), prepared)
    nominal, pv = _local_run(prepared)

    assert nominal == run.total_nominal
    assert pv == run.total_pv
    assert result_digest(prepared.sim_method_version, nominal, pv) == run.result_digest


# ===========================================================================
# contribution mutations
# ===========================================================================
def test_02_a_cost_sample_read_as_a_total_instead_of_a_unit_is_caught() -> None:
    """The engine's sample lies in the UNIT support, not the total support.

    Recovering the sample from the contribution shows which scale it was drawn
    on: with `Quantity = 4` a unit sample lives in `[80, 150]` and a total sample
    would live in `[320, 600]`. The two supports do not overlap, so no run can
    be ambiguous about it.
    """
    prepared, _ = _prepare(
        _model(costs=[_cost("CL-001", "Triangular", 80.0, 100.0, 150.0, quantity=4.0)]),
        iterations=1000,
    )
    run = run_simulation(_ref(), prepared)
    driver = prepared.cost_drivers[0]
    assert driver.knom == 1.0

    recovered = [total / driver.quantity for total in run.total_nominal]
    assert all(80.0 <= value <= 150.0 for value in recovered), "the sample is not a unit cost"
    assert not any(320.0 <= value <= 600.0 for value in recovered), (
        "the recovered values sit in the TOTAL support"
    )
    # Read as a total, the retained totals would be the samples themselves.
    assert all(total > 150.0 for total in run.total_nominal)


def test_03_an_omitted_quantity_is_caught() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    accepted = run_simulation(_ref(), prepared).total_nominal
    mutated, _ = _local_run(prepared, mutation="quantity_omitted")
    assert mutated != accepted
    assert sum(1 for a, b in zip(accepted, mutated) if a != b) == len(accepted)


def test_04_a_quantity_applied_twice_is_caught() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    accepted = run_simulation(_ref(), prepared).total_nominal
    mutated, _ = _local_run(prepared, mutation="quantity_twice")
    assert mutated != accepted
    assert sum(1 for a, b in zip(accepted, mutated) if a != b) == len(accepted)


def test_05_a_quantity_introduced_on_a_risk_is_caught() -> None:
    """Risks have no Quantity. Any factor other than 1 changes every occurrence."""
    prepared, _ = _prepare(
        _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.6)]),
        iterations=1000,
    )
    accepted = run_simulation(_ref(), prepared).total_nominal
    mutated, _ = _local_run(prepared, mutation="risk_quantity")

    assert mutated != accepted
    occurred = [index for index, total in enumerate(accepted) if total != 0.0]
    assert occurred, "the fixture never occurs, so the control is vacuous"
    assert all(mutated[index] == accepted[index] * 3.0 for index in occurred)
    for driver in prepared.risk_drivers:
        assert driver.quantity is None, "a Risk carries a Quantity into the hot loop"


def test_06_probability_folded_into_the_k_factors_is_caught() -> None:
    """Folding Probability into K removes the Bernoulli entirely, so the run stops
    being stochastic in occurrence and every iteration contributes."""
    prepared, _ = _prepare(
        _model(risks=[_risk("R-001", "Uniform", 100.0, None, 400.0, probability=0.4)]),
        iterations=1000,
    )
    accepted = run_simulation(_ref(), prepared).total_nominal
    mutated, _ = _local_run(prepared, mutation="probability_in_k")

    assert mutated != accepted
    assert 0.0 in set(accepted), "the accepted run must contain non-occurrences"
    assert 0.0 not in set(mutated), "the folded variant should never be zero"

    for driver in prepared.risk_drivers:
        assert driver.knom == 1.0 and driver.kpv == 1.0, (
            "Probability reached Knom or Kpv during preparation"
        )
        assert driver.probability == 0.4


def test_07_pv_derived_from_the_nominal_total_is_caught() -> None:
    """Two drivers whose Kpv/Knom ratios differ; one ratio cannot serve both."""
    model = _model(
        costs=[
            _cost("CL-001", "Uniform", 100.0, None, 100.0, quantity=1.0, weights=(1.0, 0.0)),
            _cost("CL-002", "Triangular", 80.0, 100.0, 150.0, quantity=1.0,
                  weights=(0.0, 1.0)),
        ],
        duration=2, start_year=2027, inflation={2027: 0.0, 2028: 0.0},
    )
    prepared, _ = _prepare(model, iterations=1000)
    first, second = prepared.cost_drivers
    assert first.kpv / first.knom != second.kpv / second.knom, "the fixture is degenerate"

    run = run_simulation(_ref(), prepared)
    _, mutated_pv = _local_run(prepared, mutation="pv_from_nominal")
    assert mutated_pv != run.total_pv
    assert run.total_pv != run.total_nominal


# ===========================================================================
# D6-18b and stream mutations
# ===========================================================================
def test_08_sampling_severity_only_on_occurrence_is_caught() -> None:
    prepared, _ = _prepare(
        _model(risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.3)]),
        iterations=1500,
    )
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="severity_only_on_occurrence")

    assert mutated != run.total_nominal
    occurred = sum(1 for total in run.total_nominal if total != 0.0)
    assert 0 < occurred < 1500
    severity = next(record for record in run.diagnostics if record.role == "severity")
    assert severity.uniforms_consumed == 1500, (
        "the accepted engine did not sample severity every iteration"
    )


def test_09_merging_the_occurrence_and_severity_streams_is_caught() -> None:
    prepared, _ = _prepare(
        _model(risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability=0.5)]),
        iterations=1200,
    )
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="merged_streams")

    assert mutated != run.total_nominal
    driver = prepared.risk_drivers[0]
    assert driver.occurrence_initial_state != driver.value_initial_state
    assert driver.occurrence_stream_index != driver.value_stream_index


def test_10_probability_only_comparability_fails_under_the_severity_mutation() -> None:
    """The invariance D6-18b buys is exactly what the mutation destroys."""
    accepted_states = {}
    mutated_states = {}
    for probability in (0.2, 0.8):
        prepared, _ = _prepare(
            _model(
                risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0,
                             probability=probability)]
            ),
            iterations=1000,
        )
        run = run_simulation(_ref(), prepared)
        severity = next(record for record in run.diagnostics if record.role == "severity")
        accepted_states[probability] = severity.final_state

        # The mutated engine's severity stream depends on the occurrences.
        driver = prepared.risk_drivers[0]
        state = driver.value_initial_state
        occurrence_state = driver.occurrence_initial_state
        for _ in range(1000):
            occurrence = bernoulli_occurs(_ref(), occurrence_state, driver.probability)
            occurrence_state = occurrence.state
            if occurrence.occurred:
                state = _sample(state, driver).state
        mutated_states[probability] = state

    assert accepted_states[0.2] == accepted_states[0.8], (
        "the accepted engine lost comparability"
    )
    assert mutated_states[0.2] != mutated_states[0.8], (
        "the control is vacuous: the mutated engine stayed comparable"
    )


# ===========================================================================
# accumulation-order mutations
#
# Binary64 addition is not associative, but it is not UNIVERSALLY sensitive
# either: `1e16 + 1` ties straight back to `1e16`, so a small term is absorbed
# or survives depending on WHEN the two large terms cancel. Each order mutation
# therefore gets a construction built for it, and each one states both sums.
#
#   THREE-TERM  canonical [1e16, 1, -1e16] = 0.0
#               risks first [-1e16, 1e16, 1] = 1.0
#
#   FOUR-TERM   canonical [1e16, 1, -1e16, 1] = 1.0
#               reversed  [1, -1e16, 1, 1e16] = 0.0
#               register  [1, 1, 1e16, -1e16] = 2.0
# ===========================================================================
_BIG = 1.0e16


def _three_term_fixture(cost_order=None):
    """Canonical `[1e16, 1, -1e16]`, with the risk carrying the negative term."""
    costs = [
        _cost("CL-001", "Uniform", _BIG, None, _BIG, quantity=1.0),
        _cost("CL-002", "Uniform", 1.0, None, 1.0, quantity=1.0),
    ]
    return _model(
        costs=cost_order if cost_order is not None else costs,
        risks=[_risk("R-001", "Uniform", -_BIG, None, -_BIG, probability=1.0)],
    )


def _four_term_fixture():
    """Canonical `[1e16, 1, -1e16, 1]`, two costs and two certain risks."""
    return _model(
        costs=[
            _cost("CL-001", "Uniform", _BIG, None, _BIG, quantity=1.0),
            _cost("CL-002", "Uniform", 1.0, None, 1.0, quantity=1.0),
        ],
        risks=[
            _risk("R-001", "Uniform", -_BIG, None, -_BIG, probability=1.0),
            _risk("R-002", "Uniform", 1.0, None, 1.0, probability=1.0),
        ],
    )


def test_11_using_the_physical_row_order_is_caught() -> None:
    """The register order `[CL-002, R-002, CL-001, R-001]` sums to `2.0`.

    Canonical order sums to `1.0`, so a kernel that walked the register instead
    of sorting would report a different total for the same model.
    """
    register = ("CL-002", "R-002", "CL-001", "R-001")
    prepared, _ = _prepare(_four_term_fixture(), iterations=1000)
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="physical_row_order", row_order=register)

    assert [driver.permanent_id for driver in prepared.drivers] == [
        "CL-001", "CL-002", "R-001", "R-002"
    ]
    assert set(run.total_nominal) == {1.0}
    assert set(mutated) == {2.0}, sorted(set(mutated))
    assert mutated != run.total_nominal


def test_12_accumulating_risks_before_costs_is_caught() -> None:
    """Three terms: canonical `0.0`, risks first `1.0`."""
    prepared, _ = _prepare(_three_term_fixture(), iterations=1000)
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="risks_first")

    assert safe_signed_sum([_BIG, 1.0, -_BIG], "canonical") == 0.0
    assert safe_signed_sum([-_BIG, _BIG, 1.0], "risks first") == 1.0
    assert set(run.total_nominal) == {0.0}
    assert set(mutated) == {1.0}
    assert mutated != run.total_nominal


def test_13_reversing_the_accumulation_is_caught() -> None:
    """Four terms: canonical `1.0`, fully reversed `0.0`."""
    prepared, _ = _prepare(_four_term_fixture(), iterations=1000)
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="reversed_accumulation")

    assert safe_signed_sum([_BIG, 1.0, -_BIG, 1.0], "canonical") == 1.0
    assert safe_signed_sum([1.0, -_BIG, 1.0, _BIG], "reversed") == 0.0
    assert set(run.total_nominal) == {1.0}
    assert set(mutated) == {0.0}
    assert mutated != run.total_nominal


def test_14_the_order_controls_do_not_claim_reversal_always_matters() -> None:
    """The difference is REQUIRED ONLY ON THE CONSTRUCTED FIXTURE.

    On an ordinary model reversing the contributions changes nothing on most
    iterations, and saying so is part of the evidence: row-order invariance is
    universal precisely because canonical accumulation order does not depend on
    the register.
    """
    prepared, _ = _prepare(_mixed(), iterations=1000)
    run = run_simulation(_ref(), prepared)
    reversed_totals, _ = _local_run(prepared, mutation="reversed_accumulation")

    identical = sum(1 for a, b in zip(run.total_nominal, reversed_totals) if a == b)
    assert identical > 0, "the ordinary fixture is unexpectedly order-sensitive everywhere"


def test_15_naive_left_to_right_accumulation_is_caught_on_the_constructed_fixture() -> None:
    """BOTH HALVES, recorded honestly.

    On an ordinary model naive left-to-right addition and the accepted
    accumulation agree on every iteration, bit for bit, because tier 1 of
    `safe_signed_sum` IS left-to-right addition with a range check. The
    constructed fixture is the one where they part: a partial sum leaves Double
    range, the naive path reports infinity, and the accepted primitive returns
    the representable total the model actually has.
    """
    ordinary, _ = _prepare(_mixed(), iterations=1000)
    ordinary_run = run_simulation(_ref(), ordinary)
    ordinary_mutated, _ = _local_run(ordinary, mutation="naive_accumulation")
    assert ordinary_mutated == ordinary_run.total_nominal, (
        "the two accumulations already differ on an ordinary model"
    )

    constructed = _model(
        costs=[
            _cost("CL-001", "Uniform", 1.5e308, None, 1.5e308, quantity=1.0),
            _cost("CL-002", "Uniform", 1.5e308, None, 1.5e308, quantity=1.0),
            _cost("CL-003", "Uniform", -1.5e308, None, -1.5e308, quantity=1.0),
        ]
    )
    prepared, _ = _prepare(constructed, iterations=1000)
    run = run_simulation(_ref(), prepared)
    mutated, _ = _local_run(prepared, mutation="naive_accumulation")

    assert set(run.total_nominal) == {1.5e308}
    assert all(math.isinf(total) for total in mutated), sorted(set(mutated))[:3]
    assert mutated != run.total_nominal


def test_15a_the_mutation_does_not_depend_on_the_built_in_sum() -> None:
    """The instrument is version-independent, and the reason is recorded.

    `sum` over floats is compensated from CPython 3.12, so it is no longer a
    definition of naive accumulation. This test does not require the two to
    agree or to disagree on the host runtime - it requires that the MUTATION is
    the explicit accumulator either way, so the control behaves identically on
    3.11 and on 3.13.
    """
    terms = [0.1] * 20
    explicit = naive_left_to_right(terms)

    # The explicit accumulator is pinned to a literal, so it cannot drift with
    # the runtime, and it agrees with the accepted tier-1 accumulation exactly.
    assert explicit == 2.0000000000000004, explicit
    assert explicit == safe_signed_sum(terms, "tier 1"), (
        "the accepted tier-1 accumulation is not plain left-to-right addition"
    )

    # Whether the HOST's `sum` agrees is a property of the interpreter, not of
    # PCCM. It is neither asserted nor relied on - only that when it disagrees,
    # it disagrees by being compensated, which is precisely why it cannot serve
    # as the mutation.
    if sum(terms) != explicit:
        assert sum(terms) == 2.0, (
            f"an unexpected built-in sum result {sum(terms)!r} on "
            f"Python {sys.version_info.major}.{sys.version_info.minor}"
        )

    # THE STRUCTURAL GUARANTEE: the mutation instrument never calls `sum` or
    # `math.fsum` at all, so no interpreter change can move it.
    module = ast.parse(Path(__file__).read_text(encoding="utf-8"))
    guarded = {"naive_left_to_right", "_local_run"}
    for node in ast.walk(module):
        if not isinstance(node, ast.FunctionDef) or node.name not in guarded:
            continue
        for inner in ast.walk(node):
            if isinstance(inner, ast.Call):
                name = getattr(inner.func, "id", None) or getattr(inner.func, "attr", None)
                assert name not in {"sum", "fsum"}, (
                    f"{node.name} calls {name!r}; the mutation must be explicit addition"
                )
        guarded.discard(node.name)
    assert not guarded, f"the guarded functions were not found: {sorted(guarded)}"


# ===========================================================================
# result-digest mutations
# ===========================================================================
_DIGEST_NOMINAL = (1234567.8901234567, 0.0, -0.5, 1e300, 5e-324)
_DIGEST_PV = (1000000.0, 0.0, -0.25, 1e299, 1e-320)
_ACCEPTED_DIGEST = "3181AF89642DE500"


def test_16_the_accepted_digest_is_the_retained_step0_vector() -> None:
    assert result_digest(1, _DIGEST_NOMINAL, _DIGEST_PV) == _ACCEPTED_DIGEST


def test_17_dropping_the_iteration_index_from_a_record_is_caught() -> None:
    records = [
        (number_field(nominal), number_field(pv))
        for nominal, pv in zip(_DIGEST_NOMINAL, _DIGEST_PV)
    ]
    stream = (
        text_field("PCCM-RD").encode()
        + integer_field(1).encode()
        + encode_section("RESULT", records)
    )
    assert fingerprint(stream) != _ACCEPTED_DIGEST
    assert stream != result_digest_stream(1, _DIGEST_NOMINAL, _DIGEST_PV)


def test_18_swapping_nominal_and_pv_is_caught() -> None:
    swapped = result_digest(1, _DIGEST_PV, _DIGEST_NOMINAL)
    assert swapped != _ACCEPTED_DIGEST
    assert swapped == "63A0E93074F0C2EA", "this is the retained swapped vector"


def test_19_omitting_the_digest_version_is_caught() -> None:
    records = [
        (integer_field(index + 1), number_field(nominal), number_field(pv))
        for index, (nominal, pv) in enumerate(zip(_DIGEST_NOMINAL, _DIGEST_PV))
    ]
    stream = text_field("PCCM-RD").encode() + encode_section("RESULT", records)
    assert fingerprint(stream) != _ACCEPTED_DIGEST
    # And a DIFFERENT version is a different digest, so the field is load-bearing.
    assert result_digest(2, _DIGEST_NOMINAL, _DIGEST_PV) == "7E8D58C46CCDD798"
    assert result_digest(2, _DIGEST_NOMINAL, _DIGEST_PV) != _ACCEPTED_DIGEST


def test_20_sorting_the_samples_before_the_digest_is_caught() -> None:
    """A sorted digest cannot tell two runs apart that produced the same multiset
    of totals in different orders - which is most of what the digest is for."""
    order = sorted(range(len(_DIGEST_NOMINAL)), key=lambda index: _DIGEST_NOMINAL[index])
    sorted_nominal = tuple(_DIGEST_NOMINAL[index] for index in order)
    sorted_pv = tuple(_DIGEST_PV[index] for index in order)
    assert sorted_nominal != _DIGEST_NOMINAL, "the fixture is already sorted"
    assert result_digest(1, sorted_nominal, sorted_pv) != _ACCEPTED_DIGEST

    # The retained reversal vector makes the same point from the other side.
    assert result_digest(1, tuple(reversed(_DIGEST_NOMINAL)), tuple(reversed(_DIGEST_PV))) == (
        "4E0FEE211853E8F6"
    )


def test_21_a_real_run_digest_changes_when_its_arrays_are_reordered() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1100)
    run = run_simulation(_ref(), prepared)
    sorted_digest = result_digest(
        prepared.sim_method_version, sorted(run.total_nominal), sorted(run.total_pv)
    )
    assert sorted_digest != run.result_digest
    assert list(run.total_nominal) != sorted(run.total_nominal)


# ===========================================================================
# statistics mutations
# ===========================================================================
_STATS_SAMPLE = [2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0]
_EXTREME_SAMPLE = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]


def test_22_a_nearest_rank_percentile_is_caught() -> None:
    """Nearest rank picks an observation; type 7 interpolates between two."""
    values = [1.0, 2.0, 3.0, 4.0]
    for p in (0.25, 0.4, 0.9):
        rank = min(len(values) - 1, max(0, math.ceil(p * len(values)) - 1))
        nearest = sorted(values)[rank]
        accepted = percentile_type7(values, p)
        assert accepted != nearest, (p, accepted, nearest)
    assert percentile_type7(values, 0.9) == 3.7
    assert 3.7 not in values, "type 7 produced an observation, not an interpolation"


def test_23_the_unsafe_difference_interpolation_is_caught() -> None:
    low, high = -1.7e308, 1.7e308
    for f in (0.25, 0.5, 0.75):
        convex = (1.0 - f) * low + f * high
        unsafe = low + f * (high - low)
        assert math.isfinite(convex)
        assert not math.isfinite(unsafe)
        assert percentile_type7([low, high], f) == convex


def test_24_a_population_standard_deviation_is_caught() -> None:
    count = len(_STATS_SAMPLE)
    mean = sum(_STATS_SAMPLE) / count
    population = math.sqrt(sum((value - mean) ** 2 for value in _STATS_SAMPLE) / count)
    accepted = sample_standard_deviation(_STATS_SAMPLE)

    assert math.isclose(accepted, math.sqrt(32.0 / 7.0), rel_tol=1e-15)
    assert math.isclose(population, 2.0, rel_tol=1e-15)
    assert not math.isclose(accepted, population, rel_tol=1e-9)
    assert accepted > population, "the n - 1 divisor must give the larger value"


def test_25_a_naive_sum_of_squares_deviation_is_caught() -> None:
    """The textbook one-pass form reports infinity where the answer exists."""
    assert math.isinf(sum(value * value for value in _EXTREME_SAMPLE))
    accepted = sample_standard_deviation(_EXTREME_SAMPLE)
    assert math.isfinite(accepted) and accepted > 1.6e308

    # On an ordinary sample the naive form agrees, which is why the extreme
    # fixture is the one that carries the proof.
    count = len(_STATS_SAMPLE)
    mean = sum(_STATS_SAMPLE) / count
    squares = sum(value * value for value in _STATS_SAMPLE)
    naive = math.sqrt((squares - count * mean * mean) / (count - 1))
    assert math.isclose(naive, sample_standard_deviation(_STATS_SAMPLE), rel_tol=1e-12)


def test_26_a_naive_mean_accumulation_is_caught() -> None:
    naive = 0.0
    for value in _EXTREME_SAMPLE:
        naive += value
    assert math.isinf(naive)
    assert math.isinf(naive / len(_EXTREME_SAMPLE))
    assert sample_mean(_EXTREME_SAMPLE) == 8.5e307

    assert naive_left_to_right(_STATS_SAMPLE) / len(_STATS_SAMPLE) == sample_mean(
        _STATS_SAMPLE
    ), "the two agree wherever the naive sum exists, which is the whole point"


def test_27_an_unguarded_welford_deviation_is_caught() -> None:
    """`delta = x - mean` overflows on the same sample the accepted path handles."""
    mean = sample_mean(_EXTREME_SAMPLE)
    assert math.isinf(_EXTREME_SAMPLE[0] - mean)
    assert math.isfinite(sample_standard_deviation(_EXTREME_SAMPLE))


# ===========================================================================
# reporting mutations
# ===========================================================================
def test_28_a_contingency_baseline_of_the_mean_is_caught() -> None:
    prepared, result = _prepare(_mixed(), iterations=1500)
    run = run_simulation(_ref(), prepared)
    base = deterministic_base_of(result)
    accepted = contingency_at(run.summary, "P80", base)

    mean_based = contingency_at(
        run.summary, "P80", DeterministicBase(run.summary.nominal.mean, run.summary.pv.mean)
    )
    expected_based = contingency_at(
        run.summary, "P80", DeterministicBase(result.totals.e_nom, result.totals.e_pv)
    )
    a_plus_emv = contingency_at(
        run.summary,
        "P80",
        DeterministicBase(
            result.totals.a_nom + result.totals.d_nom,
            result.totals.a_pv + result.totals.d_pv,
        ),
    )

    assert accepted.nominal != mean_based.nominal
    assert accepted.nominal != expected_based.nominal
    assert accepted.nominal != a_plus_emv.nominal
    assert accepted.base_nominal == result.totals.a_nom


def test_29_a_clamped_contingency_is_caught() -> None:
    prepared, _ = _prepare(
        _model(costs=[_cost("CL-001", "Triangular", 10.0, 20.0, 30.0, quantity=1.0)]),
        iterations=1500,
    )
    run = run_simulation(_ref(), prepared)
    contingency = contingency_at(run.summary, "P50", DeterministicBase(1.0e6, 1.0e6))

    assert contingency.nominal < 0.0
    assert max(contingency.nominal, 0.0) == 0.0, "the control is vacuous"
    assert contingency.nominal != 0.0


def test_30_the_selected_confidence_level_entering_execution_is_caught() -> None:
    """A mutated engine in which the selection reaches the seed produces a
    different digest for every level; the accepted engine produces one."""
    prepared, result = _prepare(_mixed(), seed=5000, iterations=1000)
    run = run_simulation(_ref(), prepared)
    base = deterministic_base_of(result)

    accepted_digests = set()
    for level in prepared.ladder.selectable:
        contingency_at(run.summary, level, base)
        accepted_digests.add(run.result_digest)
    assert len(accepted_digests) == 1, "the Selected CL reached the retained samples"

    mutated_digests = set()
    for offset, _level in enumerate(prepared.ladder.selectable):
        leaked, _ = _prepare(_mixed(), seed=5000 + offset, iterations=1000)
        mutated_digests.add(run_simulation(_ref(), leaked).result_digest)
    assert len(mutated_digests) == len(prepared.ladder.selectable), (
        "the control is vacuous: the leak produced no divergence"
    )


# ===========================================================================
# scope and static guards - section 34
# ===========================================================================
def _semantic_identifiers(path: Path) -> set[str]:
    """Every NAME the module actually uses - not the prose it is written in.

    Docstrings and message text are excluded on purpose. A module that documents
    what it does NOT do has to be allowed to say the words; what matters is
    whether it CALLS or TOUCHES anything so named. Imported modules, attribute
    accesses, bound names, arguments and definitions are all included.

    SCOPE OF THE PROOF. This is a static check on identifiers, so it cannot see
    a name assembled at run time from strings. The complementary runtime guard is
    in the conformance suite, where a whole run executes with `open` replaced and
    opens nothing at all.
    """
    tree = ast.parse(path.read_text(encoding="utf-8"))
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            names.add(node.id)
        elif isinstance(node, ast.Attribute):
            names.add(node.attr)
        elif isinstance(node, ast.arg):
            names.add(node.arg)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, ast.keyword) and node.arg:
            names.add(node.arg)
        elif isinstance(node, ast.alias):
            names.add(node.name)
            if node.asname:
                names.add(node.asname)
        elif isinstance(node, ast.ImportFrom) and node.module:
            names.add(node.module)
    return names


_ENGINE_MODULES = (SIM_ORACLE, SIM_STATS)


def test_31_the_engine_touches_no_workbook_or_com_object() -> None:
    forbidden = (
        "Range", "ListObject", "ListObjects", "Application", "ThisWorkbook",
        "ActiveWorkbook", "Workbook", "Workbooks", "Worksheet", "Worksheets",
        "Cells", "Dispatch", "EnsureDispatch", "load_workbook", "openpyxl",
        "win32com", "comtypes", "pythoncom", "xlwings",
    )
    for path in _ENGINE_MODULES:
        names = _semantic_identifiers(path)
        for token in forbidden:
            assert token not in names, f"{path.name} uses {token!r}"


def test_32_the_engine_imports_no_evidence_and_no_emission_machinery() -> None:
    forbidden = (
        "evidence", "vba_source", "stage_b_emit", "emit_stage_b", "workbook_builder",
        "build_workbook", "calc_emit", "emit_calc_artifacts", "gate_b_inspection",
        "numpy", "random", "statistics",
    )
    for path in _ENGINE_MODULES:
        names = _semantic_identifiers(path)
        for token in forbidden:
            assert token not in names, f"{path.name} references {token!r}"
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if isinstance(node, (ast.Import, ast.ImportFrom)):
                module = getattr(node, "module", "") or ""
                assert "evidence" not in module, path.name
                for alias in node.names:
                    assert "evidence" not in alias.name, path.name


def test_33_the_engine_publishes_nothing() -> None:
    """No `_SimData`, no Results, no status, no counter, no run id, no VBA."""
    forbidden = (
        "_SimData", "SimData", "write_sim_data", "publish", "persist",
        "next_auto_nonce", "run_id", "allocate_run_id", "simulation_status",
        "attempt_metadata", "PCCM_RunSimulation", "write_text", "write_bytes", "save",
    )
    for path in _ENGINE_MODULES:
        names = _semantic_identifiers(path)
        for token in forbidden:
            assert token not in names, f"{path.name} uses {token!r}"


def test_34_the_engine_contains_no_sensitivity_or_annual_stochastic_machinery() -> None:
    forbidden = (
        "sensitivity", "tornado", "annual_samples", "annual_matrix", "annual",
        "sample_matrix", "per_driver_samples", "driver_samples", "_annual_series",
    )
    for path in _ENGINE_MODULES:
        names = _semantic_identifiers(path)
        for token in forbidden:
            assert token not in names, f"{path.name} references {token!r}"


def _sequence_lengths(value, depth: int = 0) -> list:
    """Every sequence length reachable in a result tree, to a bounded depth."""
    if depth > 4:
        return []
    found: list = []
    if isinstance(value, (tuple, list)):
        found.append(len(value))
        for item in value[:8]:
            found.extend(_sequence_lengths(item, depth + 1))
    elif isinstance(value, dict):
        found.append(len(value))
    elif hasattr(value, "__dataclass_fields__"):
        for field in value.__dataclass_fields__:
            found.extend(_sequence_lengths(getattr(value, field), depth + 1))
    return found


def test_35_no_per_driver_per_iteration_sample_matrix_survives_a_run() -> None:
    """The only long arrays are the two retained totals.

    Five drivers over 1,000 iterations would give a 5,000-element sample matrix
    and seven components a 7,000-element one. Neither length exists anywhere in
    the result, at any depth.
    """
    iterations = 1000
    prepared, _ = _prepare(_mixed(), iterations=iterations)
    run = run_simulation(_ref(), prepared)

    drivers = len(prepared.cost_drivers) + len(prepared.risk_drivers)
    components = len(run.diagnostics)
    assert drivers == 5 and components == 7

    lengths = _sequence_lengths(run)
    assert lengths.count(iterations) == 2, (
        f"expected exactly two arrays of length {iterations}, saw {lengths.count(iterations)}"
    )
    assert max(lengths) == iterations
    for forbidden in (drivers * iterations, components * iterations, 2 * drivers * iterations):
        assert forbidden not in lengths, f"a matrix of {forbidden} values was retained"


def test_36_the_prepared_model_retains_no_worksheet_or_workbook_object() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    allowed = {"builtins", "pccm_builder.sim_rng", "pccm_builder.sim_sample"}
    for driver in prepared.drivers:
        for field in driver.__dataclass_fields__:
            value = getattr(driver, field)
            assert not hasattr(value, "Range"), field
            assert not hasattr(value, "ListObjects"), field
            assert type(value).__module__ in allowed, (field, type(value))
    assert not hasattr(prepared, "worksheet")
    assert not hasattr(prepared, "workbook")


def test_37_the_engine_holds_no_module_level_mutable_simulation_state() -> None:
    """No module-level list, dict or set exists for a run to accumulate into."""
    import pccm_builder.sim_oracle as engine
    import pccm_builder.sim_stats as statistics_module

    for module in (engine, statistics_module):
        mutable = [
            name for name in dir(module)
            if not name.startswith("__")
            and isinstance(getattr(module, name, None), (list, dict, set))
        ]
        assert not mutable, f"{module.__name__} holds mutable module state: {mutable}"

    tree = ast.parse(SIM_ORACLE.read_text(encoding="utf-8"))
    for node in tree.body:
        if isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = node.targets if isinstance(node, ast.Assign) else [node.target]
            for target in targets:
                assert isinstance(target, ast.Name), ast.dump(target)
                assert isinstance(node.value, (ast.Constant, ast.Call)), (
                    f"module-level {target.id} is not a constant or a compiled literal"
                )


# ===========================================================================
# RNG reference binding
#
# A prepared model records `rng_version = 1`. Without a binding that claim is
# unenforceable: the same prepared model run under a reference with `a12 + 1`
# produced a completely different digest and still reported version 1. Two
# different generators cannot both be version 1.
# ===========================================================================
class _BehaviourOverride(RngReference):
    """Every accepted constant, a completely different generator.

    This is the hole a field snapshot cannot see: `rng_reference_signature`
    returns the SAME tuple for this object and for the accepted reference,
    because every operational field is identical. Only the method body differs -
    and the method body is the generator. On one Uniform Cost Line the accepted
    reference produces `88.8907785432604, 102.29692957777561, ...` and digest
    `30F068CF20B784C6`; this one produces `115.0` forever and digest
    `03B41CA3043133E0`. Both would claim `rng_version = 1`.
    """

    def next_uniform(self, state):  # type: ignore[override]
        self.validate_state(state)
        return Draw(state, 0.5)


def _override_reference() -> RngReference:
    base = _ref()
    return _BehaviourOverride(
        **{field.name: getattr(base, field.name) for field in dataclasses.fields(base)}
    )


def _variant(**changes) -> RngReference:
    """The accepted reference with exactly one operational field replaced.

    Always the EXACT accepted class, so these controls exercise the field
    signature rather than the implementation-identity check.
    """
    base = _ref()
    values = {field.name: getattr(base, field.name) for field in dataclasses.fields(base)}
    values.update(changes)
    return RngReference(**values)


def _bumped_jump(reference: RngReference) -> tuple:
    rows = [list(row) for row in reference.jump_a1]
    rows[1][2] = (rows[1][2] + 1) % reference.m1
    return tuple(tuple(row) for row in rows)


def _wrong_references() -> tuple:
    """Every operational axis the field signature is required to cover."""
    reference = _ref()
    return (
        ("recurrence a12", _variant(a12=reference.a12 + 1)),
        ("recurrence a13n", _variant(a13n=reference.a13n + 1)),
        ("modulus m1", _variant(m1=reference.m1 - 1)),
        ("normalisation", _variant(norm=reference.norm * (1.0 + 2.0e-16))),
        ("jump matrix element", _variant(jump_a1=_bumped_jump(reference))),
        ("AUTO multiplier", _variant(auto_multiplier=reference.auto_multiplier + 1)),
        ("seed domain", _variant(seed_max=reference.seed_max - 1)),
        ("kind order", _variant(kind_order=tuple(reversed(reference.kind_order)))),
        (
            "role order",
            _variant(
                role_order={
                    kind: (tuple(reversed(roles)) if kind == "RISK" else roles)
                    for kind, roles in reference.role_order.items()
                }
            ),
        ),
    )


def test_38_a_prepared_model_refuses_a_different_reference_at_run_time() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    accepted = run_simulation(_ref(), prepared)

    for label, wrong in _wrong_references():
        try:
            run_simulation(wrong, prepared)
        except SimOracleError as error:
            assert "prepared against" in str(error), (label, str(error))
            continue
        raise AssertionError(f"{label}: the run accepted a foreign reference")

    # The control is not vacuous: the accepted reference still runs, unchanged.
    assert run_simulation(_ref(), prepared).result_digest == accepted.result_digest


def test_39_the_binding_is_checked_before_the_first_draw() -> None:
    """Proved by instrumenting the ACCEPTED class, not by accepting a subclass.

    `RngReference.next_uniform` is replaced for the duration of the call so that
    any draw at all raises. A foreign reference must still produce the BINDING
    refusal, which can only happen if the binding is checked first. The spy is
    then shown to be live: the same instrumented class, handed the reference the
    model was actually prepared against, does reach the sampler.
    """
    prepared, _ = _prepare(_mixed(), iterations=1000)
    wrong = _variant(a12=_ref().a12 + 1)
    real_next_uniform = RngReference.next_uniform

    def refuse_to_draw(self, state):
        raise AssertionError("a draw was taken before the binding was verified")

    RngReference.next_uniform = refuse_to_draw          # type: ignore[assignment]
    try:
        try:
            run_simulation(wrong, prepared)
        except SimOracleError as error:
            assert "a12" in str(error), str(error)
        else:
            raise AssertionError("a foreign reference was accepted")

        # The spy is live, so the test above proved ORDERING rather than nothing.
        try:
            run_simulation(_ref(), prepared)
        except AssertionError as error:
            assert "before the binding was verified" in str(error)
        else:
            raise AssertionError("the instrumented sampler was never reached")
    finally:
        RngReference.next_uniform = real_next_uniform   # type: ignore[assignment]

    # And the class is intact afterwards.
    assert run_simulation(_ref(), prepared).total_nominal == run_simulation(
        _ref(), prepared
    ).total_nominal


def test_39a_a_behavioural_subclass_cannot_claim_the_accepted_rng_version() -> None:
    """The field signature is identical; only the implementation differs.

    A snapshot of data cannot see a method body, so the boundaries require the
    EXACT accepted type. `RngReference` is the Step-2 oracle, not an extension
    point.
    """
    override = _override_reference()
    base = _ref()
    for field in dataclasses.fields(base):
        assert getattr(override, field.name) == getattr(base, field.name), field.name
    assert isinstance(override, RngReference), "the control is not a subclass"
    assert type(override) is not RngReference

    # Preparation: refused before stream construction and before any draw.
    try:
        prepare_simulation(
            override, _sim(), _inputs(), _mixed(), _tolerances(),
            effective_seed=12345, iterations=1000,
        )
    except SimOracleError as error:
        assert "exactly RngReference" in str(error), str(error)
    else:
        raise AssertionError("preparation accepted a behavioural subclass")

    # Run: refused too, on a model prepared against the accepted reference.
    prepared, _ = _prepare(_mixed(), iterations=1000)
    try:
        run_simulation(override, prepared)
    except SimOracleError as error:
        assert "exactly RngReference" in str(error), str(error)
    else:
        raise AssertionError("the run accepted a behavioural subclass")

    # And the signature helper itself refuses, so it cannot imply authority.
    try:
        rng_reference_signature(override)
    except SimOracleError as error:
        assert "exactly RngReference" in str(error)
        return
    raise AssertionError("the signature helper accepted a behavioural subclass")


def test_39b_the_subclass_really_is_a_different_generator() -> None:
    """The control is not vacuous: the two produce different numbers.

    Run through the SAMPLERS directly - the engine refuses the override, which
    is the point - so the divergence is demonstrated without the oracle ever
    accepting it.
    """
    override = _override_reference()
    state = _ref().fixed_seed_to_state(12345)

    accepted_uniform = _ref().next_uniform(state).uniform
    override_uniform = override.next_uniform(state).uniform
    assert override_uniform == 0.5
    assert accepted_uniform != 0.5

    accepted_sample = sample_uniform(_ref(), state, 80.0, 150.0, None).value
    override_sample = sample_uniform(override, state, 80.0, 150.0, None).value
    assert override_sample == 115.0, override_sample
    assert accepted_sample == 88.8907785432604, accepted_sample

    # Identical field snapshots, so no data check could have separated them.
    assert rng_reference_signature(_ref()) == rng_reference_signature(
        RngReference(
            **{field.name: getattr(override, field.name)
               for field in dataclasses.fields(_ref())}
        )
    )


def test_40_a_foreign_reference_is_refused_at_preparation() -> None:
    """Refused before stream construction and before any draw."""
    for label, wrong in _wrong_references():
        try:
            _prepare(_mixed(), iterations=1000)  # sanity: the accepted path works
            prepare_simulation(
                wrong, _sim(), _inputs(), _mixed(), _tolerances(),
                effective_seed=12345, iterations=1000,
            )
        except SimOracleError as error:
            assert "derive" in str(error), (label, str(error))
            continue
        raise AssertionError(f"{label}: preparation accepted a foreign reference")


def test_41_the_binding_refusal_names_the_field_that_moved() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    expected = {
        "recurrence a12": "a12",
        "jump matrix element": "jump.a1_p127",
        "seed domain": "seed_max",
        "role order": "components.role_order",
        "kind order": "components.kind_order",
    }
    for label, wrong in _wrong_references():
        if label not in expected:
            continue
        try:
            run_simulation(wrong, prepared)
        except SimOracleError as error:
            assert expected[label] in str(error), (label, str(error))
            continue
        raise AssertionError(f"{label} was accepted")


def test_42_the_divergence_the_binding_prevents_is_real() -> None:
    """The defect this closes, demonstrated: two generators, one version number.

    Each mutated reference is shown to be a genuinely different generator - a
    different uniform from the same state, or a different stream ladder from the
    same seed - while every one of them still carries the contract's
    `rng_version`. That is why mixing them has to be refused rather than
    detected afterwards.
    """
    prepared, _ = _prepare(_mixed(), seed=12345, iterations=1000)
    accepted = run_simulation(_ref(), prepared)
    assert prepared.rng_version == 1 and accepted.rng_version == 1

    start = _ref().fixed_seed_to_state(12345)
    baseline_uniform = _ref().next_uniform(start).uniform
    baseline_jump = _ref().jump_to_next_stream(start)

    diverged = 0
    for label, wrong in _wrong_references():
        assert rng_reference_signature(wrong) != prepared.rng_signature, label
        try:
            if wrong.next_uniform(start).uniform != baseline_uniform:
                diverged += 1
                continue
        except Exception:  # a modulus change can make the start state illegal
            diverged += 1
            continue
        if wrong.jump_to_next_stream(start) != baseline_jump:
            diverged += 1

    assert diverged >= 5, (
        f"only {diverged} of the mutated references changed a number; the control "
        "would be proving little"
    )


# ===========================================================================
# immutability of the successful result
# ===========================================================================
def test_43_a_reported_percentile_cannot_be_rewritten() -> None:
    stats = describe([float(index) for index in range(1000)], (("P50", 0.5),), "immutable")
    try:
        stats.percentiles["P50"] = -999.0
    except TypeError:
        pass
    else:
        raise AssertionError("a reported percentile was rewritten")
    assert not isinstance(stats.percentiles, dict)
    assert stats.percentiles["P50"] == percentile_type7(
        [float(index) for index in range(1000)], 0.5
    )


def test_44_a_manually_built_statistic_copies_its_percentile_mapping() -> None:
    """A caller that keeps its own dictionary cannot reach in afterwards."""
    owned = {"P50": 1.0, "P90": 2.0}
    stats = MeasureStatistics(
        count=3, mean=1.0, sample_standard_deviation=0.0,
        minimum=1.0, maximum=2.0, percentiles=owned,
    )
    owned["P50"] = -999.0
    assert stats.percentiles["P50"] == 1.0, "the caller's dictionary reached the record"
    try:
        stats.percentiles["P90"] = -999.0
    except TypeError:
        return
    raise AssertionError("a manually built record exposed a mutable mapping")


def test_45_mutating_reported_data_cannot_change_a_contingency() -> None:
    """The defect this closes: a rewritten P50 fed a contingency with no new
    simulation, no new seed and no new digest."""
    prepared, result = _prepare(_mixed(), iterations=1500)
    run = run_simulation(_ref(), prepared)
    base = deterministic_base_of(result)

    saved_digest = run.result_digest
    saved_p50 = run.summary.nominal.percentiles["P50"]
    saved_contingency = contingency_at(run.summary, "P50", base).nominal

    try:
        run.summary.nominal.percentiles["P50"] = -999.0
    except TypeError:
        pass
    else:
        raise AssertionError("the reported P50 was rewritten after the run")

    assert run.result_digest == saved_digest
    assert run.summary.nominal.percentiles["P50"] == saved_p50
    assert contingency_at(run.summary, "P50", base).nominal == saved_contingency
    assert saved_contingency != -999.0 - base.nominal, "the control is vacuous"


def _mutable_containers(value, path: str = "run", depth: int = 0) -> list:
    """Every `list`, `dict` or `set` reachable in a result tree."""
    if depth > 6:
        return []
    if isinstance(value, (list, dict, set)):
        return [(path, type(value).__name__)]
    found: list = []
    if isinstance(value, tuple):
        for index, item in enumerate(value[:6]):
            found.extend(_mutable_containers(item, f"{path}[{index}]", depth + 1))
    elif hasattr(value, "__dataclass_fields__"):
        for field in value.__dataclass_fields__:
            found.extend(
                _mutable_containers(getattr(value, field), f"{path}.{field}", depth + 1)
            )
    return found


def test_46_the_whole_successful_result_tree_is_immutable() -> None:
    prepared, _ = _prepare(_mixed(), iterations=1000)
    run = run_simulation(_ref(), prepared)

    assert _mutable_containers(run, "run") == []
    assert _mutable_containers(prepared, "prepared") == []
    for record in run.diagnostics:
        assert isinstance(record.initial_state.words, tuple)
        assert isinstance(record.final_state.words, tuple)
    assert isinstance(run.total_nominal, tuple) and isinstance(run.total_pv, tuple)

    # The guard is not vacuous: it finds a mutable container when one is there.
    assert _mutable_containers(({"a": 1},), "probe") == [("probe[0]", "dict")]


# ===========================================================================
# the constant-sample invariant, as a control
# ===========================================================================
def test_47_manufactured_dispersion_on_a_constant_sample_is_caught() -> None:
    """The numbers the accepted implementation used to produce, named.

    The normalised two-pass path is still exactly what runs for a sample that
    varies; the invariant only intercepts one that does not.
    """
    values = [1.5e308] * 1000
    scale = 8.98846567431158e307                     # 2**1023
    scaled = [value / scale for value in values]
    total = safe_signed_sum(scaled, "drifted")
    centre = total / len(values)
    drifted_mean = centre * scale
    drifted_sd = math.sqrt(
        safe_signed_sum([(x - centre) * (x - centre) for x in scaled], "drifted")
        / (len(values) - 1)
    ) * scale

    assert drifted_mean == 1.4999999999999677e308, drifted_mean
    assert drifted_sd == 3.2348791455812365e294, drifted_sd
    assert sample_mean(values) == 1.5e308
    assert sample_standard_deviation(values) == 0.0

    ordinary = [0.1] * 1000
    naive = 0.0
    for value in ordinary:
        naive += value
    assert naive / len(ordinary) != 0.1, "the control is vacuous at ordinary scale"
    assert sample_mean(ordinary) == 0.1
    assert sample_standard_deviation(ordinary) == 0.0


def test_48_the_invariant_does_not_flatten_a_real_dispersion() -> None:
    """One ulp of variation is a real dispersion and is not intercepted.

    A NOTE ON WHAT IS AND IS NOT PROMISED. Once a sample varies, the ordinary
    left-to-right accumulation is what runs, and it carries the usual
    `O(n * eps)` relative drift. At near-`Double`-maximum magnitudes that drift
    can put the mean marginally OUTSIDE the closed interval `[min, max]` - for
    `[1.5e308] * 999 + [nextafter(1.5e308, 0)]` the mean comes back as
    `1.4999999999999677e308`, about `2e-14` relative below the minimum. That is
    accumulation arithmetic, not a scale-safety defect, and removing it would
    take a compensated summation this module deliberately does not own. The
    bracket is therefore asserted to within that relative drift, and stated
    rather than hidden.
    """
    for values in (
        [1.0] * 999 + [math.nextafter(1.0, 2.0)],
        [1.5e308] * 999 + [math.nextafter(1.5e308, 0.0)],
        [1.0e-300] * 999 + [2.0e-300],
    ):
        assert len(set(values)) == 2
        assert sample_standard_deviation(values) > 0.0, values[0]
        # The dispersion survives even where the MEAN cannot show it: one ulp in
        # a thousand observations is below the resolution of the average. The
        # percentiles still separate the two distinct values exactly.
        assert percentile_type7(values, 0.0) != percentile_type7(values, 1.0)
        assert percentile_type7(values, 0.0) == min(values)
        assert percentile_type7(values, 1.0) == max(values)

        mean = sample_mean(values)
        low, high = min(values), max(values)
        allowance = 1e-12 * max(abs(low), abs(high))
        assert low - allowance <= mean <= high + allowance, (values[0], mean)


def test_49_a_dispersion_with_no_double_is_refused_not_reported_as_zero() -> None:
    """The invariant must not be reachable by underflow.

    `[5e-324] * 999 + [1e-323]` varies, so it is NOT a constant sample - but its
    true sample deviation is about `1.6e-325`, below the smallest subnormal. The
    only two wrong answers are `0.0`, which would claim the sample has no
    dispersion, and a silent underflow. It is refused instead, naming the stage.
    """
    values = [5e-324] * 999 + [1e-323]
    assert len(set(values)) == 2

    try:
        sample_standard_deviation(values)
    except NumericalRangeRefusal as error:
        assert "rescale" in str(error), str(error)
    else:
        raise AssertionError("an unrepresentable deviation was returned")

    # The mean of the same sample IS representable and is produced.
    assert 5e-324 <= sample_mean(values) <= 1e-323


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
