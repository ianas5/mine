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
Substituting the built-in `sum` for the accepted accumulation changes nothing on
an ordinary fixture, and both halves of that fact are recorded rather than one.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import math
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
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
from pccm_builder.calc_numeric import safe_signed_sum  # noqa: E402
from pccm_builder.calc_oracle import (  # noqa: E402
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    FxRow,
    RiskDriver,
)
from pccm_builder.sim_oracle import (  # noqa: E402
    DeterministicBase,
    contingency_at,
    deterministic_base_of,
    prepare_simulation,
    result_digest,
    result_digest_stream,
    run_simulation,
)
from pccm_builder.sim_sample import (  # noqa: E402
    FAMILY_TRIANGULAR,
    FAMILY_UNIFORM,
)
from pccm_builder.sim_stats import (  # noqa: E402
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

        if mutation == "builtin_sum":
            nominal.append(sum(nominal_terms))
            pv.append(sum(pv_terms))
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


def test_15_the_builtin_sum_is_caught_only_on_the_constructed_fixture() -> None:
    """BOTH HALVES, recorded honestly.

    On an ordinary model `sum` and the accepted accumulation agree on every
    iteration, because tier 1 of `safe_signed_sum` IS left-to-right addition.
    The constructed fixture is the one where they part: a partial sum leaves
    Double range, `sum` reports infinity, and the accepted primitive returns the
    representable total the model actually has.
    """
    ordinary, _ = _prepare(_mixed(), iterations=1000)
    ordinary_run = run_simulation(_ref(), ordinary)
    ordinary_mutated, _ = _local_run(ordinary, mutation="builtin_sum")
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
    mutated, _ = _local_run(prepared, mutation="builtin_sum")

    assert set(run.total_nominal) == {1.5e308}
    assert all(math.isinf(total) for total in mutated), sorted(set(mutated))[:3]
    assert mutated != run.total_nominal


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

    assert sum(_STATS_SAMPLE) / len(_STATS_SAMPLE) == sample_mean(_STATS_SAMPLE), (
        "the two agree wherever the naive sum exists, which is the whole point"
    )


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


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
