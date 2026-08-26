#!/usr/bin/env python3
"""PCCM Phase 6 Step-3 conformance tests for the Python sampler reference.

`builder/pccm_builder/sim_sample.py` turns raw MRG32k3a uniforms into
distribution samples and Bernoulli decisions. These tests prove it reproduces the
retained Step-0 Cheng vectors EXACTLY - value, proposal attempts, uniforms
consumed, cumulative uniforms and the post-sample six-word state - and that every
family's RNG consumption is what the contract says it is.

CONSUMPTION IS TESTED, NOT JUST VALUES. Under a rejection sampler the number of
uniforms consumed depends on the values drawn, so a later implementation has to
reproduce the consumption as well as the sample; a test that checked only the
number would pass on an implementation that desynchronises the stream.

The retained evidence is a TEST oracle only; a separate test asserts the sampler
reads nothing at run time.

NO MONTE CARLO HERE. No iteration engine, no contribution, no `Quantity`,
`Knom`, `Kpv`, totals, digest or statistic.

Runs standalone or under pytest.
"""

from __future__ import annotations

import ast
import json
import math
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    ACCEPTED_FAMILIES,
    RngReference,
    RngState,
    SimSampleError,
    bernoulli_occurs,
    load_contract,
    load_sim_contract,
    prepare_beta_pert,
    sample_beta_pert,
    sample_distribution,
    sample_prepared_beta,
    sample_triangular,
    sample_uniform,
)
from pccm_builder.sim_sample import (  # noqa: E402
    FAMILY_BETA_PERT,
    FAMILY_TRIANGULAR,
    FAMILY_UNIFORM,
    _bernoulli_from_u,
    _triangular_from_u,
    _uniform_from_u,
    is_degenerate,
)

CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0"
SIM_SAMPLE = PCCM_ROOT / "builder" / "pccm_builder" / "sim_sample.py"

_REF: RngReference | None = None


def _ref() -> RngReference:
    global _REF
    if _REF is None:
        _REF = RngReference.from_contracts(
            load_sim_contract(SIM_PATH), load_contract(CONTRACT_PATH)
        )
    return _REF


def _state(seed: int = 12345) -> RngState:
    return _ref().fixed_seed_to_state(seed)


def _cheng_vectors() -> dict:
    return json.loads(
        (EVIDENCE / "vectors" / "cheng_vectors.json").read_text(encoding="utf-8")
    )


# ===========================================================================
# A. dispatch
# ===========================================================================
def test_01_exactly_three_families_dispatch() -> None:
    assert ACCEPTED_FAMILIES == ("Uniform", "Triangular", "Beta-PERT")
    for family in ACCEPTED_FAMILIES:
        result = sample_distribution(_ref(), _state(), family, 0.0, 50.0, 100.0)
        assert 0.0 <= result.value <= 100.0, family


def test_02_an_unknown_family_is_refused() -> None:
    for bad in ("Normal", "uniform", "UNIFORM", "BetaPERT", "Beta PERT", "", None, 3):
        try:
            sample_distribution(_ref(), _state(), bad, 0.0, 50.0, 100.0)
        except SimSampleError:
            continue
        raise AssertionError(f"family {bad!r} was accepted")


# ===========================================================================
# B - C. degeneracy, and the ignored Uniform Most Likely
# ===========================================================================
def test_03_degeneracy_is_family_specific() -> None:
    assert is_degenerate(FAMILY_UNIFORM, 100.0, None, 100.0)
    assert is_degenerate(FAMILY_UNIFORM, 100.0, 500.0, 100.0), (
        "a populated, unrelated Most Likely must not make a Uniform non-degenerate"
    )
    assert is_degenerate(FAMILY_TRIANGULAR, 7.0, 7.0, 7.0)
    assert not is_degenerate(FAMILY_TRIANGULAR, 7.0, 8.0, 9.0)
    assert is_degenerate(FAMILY_BETA_PERT, 7.0, 7.0, 7.0)


def test_04_a_degenerate_driver_consumes_nothing_in_every_family() -> None:
    start = _state()
    for family, a, m, b in (
        (FAMILY_UNIFORM, 100.0, None, 100.0),
        (FAMILY_UNIFORM, 100.0, 500.0, 100.0),
        (FAMILY_UNIFORM, -3.5, None, -3.5),
        (FAMILY_TRIANGULAR, 7.0, 7.0, 7.0),
        (FAMILY_BETA_PERT, 7.0, 7.0, 7.0),
        (FAMILY_BETA_PERT, -2.0, -2.0, -2.0),
    ):
        result = sample_distribution(_ref(), start, family, a, m, b)
        assert result.value == a, (family, a)
        assert result.uniforms_consumed == 0
        assert result.proposal_attempts == 0
        assert result.state == start, "a degenerate driver advanced the stream"


def test_05_a_degenerate_beta_forms_no_r_and_no_parameterisation() -> None:
    """0/0 cannot arise, because `r` is never formed."""
    shape = prepare_beta_pert(7.0, 7.0, 7.0)
    assert shape.degenerate is True
    assert shape.alpha == 0.0 and shape.beta == 0.0
    assert shape.dispatch == ""


def test_06_the_ignored_uniform_most_likely_changes_nothing() -> None:
    """Same Min/Max, different Most Likely: identical value, state and consumption."""
    start = _state()
    baseline = sample_uniform(_ref(), start, 10.0, 90.0, None)
    for ignored in (None, 0.0, 50.0, 1e300, -1e300, 90.0, 10.0):
        result = sample_uniform(_ref(), start, 10.0, 90.0, ignored)
        assert float.hex(result.value) == float.hex(baseline.value), ignored
        assert result.state == baseline.state, ignored
        assert result.uniforms_consumed == baseline.uniforms_consumed
    # And through the dispatcher, which is what a driver actually calls.
    through = sample_distribution(_ref(), start, FAMILY_UNIFORM, 10.0, 12345.0, 90.0)
    assert float.hex(through.value) == float.hex(baseline.value)
    assert through.state == baseline.state


# ===========================================================================
# D - E. Uniform and Triangular transforms
# ===========================================================================
def test_07_uniform_consumes_exactly_one_uniform() -> None:
    start = _state()
    result = sample_uniform(_ref(), start, 0.0, 100.0)
    assert result.uniforms_consumed == 1
    assert result.proposal_attempts == 0
    assert result.state == _ref().next_uniform(start).state


def test_08_the_uniform_transform_is_the_convex_form() -> None:
    for u in (0.001, 0.25, 0.5, 0.75, 0.999):
        for a, b in ((0.0, 100.0), (-100.0, 100.0), (-90.0, -10.0), (1e-300, 2e-300)):
            expected = (1.0 - u) * a + u * b
            assert float.hex(_uniform_from_u(u, a, b)) == float.hex(expected)


def test_09_uniform_supports_negative_and_zero_crossing_ranges() -> None:
    for a, b in ((-100.0, -10.0), (-50.0, 50.0), (-1.0, 0.0), (0.0, 1.0)):
        result = sample_uniform(_ref(), _state(), a, b)
        assert a <= result.value <= b, (a, b)


def test_10_uniform_stays_finite_over_the_extreme_domain() -> None:
    """The convex form's whole purpose: `a + u*(b-a)` overflows here."""
    a, b = -1.0e308, 1.0e308
    assert math.isinf(b - a), "the naive difference should overflow, or this proves nothing"
    for u in (0.001, 0.5, 0.999):
        value = _uniform_from_u(u, a, b)
        assert math.isfinite(value), u
        assert a <= value <= b


def test_11_triangular_consumes_exactly_one_uniform() -> None:
    start = _state()
    result = sample_triangular(_ref(), start, 0.0, 30.0, 100.0)
    assert result.uniforms_consumed == 1
    assert result.state == _ref().next_uniform(start).state


def test_12_the_triangular_branch_point_and_its_neighbourhood() -> None:
    a, m, b = 0.0, 30.0, 100.0
    c = (m - a) / (b - a)
    at = _triangular_from_u(c, a, m, b)
    assert math.isclose(at, m, rel_tol=1e-12), f"u = c must land on the mode, got {at}"
    below = _triangular_from_u(math.nextafter(c, 0.0), a, m, b)
    above = _triangular_from_u(math.nextafter(c, 1.0), a, m, b)
    assert below <= at <= above
    # `u <= c` takes the LOWER branch; strictness at the boundary is contracted.
    assert _triangular_from_u(c, a, m, b) == at


def test_13_triangular_boundary_modes() -> None:
    """`m = a` -> c = 0, upper branch always; `m = b` -> c = 1, lower branch always."""
    for u in (0.001, 0.5, 0.999):
        left = _triangular_from_u(u, 0.0, 0.0, 100.0)
        assert 0.0 <= left <= 100.0
        right = _triangular_from_u(u, 0.0, 100.0, 100.0)
        assert 0.0 <= right <= 100.0
    # m = a: c = 0, so every u > 0 takes the UPPER branch, evaluated in the
    # conditioned space (s = 100) and rescaled after.
    assert _triangular_from_u(0.5, 0.0, 0.0, 100.0) == (1.0 - math.sqrt(0.5)) * 100.0
    # m = b: c = 1, so every u <= 1 takes the LOWER branch.
    assert _triangular_from_u(0.5, 0.0, 100.0, 100.0) == math.sqrt(0.5) * 100.0
    # And the branch actually taken is the contracted one, not merely a value
    # that happens to match: at m = a the lower branch would give `a` exactly.
    assert _triangular_from_u(0.5, 0.0, 0.0, 100.0) != 0.0


def test_14_triangular_supports_negative_and_zero_crossing_ranges() -> None:
    for a, m, b in ((-100.0, -50.0, -10.0), (-50.0, 0.0, 50.0), (-10.0, -10.0, 10.0)):
        for u in (0.001, 0.5, 0.999):
            value = _triangular_from_u(u, a, m, b)
            assert a <= value <= b, (a, m, b, u)


def test_15_triangular_stays_finite_over_the_extreme_domain() -> None:
    a, m, b = -1.0e308, 0.0, 1.0e308
    assert math.isinf((b - a) * (m - a)) or math.isinf(b - a)
    for u in (0.001, 0.25, 0.5, 0.75, 0.999):
        value = _triangular_from_u(u, a, m, b)
        assert math.isfinite(value), u
        assert a <= value <= b
    # Subnormal scale.
    small = _triangular_from_u(0.5, 5e-324, 1e-323, 1.5e-323)
    assert math.isfinite(small)


# ===========================================================================
# G - I. Beta-PERT parameterisation and dispatch
# ===========================================================================
def test_16_the_pert_family_invariant_holds() -> None:
    for r in (0.0, 0.01, 0.1, 0.25, 0.5, 0.75, 0.9, 0.99, 1.0):
        shape = prepare_beta_pert(0.0, r, 1.0)
        assert math.isclose(shape.alpha, 1.0 + 4.0 * r, rel_tol=1e-15)
        assert math.isclose(shape.beta, 1.0 + 4.0 * (1.0 - r), rel_tol=1e-15)
        assert math.isclose(shape.alpha + shape.beta, 6.0, rel_tol=1e-15)
        assert 1.0 <= shape.alpha <= 5.0 and 1.0 <= shape.beta <= 5.0


def test_17_dispatch_sends_the_interior_to_bb_and_the_boundary_to_bc() -> None:
    for r in (0.01, 0.1, 0.25, 0.5, 0.75, 0.9, 0.99):
        assert prepare_beta_pert(0.0, r, 1.0).dispatch == "BB", r
    # EQUALITY BELONGS TO BC, and both endpoints reach it by the rule - not by a
    # special case bolted on afterwards.
    at_a = prepare_beta_pert(0.0, 0.0, 1.0)
    at_b = prepare_beta_pert(0.0, 1.0, 1.0)
    assert (at_a.alpha, at_a.beta, at_a.dispatch) == (1.0, 5.0, "BC")
    assert (at_b.alpha, at_b.beta, at_b.dispatch) == (5.0, 1.0, "BC")


def test_18_the_two_dispatches_orient_oppositely() -> None:
    """BB uses (min, max); BC uses (max, min). Inverting one mirrors the result."""
    bb = prepare_beta_pert(0.0, 0.25, 1.0)
    assert (bb.cheng_a, bb.cheng_b) == (min(bb.alpha, bb.beta), max(bb.alpha, bb.beta))
    bc = prepare_beta_pert(0.0, 0.0, 1.0)
    assert (bc.cheng_a, bc.cheng_b) == (max(bc.alpha, bc.beta), min(bc.alpha, bc.beta))


def test_19_preparing_a_shape_draws_nothing() -> None:
    """The whole point of preparing once per driver."""
    start = _state()
    prepare_beta_pert(0.0, 0.25, 1.0)
    assert _state() == start


# ===========================================================================
# J - L. the retained Cheng vectors
# ===========================================================================
def test_20_the_retained_cheng_cases_are_the_required_five() -> None:
    labels = [c["label"] for c in _cheng_vectors()["cases"]]
    assert len(labels) == 5
    dispatches = [c["dispatch"] for c in _cheng_vectors()["cases"]]
    assert dispatches.count("BB") == 3 and dispatches.count("BC") == 2
    shapes = [(c["alpha"], c["beta"]) for c in _cheng_vectors()["cases"]]
    assert (2.0, 4.0) in shapes and (3.0, 3.0) in shapes and (1.04, 4.96) in shapes
    assert (1.0, 5.0) in shapes and (5.0, 1.0) in shapes


def test_21_every_retained_cheng_sample_reproduces_exactly() -> None:
    """Value, attempts, uniforms, cumulative uniforms and post-state - all five.

    `a = 0, b = 1` makes the convex rescale the identity, so the retained Beta
    variate is directly comparable.
    """
    for case in _cheng_vectors()["cases"]:
        shape = prepare_beta_pert(0.0, case["r"], 1.0)
        assert shape.dispatch == case["dispatch"], case["label"]
        assert math.isclose(shape.alpha, case["alpha"], rel_tol=1e-15), case["label"]
        state = RngState(tuple(case["initial_state"]))
        cumulative = 0
        for retained in case["samples"]:
            result = sample_prepared_beta(_ref(), state, shape)
            cumulative += result.uniforms_consumed
            where = f"{case['label']} sample {retained['index']}"
            assert repr(result.value) == retained["accepted_sample"], where
            assert result.proposal_attempts == retained[
                "proposal_attempts_for_this_sample"
            ], where
            assert result.uniforms_consumed == retained["uniforms_for_this_sample"], where
            assert cumulative == retained["cumulative_uniforms"], where
            assert result.state.as_list() == retained["rng_state_after_sample"], where
            state = result.state
        assert state.as_list() == case["final_state"], case["label"]
        assert cumulative == case["total_uniforms"], case["label"]


def test_22_the_retained_vectors_exercise_retries_and_immediate_acceptance() -> None:
    """Without a retry in the corpus, the consumption rule would be untested."""
    for case in _cheng_vectors()["cases"]:
        attempts = [s["proposal_attempts_for_this_sample"] for s in case["samples"]]
        assert 1 in attempts, case["label"]
        assert any(a > 1 for a in attempts), case["label"]


def test_23_two_uniforms_per_proposal_attempt_always() -> None:
    for case in _cheng_vectors()["cases"]:
        shape = prepare_beta_pert(0.0, case["r"], 1.0)
        state = RngState(tuple(case["initial_state"]))
        for _ in range(len(case["samples"])):
            result = sample_prepared_beta(_ref(), state, shape)
            assert result.uniforms_consumed == 2 * result.proposal_attempts
            state = result.state


def test_24_a_rejected_proposal_advances_the_state_by_what_it_consumed() -> None:
    """No rewind. The retry continues from where the rejection left the stream."""
    case = next(c for c in _cheng_vectors()["cases"] if c["dispatch"] == "BC")
    shape = prepare_beta_pert(0.0, case["r"], 1.0)
    state = RngState(tuple(case["initial_state"]))
    for retained in case["samples"]:
        before = state
        result = sample_prepared_beta(_ref(), state, shape)
        _, expected = _ref().uniforms(before, result.uniforms_consumed)
        assert result.state == expected, retained["index"]
        state = result.state


def test_25_sample_beta_pert_matches_the_prepared_path() -> None:
    shape = prepare_beta_pert(0.0, 0.25, 1.0)
    start = _state()
    prepared = sample_prepared_beta(_ref(), start, shape)
    direct = sample_beta_pert(_ref(), start, 0.0, 0.25, 1.0)
    assert float.hex(direct.value) == float.hex(prepared.value)
    assert direct.state == prepared.state
    assert direct.uniforms_consumed == prepared.uniforms_consumed
    assert direct.proposal_attempts == prepared.proposal_attempts


# ===========================================================================
# M - N. rescale over the accepted domain
# ===========================================================================
def test_26_beta_rescales_onto_negative_and_zero_crossing_supports() -> None:
    state = _state()
    for a, m, b in (
        (-100.0, -50.0, -10.0),
        (-50.0, 0.0, 50.0),
        (-10.0, -10.0, 10.0),
        (-10.0, 10.0, 10.0),
        (100.0, 150.0, 200.0),
    ):
        result = sample_beta_pert(_ref(), state, a, m, b)
        assert a <= result.value <= b, (a, m, b, result.value)
        assert math.isfinite(result.value)


def test_27_beta_stays_finite_over_the_extreme_domain() -> None:
    a, m, b = -1.0e308, 0.0, 1.0e308
    result = sample_beta_pert(_ref(), _state(), a, m, b)
    assert math.isfinite(result.value)
    assert a <= result.value <= b
    small = sample_beta_pert(_ref(), _state(), 5e-324, 1e-323, 1.5e-323)
    assert math.isfinite(small.value)


def test_28_the_endpoint_modes_rescale_correctly() -> None:
    for a, m, b in ((0.0, 0.0, 100.0), (0.0, 100.0, 100.0), (-5.0, -5.0, 5.0)):
        shape = prepare_beta_pert(a, m, b)
        assert shape.dispatch == "BC", (a, m, b)
        result = sample_beta_pert(_ref(), _state(), a, m, b)
        assert a <= result.value <= b


# ===========================================================================
# O. Bernoulli
# ===========================================================================
def test_29_bernoulli_consumes_exactly_one_uniform() -> None:
    start = _state()
    result = bernoulli_occurs(_ref(), start, 0.5)
    assert result.uniforms_consumed == 1
    assert result.state == _ref().next_uniform(start).state
    assert result.uniform == _ref().next_uniform(start).uniform


def test_30_probability_zero_never_occurs_and_one_always_does() -> None:
    """Exactly, with no special case - because raw MRG output is inside (0,1)."""
    state = _state()
    for _ in range(500):
        never = bernoulli_occurs(_ref(), state, 0.0)
        always = bernoulli_occurs(_ref(), state, 1.0)
        assert never.occurred is False
        assert always.occurred is True
        state = never.state


def test_31_the_bernoulli_comparison_is_strict() -> None:
    p = 0.25
    assert _bernoulli_from_u(math.nextafter(p, 0.0), p) is True
    assert _bernoulli_from_u(p, p) is False, "u == p must NOT occur under strict <"
    assert _bernoulli_from_u(math.nextafter(p, 1.0), p) is False
    assert _bernoulli_from_u(0.0, 0.0) is False
    assert _bernoulli_from_u(math.nextafter(1.0, 0.0), 1.0) is True


def test_32_the_occurrence_rate_tracks_the_probability() -> None:
    state = _state()
    for p in (0.1, 0.5, 0.9):
        hits = 0
        current = state
        for _ in range(20000):
            result = bernoulli_occurs(_ref(), current, p)
            hits += result.occurred
            current = result.state
        assert abs(hits / 20000 - p) < 0.02, p


def test_33_an_out_of_range_probability_is_refused_not_clamped() -> None:
    for bad in (-0.001, 1.001, -1, 2, float("nan"), float("inf"), "0.5", None, True):
        try:
            bernoulli_occurs(_ref(), _state(), bad)
        except SimSampleError:
            continue
        raise AssertionError(f"probability {bad!r} was accepted")


# ===========================================================================
# P. parameter refusal
# ===========================================================================
def test_34_misordered_parameters_are_refused_not_repaired() -> None:
    cases = (
        (FAMILY_UNIFORM, 100.0, None, 10.0),
        (FAMILY_TRIANGULAR, 0.0, 200.0, 100.0),
        (FAMILY_TRIANGULAR, 0.0, -5.0, 100.0),
        (FAMILY_BETA_PERT, 0.0, 200.0, 100.0),
        (FAMILY_BETA_PERT, 100.0, 50.0, 0.0),
    )
    for family, a, m, b in cases:
        try:
            sample_distribution(_ref(), _state(), family, a, m, b)
        except SimSampleError:
            continue
        raise AssertionError(f"{family} accepted {a}, {m}, {b}")


def test_35_non_finite_and_malformed_parameters_are_refused() -> None:
    for value in (float("nan"), float("inf"), float("-inf"), "10", None, True):
        for family in (FAMILY_UNIFORM, FAMILY_TRIANGULAR, FAMILY_BETA_PERT):
            try:
                sample_distribution(_ref(), _state(), family, value, 50.0, 100.0)
            except SimSampleError:
                continue
            raise AssertionError(f"{family} accepted Min = {value!r}")


def test_36_a_missing_most_likely_is_refused_for_the_families_that_need_one() -> None:
    for family in (FAMILY_TRIANGULAR, FAMILY_BETA_PERT):
        try:
            sample_distribution(_ref(), _state(), family, 0.0, None, 100.0)
        except SimSampleError:
            continue
        raise AssertionError(f"{family} accepted a blank Most Likely")


def test_37_an_invalid_rng_state_is_refused_through_the_sampler() -> None:
    bad = RngState((0, 0, 0, 1, 1, 1))
    from pccm_builder import SimRngError

    for call in (
        lambda: sample_uniform(_ref(), bad, 0.0, 100.0),
        lambda: sample_triangular(_ref(), bad, 0.0, 50.0, 100.0),
        lambda: sample_beta_pert(_ref(), bad, 0.0, 50.0, 100.0),
        lambda: bernoulli_occurs(_ref(), bad, 0.5),
    ):
        try:
            call()
        except (SimSampleError, SimRngError):
            continue
        raise AssertionError("an invalid state reached a sampler")


def test_38_a_uniform_populated_most_likely_is_ignored_not_validated() -> None:
    """It is not a mode, so it is not order-checked."""
    result = sample_uniform(_ref(), _state(), 10.0, 90.0, 5000.0)
    assert 10.0 <= result.value <= 90.0
    result = sample_uniform(_ref(), _state(), 10.0, 90.0, -5000.0)
    assert 10.0 <= result.value <= 90.0


# ===========================================================================
# Q - S. scope discipline
# ===========================================================================
def test_39_the_sampler_reads_no_evidence_at_run_time() -> None:
    tree = ast.parse(SIM_SAMPLE.read_text(encoding="utf-8"))
    readers = ("open", "read_text", "read_bytes", "load", "loads", "glob", "rglob")
    for node in ast.walk(tree):
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            names = [a.name for a in node.names] + [getattr(node, "module", "") or ""]
            for name in names:
                assert "evidence" not in name and "phase6_step0" not in name, name
        if isinstance(node, ast.Call):
            name = getattr(node.func, "attr", None) or getattr(node.func, "id", None)
            assert name not in readers, f"sim_sample.py calls {name}() - it reads nothing"


def test_40_the_sampler_uses_no_third_party_or_stdlib_rng() -> None:
    source = SIM_SAMPLE.read_text(encoding="utf-8")
    tree = ast.parse(source)
    imported = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            imported.update(a.name.split(".")[0] for a in node.names)
        elif isinstance(node, ast.ImportFrom) and node.module:
            imported.add(node.module.split(".")[0])
    for banned in ("random", "secrets", "numpy", "scipy"):
        assert banned not in imported, f"sim_sample.py imports {banned}"
    for banned in ("default_rng", "Randomize", "Rnd(", "random.random"):
        assert banned not in source, banned


def test_41_no_monte_carlo_or_contribution_implementation_exists() -> None:
    """Step 3 provides the pieces. Step 4 orchestrates them."""
    source = SIM_SAMPLE.read_text(encoding="utf-8")
    tree = ast.parse(source)
    defined = {
        node.name
        for node in ast.walk(tree)
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    }
    for banned in (
        "iterate", "run_simulation", "simulate", "contribution", "accumulate",
        "total_nominal", "total_pv", "percentile", "result_digest", "contingency",
        "statistics",
    ):
        assert not any(banned in name for name in defined), f"sim_sample.py defines {banned}"
    # SEMANTIC, not a substring scan: the module's own docstring lists what it
    # must not implement, which is exactly the sentence a text scan trips over.
    identifiers = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Name):
            identifiers.add(node.id)
        elif isinstance(node, ast.Attribute):
            identifiers.add(node.attr)
        elif isinstance(node, ast.arg):
            identifiers.add(node.arg)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            identifiers.add(node.name)
    for banned in ("Quantity", "Knom", "Kpv", "total_nominal", "total_pv",
                   "result_digest", "quantity", "knom", "kpv"):
        assert banned not in identifiers, f"sim_sample.py binds {banned}"
    # No loop over iterations: the only loops are a rejection retry and a
    # parameter sweep, neither of which is a Monte Carlo.
    assert "for iteration" not in source
    assert "_SimData" not in {n for n in identifiers}


def test_42_the_sampler_holds_no_global_mutable_state() -> None:
    source = SIM_SAMPLE.read_text(encoding="utf-8")
    tree = ast.parse(source)
    for node in tree.body:
        if isinstance(node, ast.Assign):
            for target in node.targets:
                if isinstance(target, ast.Name):
                    assert target.id.isupper() or target.id.startswith("_"), target.id
    assert not [n for n in ast.walk(tree) if isinstance(n, (ast.Global, ast.Nonlocal))]


def test_43_no_sampler_vba_exists() -> None:
    """Step 7 authorised the sampler module. It authorised nothing beyond it.

    Cheng is named in modSimSample and NOWHERE ELSE - modSimRng included, which
    is the generator backbone and knows no distribution. The D6-11 algorithm
    token still lives in modSimRng alone: modSimSample consumes randomness
    through the modSimRng public surface and is granted no exception of its own.
    """
    src = PCCM_ROOT / "src" / "vba"
    names = {p.name for p in src.glob("*.bas")}
    assert "modSimSample.bas" in names
    for banned in ("modSimStats.bas",
                   "modSimFingerprint.bas", "modSimReport.bas"):
        assert banned not in names, banned
    from pccm_builder.vba_source import strip_comments, strip_strings

    for path in sorted(src.glob("*.bas")):
        raw = path.read_text(encoding="utf-8", errors="replace")
        # EXECUTABLE code, comments and string literals stripped - the same
        # discipline D6-11 enforcement uses. modSimEngine says in prose that it
        # contains no Cheng arithmetic, and a rule that stopped a module naming
        # what it refuses to contain would forbid the clearest thing it can say.
        code = strip_strings(strip_comments(raw))
        if path.stem != "modSimSample":
            assert "Cheng" not in code, path.name
        if path.stem != "modSimRng":
            assert "MRG32k3a" not in code, path.name


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
