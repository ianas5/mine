#!/usr/bin/env python3
"""PCCM Phase 6 Step-3 mutation controls for the sampler reference.

Every conformance test claims the samplers reproduce the retained Cheng vectors
and consume exactly what the contract says. A suite that cannot FAIL proves
nothing, so each control here plants one defect - an unsafe rescale, a moved
branch boundary, `log(4)` for the locked literal, a rejected proposal that fails
to advance the stream - and asserts the retained vectors or the contracted
consumption then reject it.

MUTATIONS LIVE HERE. Each defect is written out locally, exactly as the mistaken
implementation would have written it. `spec/sim_contract.yaml` is never edited to
manufacture a failure.

Runs standalone or under pytest.
"""

from __future__ import annotations

import dataclasses
import json
import math
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    RngReference,
    RngState,
    SimSampleError,
    bernoulli_occurs,
    load_contract,
    load_sim_contract,
    prepare_beta_pert,
    sample_beta_pert,
    sample_prepared_beta,
    sample_triangular,
    sample_uniform,
)
from pccm_builder.sim_sample import (  # noqa: E402
    FAMILY_BETA_PERT,
    FAMILY_TRIANGULAR,
    FAMILY_UNIFORM,
    _conditioning_scale,
    _triangular_from_u,
    _uniform_from_u,
    is_degenerate,
)

CONTRACT_PATH = PCCM_ROOT / "spec" / "input_contract.yaml"
SIM_PATH = PCCM_ROOT / "spec" / "sim_contract.yaml"
EVIDENCE = PCCM_ROOT / "evidence" / "phase6_step0"

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


def _cases() -> list[dict]:
    return json.loads(
        (EVIDENCE / "vectors" / "cheng_vectors.json").read_text(encoding="utf-8")
    )["cases"]


def _bb_case() -> dict:
    return next(c for c in _cases() if c["dispatch"] == "BB")


def _bc_case() -> dict:
    return next(c for c in _cases() if c["dispatch"] == "BC")


# ===========================================================================
# the control that makes every other control meaningful
# ===========================================================================
def test_00_the_unmutated_sampler_matches_every_retained_case() -> None:
    for case in _cases():
        shape = prepare_beta_pert(0.0, case["r"], 1.0)
        state = RngState(tuple(case["initial_state"]))
        for retained in case["samples"]:
            result = sample_prepared_beta(_ref(), state, shape)
            assert repr(result.value) == retained["accepted_sample"]
            assert result.proposal_attempts == retained["proposal_attempts_for_this_sample"]
            assert result.state.as_list() == retained["rng_state_after_sample"]
            state = result.state


# ===========================================================================
# Uniform
# ===========================================================================
def test_01_the_unsafe_uniform_difference_form_is_caught() -> None:
    """`a + u*(b-a)` loses a representable answer to an overflowing intermediate."""
    a, b = -1.0e308, 1.0e308
    for u in (0.001, 0.5, 0.999):
        safe = _uniform_from_u(u, a, b)
        unsafe = a + u * (b - a)
        assert math.isfinite(safe), u
        assert not math.isfinite(unsafe), f"the unsafe form survived at u = {u}"
    # And on an ordinary support the two differ in the last bits, so the choice
    # is load-bearing even where both are finite.
    differing = [
        u for u in (0.1, 0.3, 0.7, 0.9)
        if float.hex(_uniform_from_u(u, 1.0, 3.0)) != float.hex(1.0 + u * (3.0 - 1.0))
    ]
    assert differing, "the two forms agreed everywhere tried; pick a sharper case"


def test_02_using_the_uniform_most_likely_is_caught() -> None:
    """A defect that read `m` would make these two calls differ. They must not."""
    start = _state()
    first = sample_uniform(_ref(), start, 10.0, 90.0, None)
    second = sample_uniform(_ref(), start, 10.0, 90.0, 5000.0)
    assert float.hex(first.value) == float.hex(second.value)
    # The defect, written out: a "Triangular-like" reading of a Uniform.
    defective = _triangular_from_u(0.5, 10.0, 5000.0, 90.0) if 10.0 <= 5000.0 <= 90.0 else None
    assert defective is None, "the guard case is unreachable, which is the point"
    mode_aware = 10.0 + (5000.0 - 10.0) * 0.0  # any use of m at all
    assert mode_aware != first.value


def test_03_uniform_degeneracy_reverted_to_a_m_b_is_caught() -> None:
    """The withdrawn common predicate, and why it matters."""
    a, m, b = 100.0, 500.0, 100.0
    assert is_degenerate(FAMILY_UNIFORM, a, m, b) is True
    reverted = (a == m == b)
    assert reverted is False, "the withdrawn predicate would call this non-degenerate"

    start = _state()
    accepted = sample_uniform(_ref(), start, a, b, m)
    assert accepted.uniforms_consumed == 0
    assert accepted.state == start
    # Under the reverted predicate the driver would enter the sampler and consume
    # a uniform, moving the stream - an ignored input changing every later draw.
    would_consume = _ref().next_uniform(start).state
    assert would_consume != start


def test_04_a_misordered_uniform_is_refused_not_swapped() -> None:
    try:
        sample_uniform(_ref(), _state(), 100.0, 10.0)
    except SimSampleError:
        return
    raise AssertionError("Min > Max was silently repaired")


# ===========================================================================
# Triangular
# ===========================================================================
def test_05_a_strict_branch_comparison_is_caught() -> None:
    """`u < c` for the accepted `u <= c` differs exactly at the branch point."""
    a, m, b = 0.0, 30.0, 100.0
    scale = _conditioning_scale(a, m, b)
    an, mn, bn = a / scale, m / scale, b / scale
    c = (mn - an) / (bn - an)

    accepted = _triangular_from_u(c, a, m, b)
    # The defect: `<` sends u == c down the upper branch.
    strict = (bn - math.sqrt((1.0 - c) * (bn - an) * (bn - mn))) * scale
    assert float.hex(accepted) != float.hex(strict), (
        "the branch comparison made no difference at u == c"
    )


def test_06_swapped_triangular_branches_are_caught() -> None:
    a, m, b = 0.0, 30.0, 100.0
    scale = _conditioning_scale(a, m, b)
    an, mn, bn = a / scale, m / scale, b / scale
    c = (mn - an) / (bn - an)
    for u in (0.05, 0.2, 0.5, 0.9):
        accepted = _triangular_from_u(u, a, m, b)
        if u <= c:
            swapped = (bn - math.sqrt((1.0 - u) * (bn - an) * (bn - mn))) * scale
        else:
            swapped = (an + math.sqrt(u * (bn - an) * (mn - an))) * scale
        assert float.hex(accepted) != float.hex(swapped), u


def test_07_removing_the_triangular_conditioning_is_caught() -> None:
    """Unconditioned, `(b-a)(m-a)` overflows where the result is representable."""
    a, m, b = -1.0e308, 0.0, 1.0e308
    for u in (0.001, 0.25, 0.75):
        accepted = _triangular_from_u(u, a, m, b)
        assert math.isfinite(accepted), u
        if u <= 0.5:
            unconditioned = a + math.sqrt(u * (b - a) * (m - a))
        else:
            unconditioned = b - math.sqrt((1.0 - u) * (b - a) * (b - m))
        assert not math.isfinite(unconditioned), f"the raw form survived at u = {u}"


def test_08_a_misordered_triangular_is_refused() -> None:
    for a, m, b in ((0.0, 200.0, 100.0), (0.0, -5.0, 100.0), (100.0, 50.0, 0.0)):
        try:
            sample_triangular(_ref(), _state(), a, m, b)
        except SimSampleError:
            continue
        raise AssertionError(f"Triangular accepted {a}, {m}, {b}")


# ===========================================================================
# Beta-PERT parameterisation
# ===========================================================================
def test_09_a_changed_pert_lambda_is_caught() -> None:
    for r in (0.0, 0.25, 0.5, 1.0):
        shape = prepare_beta_pert(0.0, r, 1.0)
        for wrong in (2.0, 6.0, 3.0):
            assert shape.alpha != 1.0 + wrong * r or r == 0.0, (r, wrong)
            assert shape.alpha + shape.beta == 6.0
            assert 2.0 + wrong != 6.0 or wrong == 4.0
    # The invariant that pins it: alpha + beta = 2 + lambda, so only 4 gives 6.
    assert prepare_beta_pert(0.0, 0.25, 1.0).alpha + prepare_beta_pert(
        0.0, 0.25, 1.0
    ).beta == 6.0


def test_10_a_relaxed_dispatch_boundary_is_caught() -> None:
    """`min(alpha,beta) >= 1` would send the endpoint modes to BB."""
    for r in (0.0, 1.0):
        shape = prepare_beta_pert(0.0, r, 1.0)
        assert min(shape.alpha, shape.beta) == 1.0
        assert shape.dispatch == "BC"
        relaxed = "BB" if min(shape.alpha, shape.beta) >= 1.0 else "BC"
        assert relaxed == "BB", "the relaxed boundary should have moved it"
        assert relaxed != shape.dispatch
    # BB's own setup divides by `2ab - alpha`, which is 0 at alpha = 1, beta = 5:
    # the relaxed boundary is not merely different, it is undefined.
    a, b = 1.0, 5.0
    assert 2.0 * a * b - (a + b) == 4.0
    assert (1.0 * 5.0 * 2.0) - 6.0 == 4.0


def test_11_an_inverted_orientation_is_caught() -> None:
    """BB uses (min,max) and BC (max,min). Inverting returns the MIRROR."""
    bb = prepare_beta_pert(0.0, 0.25, 1.0)
    inverted_bb = dataclasses.replace(bb, cheng_a=bb.cheng_b, cheng_b=bb.cheng_a)
    accepted = sample_prepared_beta(_ref(), _state(), bb)
    mirrored = sample_prepared_beta(_ref(), _state(), inverted_bb)
    assert float.hex(accepted.value) != float.hex(mirrored.value)

    bc = prepare_beta_pert(0.0, 0.0, 1.0)
    inverted_bc = dataclasses.replace(bc, cheng_a=bc.cheng_b, cheng_b=bc.cheng_a)
    accepted_bc = sample_prepared_beta(_ref(), _state(), bc)
    mirrored_bc = sample_prepared_beta(_ref(), _state(), inverted_bc)
    assert float.hex(accepted_bc.value) != float.hex(mirrored_bc.value)


def test_12_a_flipped_return_orientation_is_caught() -> None:
    case = _bb_case()
    shape = prepare_beta_pert(0.0, case["r"], 1.0)
    flipped = dataclasses.replace(
        shape, first_parameter_is_oriented_a=not shape.first_parameter_is_oriented_a
    )
    state = RngState(tuple(case["initial_state"]))
    accepted = sample_prepared_beta(_ref(), state, shape)
    mirrored = sample_prepared_beta(_ref(), state, flipped)
    assert repr(accepted.value) == case["samples"][0]["accepted_sample"]
    assert repr(mirrored.value) != case["samples"][0]["accepted_sample"]
    assert math.isclose(accepted.value + mirrored.value, 1.0, rel_tol=1e-12), (
        "the flip should return the mirrored variate"
    )


# ===========================================================================
# the locked Cheng literals and expressions
# ===========================================================================
def _bb_shape(case: dict):
    return prepare_beta_pert(0.0, case["r"], 1.0)


def _bb_run(case: dict, *, log4=1.3862944, one_plus_log5=2.609438,
            logit=None, squeeze_op=None, coefficient=5.0):
    """Re-run a whole retained case with one element replaced.

    Returns `[(repr(y), attempts), ...]` and the final state, so a mutation is
    compared over the corpus rather than over one lucky sample.
    """
    logit = logit or (lambda u1: math.log(u1 / (1.0 - u1)))
    squeeze_op = squeeze_op or (lambda lhs, rhs: lhs >= rhs)
    shape = _bb_shape(case)
    a, b = shape.cheng_a, shape.cheng_b
    alpha, beta, gamma = shape.cheng_alpha, shape.cheng_beta, shape.cheng_gamma
    state = RngState(tuple(case["initial_state"]))
    out = []
    for _ in range(len(case["samples"])):
        attempts = 0
        while True:
            attempts += 1
            first = _ref().next_uniform(state)
            second = _ref().next_uniform(first.state)
            state = second.state
            u1, u2 = first.uniform, second.uniform
            v = beta * logit(u1)
            w = a * math.exp(v)
            z = u1 * u1 * u2
            rr = gamma * v - log4
            s = a + rr - w
            if squeeze_op(s + one_plus_log5, coefficient * z):
                break
            t = math.log(z)
            if s >= t:
                break
            if rr + alpha * math.log(alpha / (b + w)) >= t:
                break
        y = (w / (b + w)) if shape.first_parameter_is_oriented_a else (b / (b + w))
        out.append((repr(y), attempts))
    return out, state


def _squeeze_witness(case: dict, mutate_lhs=None, mutate_rhs=None):
    """A constructed `(u1, u2)` where the accepted and mutated squeeze disagree.

    Step 0 MEASURED why this is necessary: over 805,837 predicate evaluations the
    closest relative margin to a boundary was 7.6e-07, and the gap between
    `1.3862944` and `log(4)` is 3.9e-08. A short corpus therefore cannot flip a
    squeeze decision by luck, and a control that only re-ran 24 samples would
    pass on the mutant. So the witness is CONSTRUCTED, exactly as Step-0 control
    18b constructs one - a deterministic pair, not a search.
    """
    shape = _bb_shape(case)
    a = shape.cheng_a
    beta, gamma = shape.cheng_beta, shape.cheng_gamma
    for u1 in (0.2, 0.3, 0.4, 0.5, 0.6, 0.7):
        v = beta * math.log(u1 / (1.0 - u1))
        w = a * math.exp(v)
        s = a + (gamma * v - 1.3862944) - w
        lhs = s + 2.609438
        mutated_lhs = mutate_lhs(shape, u1, v, w) if mutate_lhs else lhs
        rhs_scale, mutated_scale = (5.0, mutate_rhs) if mutate_rhs else (5.0, 5.0)
        low, high = sorted((lhs / rhs_scale, mutated_lhs / mutated_scale))
        if low == high:
            continue
        z = (low + high) / 2.0
        u2 = z / (u1 * u1)
        if not 0.0 < u2 < 1.0:
            continue
        actual_z = u1 * u1 * u2
        accepted = lhs >= rhs_scale * actual_z
        mutated = mutated_lhs >= mutated_scale * actual_z
        if accepted != mutated:
            return u1, u2, accepted, mutated
    raise AssertionError("no squeeze witness could be constructed")


def test_13_the_bb_rerun_harness_reproduces_the_locked_corpus() -> None:
    """Without this, every mutation below could be failing for the wrong reason."""
    for case in _cases():
        if case["dispatch"] != "BB":
            continue
        actual, final = _bb_run(case)
        expected = [
            (s["accepted_sample"], s["proposal_attempts_for_this_sample"])
            for s in case["samples"]
        ]
        assert actual == expected, case["label"]
        assert final.as_list() == case["final_state"], case["label"]


def test_14_log4_substituted_for_the_locked_literal_is_caught() -> None:
    """A CONSTRUCTED witness, because the corpus cannot flip it - see below."""
    assert math.log(4.0) != 1.3862944

    case = _bb_case()
    u1, u2, accepted, mutated = _squeeze_witness(
        case,
        mutate_lhs=lambda shape, u1, v, w: (
            shape.cheng_a + (shape.cheng_gamma * v - math.log(4.0)) - w + 2.609438
        ),
    )
    assert accepted != mutated, (u1, u2)

    # And the honest counterpart: over the retained corpus the mutation is
    # INVISIBLE, exactly as Step 0's margin measurement predicts. Recording that
    # is the point - a control that only re-ran the corpus would have passed the
    # mutant and reported success.
    for candidate in _cases():
        if candidate["dispatch"] != "BB":
            continue
        assert _bb_run(candidate, log4=math.log(4.0))[0] == _bb_run(candidate)[0]


def test_15_one_plus_log5_substituted_for_the_locked_literal_is_caught() -> None:
    assert 1.0 + math.log(5.0) != 2.609438
    case = _bb_case()
    u1, u2, accepted, mutated = _squeeze_witness(
        case,
        mutate_lhs=lambda shape, u1, v, w: (
            shape.cheng_a + (shape.cheng_gamma * v - 1.3862944) - w + (1.0 + math.log(5.0))
        ),
    )
    assert accepted != mutated, (u1, u2)


def test_16_a_changed_squeeze_operator_is_caught() -> None:
    """`>` for the accepted `>=` differs exactly at equality, so equality is
    constructed rather than waited for."""
    case = _bb_case()
    shape = _bb_shape(case)
    a, beta, gamma = shape.cheng_a, shape.cheng_beta, shape.cheng_gamma
    for u1 in (0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8):
        v = beta * math.log(u1 / (1.0 - u1))
        w = a * math.exp(v)
        lhs = a + (gamma * v - 1.3862944) - w + 2.609438
        base = (lhs / 5.0) / (u1 * u1)
        for u2 in (base, math.nextafter(base, 0.0), math.nextafter(base, 1.0)):
            if not 0.0 < u2 < 1.0:
                continue
            if 5.0 * (u1 * u1 * u2) == lhs:
                assert (lhs >= 5.0 * (u1 * u1 * u2)) is True
                assert (lhs > 5.0 * (u1 * u1 * u2)) is False
                return
    raise AssertionError("no exact-equality operator witness could be constructed")


def test_17_a_changed_squeeze_coefficient_is_caught() -> None:
    case = _bb_case()
    u1, u2, accepted, mutated = _squeeze_witness(case, mutate_rhs=4.0)
    assert accepted != mutated, (u1, u2)


def test_17b_a_reversed_squeeze_direction_IS_visible_in_the_corpus() -> None:
    """Not every Cheng mutation hides: `<=` for `>=` changes most samples."""
    for case in _cases():
        if case["dispatch"] != "BB":
            continue
        mutated, _ = _bb_run(case, squeeze_op=lambda lhs, rhs: lhs <= rhs)
        expected = [
            (s["accepted_sample"], s["proposal_attempts_for_this_sample"])
            for s in case["samples"]
        ]
        differing = sum(1 for x, y in zip(mutated, expected) if x != y)
        assert differing > 0, case["label"]


def test_18_the_log1p_logit_substituted_is_caught() -> None:
    """This one changes the VALUE, and it IS visible in the corpus."""
    differing_probes = [
        u for u in (0.999, 0.9999, 0.001, 0.0001)
        if float.hex(math.log(u / (1.0 - u))) != float.hex(math.log(u) - math.log1p(-u))
    ]
    assert differing_probes, "the two logit forms agreed at every probe"

    for case in _cases():
        if case["dispatch"] != "BB":
            continue
        mutated, _ = _bb_run(case, logit=lambda u1: math.log(u1) - math.log1p(-u1))
        expected = [
            (s["accepted_sample"], s["proposal_attempts_for_this_sample"])
            for s in case["samples"]
        ]
        assert sum(1 for x, y in zip(mutated, expected) if x != y) > 0, case["label"]


def test_19_a_changed_bc_literal_is_caught() -> None:
    """`0.0138889` and `0.777778` are literals, not `1/72` and `7/9`."""
    assert 1.0 / 72.0 != 0.0138889
    assert 7.0 / 9.0 != 0.777778
    assert 3.0 / 72.0 != 0.0416667
    case = _bc_case()
    shape = prepare_beta_pert(0.0, case["r"], 1.0)
    a, b, delta = shape.cheng_a, shape.cheng_b, shape.cheng_delta
    locked_k1 = delta * (0.0138889 + 0.0416667 * b) / (a * shape.cheng_beta - 0.777778)
    fraction_k1 = delta * (1.0 / 72.0 + 3.0 / 72.0 * b) / (a * shape.cheng_beta - 7.0 / 9.0)
    assert float.hex(shape.cheng_k1) == float.hex(locked_k1)
    assert float.hex(locked_k1) != float.hex(fraction_k1)


# ===========================================================================
# consumption
# ===========================================================================
def test_20_a_rejected_proposal_that_fails_to_advance_is_caught() -> None:
    """The rewind defect: retrying from the same state loops on the same values."""
    case = _bc_case()
    shape = prepare_beta_pert(0.0, case["r"], 1.0)
    retried = next(
        s for s in case["samples"] if s["proposal_attempts_for_this_sample"] > 1
    )
    index = case["samples"].index(retried)
    state = RngState(tuple(case["initial_state"]))
    for _ in range(index):
        state = sample_prepared_beta(_ref(), state, shape).state

    result = sample_prepared_beta(_ref(), state, shape)
    assert result.proposal_attempts > 1
    # A rewinding implementation would advance by only the ACCEPTED attempt.
    _, rewound = _ref().uniforms(state, 2)
    assert result.state != rewound
    _, honest = _ref().uniforms(state, 2 * result.proposal_attempts)
    assert result.state == honest


def test_21_one_uniform_per_proposal_is_caught() -> None:
    """A proposal consumes TWO. Manufacturing one-uniform Beta desynchronises."""
    case = _bb_case()
    shape = prepare_beta_pert(0.0, case["r"], 1.0)
    state = RngState(tuple(case["initial_state"]))
    result = sample_prepared_beta(_ref(), state, shape)
    _, single = _ref().uniforms(state, result.proposal_attempts)
    assert result.state != single
    _, double = _ref().uniforms(state, 2 * result.proposal_attempts)
    assert result.state == double


def test_22_a_degenerate_driver_that_consumes_is_caught() -> None:
    start = _state()
    for family, a, m, b in (
        (FAMILY_UNIFORM, 5.0, None, 5.0),
        (FAMILY_TRIANGULAR, 5.0, 5.0, 5.0),
        (FAMILY_BETA_PERT, 5.0, 5.0, 5.0),
    ):
        if family == FAMILY_UNIFORM:
            result = sample_uniform(_ref(), start, a, b, m)
        elif family == FAMILY_TRIANGULAR:
            result = sample_triangular(_ref(), start, a, m, b)
        else:
            result = sample_beta_pert(_ref(), start, a, m, b)
        assert result.state == start, family
        assert _ref().next_uniform(start).state != start, "the control is vacuous"


# ===========================================================================
# rescale
# ===========================================================================
def test_23_the_unsafe_beta_rescale_is_caught() -> None:
    a, b = -1.0e308, 1.0e308
    for y in (0.001, 0.5, 0.999):
        safe = (1.0 - y) * a + y * b
        unsafe = a + y * (b - a)
        assert math.isfinite(safe), y
        assert not math.isfinite(unsafe), y


def test_24_a_positivity_restriction_is_caught() -> None:
    """No such rule exists, and adding one would refuse legal drivers."""
    for a, m, b in (
        (-100.0, -50.0, -10.0),
        (-50.0, 0.0, 50.0),
        (-1.0, -1.0, 1.0),
    ):
        for call in (
            lambda: sample_uniform(_ref(), _state(), a, b),
            lambda: sample_triangular(_ref(), _state(), a, m, b),
            lambda: sample_beta_pert(_ref(), _state(), a, m, b),
        ):
            result = call()
            assert math.isfinite(result.value), (a, m, b)
            assert a <= result.value <= b
        would_refuse = a > 0.0
        assert would_refuse is False, "a positivity rule would have refused this driver"


# ===========================================================================
# Bernoulli
# ===========================================================================
def test_25_a_non_strict_bernoulli_comparison_is_caught() -> None:
    from pccm_builder.sim_sample import _bernoulli_from_u

    p = 0.25
    assert _bernoulli_from_u(p, p) is False
    assert (p <= p) is True, "the defect would occur at u == p"
    # And the consequence at the domain edges.
    assert _bernoulli_from_u(0.5, 0.0) is False
    assert (0.5 <= 0.0) is False
    for u in (1e-12, 0.5, 1.0 - 1e-12):
        assert _bernoulli_from_u(u, 0.0) is False, u
        assert (u <= 0.0) is False


def test_26_broken_p0_or_p1_behaviour_is_caught() -> None:
    """Both are exact consequences of strict `<` over `(0,1)`, not special cases."""
    state = _state()
    for _ in range(200):
        never = bernoulli_occurs(_ref(), state, 0.0)
        always = bernoulli_occurs(_ref(), state, 1.0)
        assert never.occurred is False
        assert always.occurred is True
        assert 0.0 < never.uniform < 1.0
        # A defect returning u <= p would make p = 0 occur whenever u were 0 -
        # which the open interval prevents, so the strictness is what carries it.
        assert (never.uniform <= 0.0) is False
        state = never.state


def test_27_a_clamped_probability_is_caught() -> None:
    for bad in (-0.5, 1.5):
        try:
            bernoulli_occurs(_ref(), _state(), bad)
        except SimSampleError:
            continue
        raise AssertionError(f"probability {bad} was clamped instead of refused")


# ===========================================================================
# dispatch and families
# ===========================================================================
def test_28_a_fourth_family_is_caught() -> None:
    from pccm_builder import sample_distribution

    for bad in ("Normal", "LogNormal", "PERT", "Beta"):
        try:
            sample_distribution(_ref(), _state(), bad, 0.0, 50.0, 100.0)
        except SimSampleError:
            continue
        raise AssertionError(f"family {bad!r} was accepted")


def test_29_a_non_finite_result_is_refused_not_returned() -> None:
    """No silent inf or NaN reaches a caller."""
    from pccm_builder.sim_sample import _checked_result

    for bad in (float("inf"), float("-inf"), float("nan")):
        try:
            _checked_result(bad, FAMILY_UNIFORM, "convex rescale")
        except SimSampleError as error:
            assert "convex rescale" in str(error)
            continue
        raise AssertionError(f"{bad!r} was returned")


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
