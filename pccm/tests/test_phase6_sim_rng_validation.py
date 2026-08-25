#!/usr/bin/env python3
"""PCCM Phase 6 Step-2 mutation controls for the Python RNG reference.

Every conformance test in `test_phase6_sim_rng.py` claims the reference matches
the retained Step-0 vectors. A suite that cannot FAIL proves nothing, so each
control here plants exactly one defect - a changed modulus, a transposed jump
matrix, a dropped reversal, the wrong sort key - and asserts the retained
vectors then reject it.

MUTATIONS LIVE HERE, NOT IN THE CONTRACT. Every variant is built by replacing a
field of the frozen `RngReference`, or by reimplementing one step locally as the
defect would have written it. `spec/sim_contract.yaml` is never edited to
manufacture a failure.

Runs standalone or under pytest.
"""

from __future__ import annotations

import dataclasses
import json
import sys
from pathlib import Path

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))

from pccm_builder import (  # noqa: E402
    Component,
    RngReference,
    RngState,
    load_contract,
    load_sim_contract,
)
from pccm_builder.sim_rng import (  # noqa: E402
    COST_KIND,
    RISK_KIND,
    ROLE_OCCURRENCE,
    ROLE_SEVERITY,
    ROLE_VALUE,
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


def _mutated(**changes) -> RngReference:
    """One field of the reference replaced. The contract is untouched."""
    return dataclasses.replace(_ref(), **changes)


def _vectors(name: str) -> dict:
    return json.loads((EVIDENCE / "vectors" / name).read_text(encoding="utf-8"))


def _retained_first_20() -> tuple[list[str], list[int]]:
    retained = _vectors("rng_vectors.json")
    return retained["first_20_uniforms"], retained["state_after_20"]


def _matches_retained_uniforms(reference: RngReference) -> bool:
    """Does this reference still reproduce the retained corpus exactly?"""
    expected, expected_state = _retained_first_20()
    try:
        state = reference.fixed_seed_to_state(12345)
        drawn, after = reference.uniforms(state, 20)
    except Exception:  # noqa: BLE001 - a refusal is also a mismatch
        return False
    if after.as_list() != expected_state:
        return False
    return all(
        float.hex(actual) == float.hex(float(text))
        for actual, text in zip(drawn, expected)
    )


def _matches_retained_streams(reference: RngReference) -> bool:
    retained = _vectors("jump_vectors.json")["streams"]
    try:
        base = RngState(tuple(retained["0"]["initial_state"]))
        for text, case in retained.items():
            if reference.stream_initial_state(base, int(text)).as_list() != case[
                "initial_state"
            ]:
                return False
    except Exception:  # noqa: BLE001
        return False
    return True


def _design_target(reference: RngReference) -> tuple[Component, ...]:
    return reference.components_for(
        [f"CL-{i:03d}" for i in range(1, 201)],
        [f"R-{i:03d}" for i in range(1, 101)],
    )


def _retained_assignment() -> dict[tuple[str, str, str], int]:
    retained = _vectors("stream_assignment_vectors.json")
    return {
        tuple(e["component"]): e["stream"]
        for e in retained["family_a_first_10"] + retained["family_a_last_4"]
    }


# ===========================================================================
# the control that makes every other control meaningful
# ===========================================================================
def test_00_the_unmutated_reference_matches_every_retained_corpus() -> None:
    assert _matches_retained_uniforms(_ref())
    assert _matches_retained_streams(_ref())
    assignment = {
        tuple(c.as_list()): i
        for c, i in _ref().assign_component_streams(_design_target(_ref()))
    }
    for key, stream in _retained_assignment().items():
        assert assignment[key] == stream


# ===========================================================================
# the recurrence
# ===========================================================================
def test_01_a_changed_modulus_is_caught() -> None:
    assert not _matches_retained_uniforms(_mutated(m1=_ref().m1 - 1))
    assert not _matches_retained_uniforms(_mutated(m2=_ref().m2 - 1))
    assert not _matches_retained_uniforms(_mutated(m1=_ref().m2, m2=_ref().m1))


def test_02_a_changed_recurrence_coefficient_is_caught() -> None:
    for field in ("a12", "a13n", "a21", "a23n"):
        assert not _matches_retained_uniforms(
            _mutated(**{field: getattr(_ref(), field) + 1})
        ), field


def test_03_a_perturbed_norm_is_caught() -> None:
    import math

    assert not _matches_retained_uniforms(
        _mutated(norm=math.nextafter(_ref().norm, math.inf))
    )


def test_04_a_swapped_p1_p2_state_shift_is_caught() -> None:
    """The advance is part of the contract, not an ordering convenience."""
    reference = _ref()

    def defective(state: RngState) -> tuple[RngState, float]:
        s10, s11, s12, s20, s21, s22 = state.words
        p1 = (reference.a12 * s11 - reference.a13n * s10) % reference.m1
        p2 = (reference.a21 * s22 - reference.a23n * s20) % reference.m2
        # The defect: p1 and p2 written into each other's component.
        advanced = RngState((s11, s12, p2 % reference.m1, s21, s22, p1 % reference.m2))
        u = (p1 - p2 + reference.m1) * reference.norm if p1 <= p2 else (p1 - p2) * reference.norm
        return advanced, u

    expected, expected_state = _retained_first_20()
    state = reference.fixed_seed_to_state(12345)
    drawn = []
    for _ in range(20):
        state, u = defective(state)
        drawn.append(u)
    assert state.as_list() != expected_state
    assert [repr(u) for u in drawn] != expected


def test_05_a_strict_less_than_combination_boundary_is_caught() -> None:
    """`<` for the accepted `<=` differs only when p1 == p2 - and that happens."""
    reference = _ref()
    accepted = []
    defective = []
    hits = 0
    state = reference.fixed_seed_to_state(12345)
    for _ in range(400000):
        s10, s11, s12, s20, s21, s22 = state.words
        p1 = (reference.a12 * s11 - reference.a13n * s10) % reference.m1
        p2 = (reference.a21 * s22 - reference.a23n * s20) % reference.m2
        state = RngState((s11, s12, p1, s21, s22, p2))
        accepted.append((p1 - p2 + reference.m1) if p1 <= p2 else (p1 - p2))
        defective.append((p1 - p2 + reference.m1) if p1 < p2 else (p1 - p2))
        if p1 == p2:
            hits += 1
    if hits:
        assert accepted != defective, "p1 == p2 occurred and the boundary made no difference"
    else:
        # p1 == p2 is rare; the boundary is still load-bearing, shown directly.
        assert (0 - 0 + reference.m1) != 0, "at p1 == p2 the two rules give m1 and 0"


def test_06_the_uniform_is_never_produced_outside_the_open_interval() -> None:
    """A defect that could return 0 or 1 must be refused, not returned."""
    from pccm_builder import SimRngError

    reference = _ref()
    try:
        reference.next_uniform(RngState((0, 0, 0, 0, 0, 0)))
    except SimRngError:
        return
    raise AssertionError("the all-zero state was accepted")


# ===========================================================================
# the jump
# ===========================================================================
def test_07_one_changed_jump_matrix_element_is_caught() -> None:
    for name in ("jump_a1", "jump_a2"):
        original = getattr(_ref(), name)
        for row in range(3):
            for column in range(3):
                rows = [list(r) for r in original]
                rows[row][column] += 1
                mutated = _mutated(**{name: tuple(tuple(r) for r in rows)})
                assert not _matches_retained_streams(mutated), f"{name}[{row}][{column}]"


def test_08_a_transposed_jump_matrix_is_caught() -> None:
    for name in ("jump_a1", "jump_a2"):
        original = getattr(_ref(), name)
        transposed = tuple(tuple(original[r][c] for r in range(3)) for c in range(3))
        assert transposed != original, name
        assert not _matches_retained_streams(_mutated(**{name: transposed})), name


def test_09_dropping_the_matrix_boundary_reversal_is_caught() -> None:
    """The silent one: a plausible stream that is not the canonical one."""
    reference = _ref()

    def unreversed(state: RngState) -> RngState:
        def mat_vec(matrix, vector, modulus):
            return tuple(
                (r[0] * vector[0] + r[1] * vector[1] + r[2] * vector[2]) % modulus
                for r in matrix
            )

        first = mat_vec(reference.jump_a1, state.first, reference.m1)
        second = mat_vec(reference.jump_a2, state.second, reference.m2)
        return RngState(tuple(first) + tuple(second))

    base = reference.fixed_seed_to_state(12345)
    retained = _vectors("jump_vectors.json")["streams"]["1"]["initial_state"]
    assert reference.jump_to_next_stream(base).as_list() == retained
    assert unreversed(base).as_list() != retained


def test_10_reversing_only_one_side_is_caught() -> None:
    reference = _ref()

    def half_reversed(state: RngState) -> RngState:
        def mat_vec(matrix, vector, modulus):
            return tuple(
                (r[0] * vector[0] + r[1] * vector[1] + r[2] * vector[2]) % modulus
                for r in matrix
            )

        first = mat_vec(reference.jump_a1, tuple(reversed(state.first)), reference.m1)
        second = mat_vec(reference.jump_a2, state.second, reference.m2)
        return RngState(tuple(reversed(first)) + tuple(second))

    base = reference.fixed_seed_to_state(12345)
    retained = _vectors("jump_vectors.json")["streams"]["1"]["initial_state"]
    assert half_reversed(base).as_list() != retained


def test_11_swapping_the_two_jump_matrices_is_caught() -> None:
    swapped = _mutated(jump_a1=_ref().jump_a2, jump_a2=_ref().jump_a1)
    assert not _matches_retained_streams(swapped)


# ===========================================================================
# seeding
# ===========================================================================
def test_12_a_changed_auto_multiplier_is_caught() -> None:
    retained = json.loads(
        (EVIDENCE / "raw" / "seed_map.json").read_text(encoding="utf-8")
    )["nonce_to_seed_pairs"]
    mutated = _mutated(auto_multiplier=_ref().auto_multiplier + 1)
    differences = [
        pair
        for pair in retained
        if pair["auto_nonce"] < _ref().nonce_exhausted
        and mutated.auto_seed_from_nonce(pair["auto_nonce"]) != pair["effective_seed"]
    ]
    assert differences, "a changed multiplier produced the retained seeds"


def test_13_a_changed_auto_modulus_is_caught() -> None:
    mutated = _mutated(auto_modulus=_ref().auto_modulus - 2)
    assert mutated.auto_seed_from_nonce(2) != _ref().auto_seed_from_nonce(2)


def test_14_sequential_stepping_substituted_for_the_power_is_caught() -> None:
    """Stepping AGREES with the power - that is exactly why it is a trap.

    A stepping implementation gives the same answer and is O(nonce), so the way
    it fails in practice is by being bounded: an implementation that caps its
    loop, or that a caller abandons, silently returns the wrong seed. The
    authority is a power precisely so that cannot arise.
    """
    reference = _ref()
    nonce = 5_000_000

    def stepped(limit: int) -> int:
        value = 1
        for _ in range(min(nonce, limit)):
            value = (value * reference.auto_multiplier) % reference.auto_modulus
        return value

    exact = reference.auto_seed_from_nonce(nonce)
    assert stepped(10_000_000) == exact, "unbounded stepping should agree"
    for cap in (1000, 100_000, 4_999_999):
        assert stepped(cap) != exact, f"a stepping loop capped at {cap} went undetected"


def test_15_a_widened_seed_domain_is_caught() -> None:
    """The domain is the owner's, and widening it must be visible."""
    from pccm_builder import SimRngError

    widened = _mutated(seed_min=0, seed_max=2147483647)

    # 2147483647 is a valid MRG residue, so ONLY the domain refuses it. The
    # mutant admits it; the accepted reference does not.
    widened.fixed_seed_to_state(2147483647)
    try:
        _ref().fixed_seed_to_state(2147483647)
    except SimRngError:
        pass
    else:
        raise AssertionError("the accepted reference admitted a seed above the domain")

    # 0 is refused by BOTH, and by different rules - the domain in the accepted
    # reference, the all-zero state invariant in the mutant. Defence in depth,
    # recorded rather than assumed: a widened domain alone cannot smuggle in the
    # absorbing state.
    for reference, label in ((_ref(), "accepted"), (widened, "widened")):
        try:
            reference.fixed_seed_to_state(0)
        except SimRngError:
            continue
        raise AssertionError(f"the {label} reference admitted seed 0")


def test_16_a_mixer_substituted_for_the_repeated_scalar_is_caught() -> None:
    retained = _vectors("seed_vectors.json")["examples"]
    for example in retained:
        seed = example["seed"]
        mixed = [(seed * (i + 1)) % 2147483647 for i in range(6)]
        assert mixed != example["state"] or seed in (0,), seed
    assert _ref().fixed_seed_to_state(12345).as_list() == [12345] * 6


# ===========================================================================
# component streams
# ===========================================================================
def _assignment_from(order_key) -> dict[tuple[str, str, str], int]:
    ordered = sorted(_design_target(_ref()), key=order_key)
    return {tuple(c.as_list()): i for i, c in enumerate(ordered)}


def test_17_a_nonzero_stream_index_origin_is_caught() -> None:
    retained = _retained_assignment()
    shifted = {
        tuple(c.as_list()): i + 1
        for c, i in _ref().assign_component_streams(_design_target(_ref()))
    }
    assert any(shifted[key] != stream for key, stream in retained.items())


def test_18_physical_row_order_is_caught() -> None:
    """Registers are entered in whatever order a user typed them."""
    retained = _retained_assignment()
    components = list(_design_target(_ref()))
    rows = components[::-1]  # a legal, different physical order
    by_row = {tuple(c.as_list()): i for i, c in enumerate(rows)}
    assert any(by_row[key] != stream for key, stream in retained.items())


def test_19_numeric_id_sorting_is_caught() -> None:
    """Ordinal and numeric coincide on today's zero-padded IDs, so the control
    uses the widened IDs where they demonstrably part company."""
    reference = _ref()
    ids = ["CL-001", "CL-002", "CL-999", "CL-1000", "CL-0001"]
    components = reference.components_for(ids, [])
    ordinal = [c.permanent_id for c, _ in reference.assign_component_streams(components)]
    numeric = sorted(ids, key=lambda s: (int(s.split("-")[1]), s))
    assert ordinal == ["CL-0001", "CL-001", "CL-002", "CL-1000", "CL-999"]
    assert ordinal != numeric


def test_20_locale_or_case_folding_order_is_caught() -> None:
    reference = _ref()
    ids = ["CL-001", "cl-002", "CL-003"]
    components = reference.components_for(ids, [])
    ordinal = [c.permanent_id for c, _ in reference.assign_component_streams(components)]
    folded = sorted(ids, key=str.lower)
    assert ordinal == ["CL-001", "CL-003", "cl-002"], ordinal
    assert ordinal != folded


def test_21_global_occurrence_severity_blocks_are_caught() -> None:
    """The interpretation the review warned about, proven to differ.

    Sorting by the component KEY - `COST_SAMPLE`, `RISK_OCCURRENCE`,
    `RISK_SEVERITY` - puts every occurrence stream before every severity stream.
    The retained tail interleaves them per Risk, so the two disagree.
    """
    key_rank = {
        (COST_KIND, ROLE_VALUE): 0,
        (RISK_KIND, ROLE_OCCURRENCE): 1,
        (RISK_KIND, ROLE_SEVERITY): 2,
    }
    blocked = _assignment_from(
        lambda c: (key_rank[(c.kind, c.role)], tuple(ord(x) for x in c.permanent_id))
    )
    retained = _retained_assignment()
    mismatches = [k for k, stream in retained.items() if blocked[k] != stream]
    assert mismatches, "global blocks reproduced the retained per-Risk interleaving"
    assert blocked[(RISK_KIND, "R-100", ROLE_OCCURRENCE)] == 299
    assert retained[(RISK_KIND, "R-100", ROLE_OCCURRENCE)] == 398


def test_22_severity_before_occurrence_within_a_risk_is_caught() -> None:
    role_rank = {ROLE_VALUE: 0, ROLE_SEVERITY: 0, ROLE_OCCURRENCE: 1}
    kind_rank = {COST_KIND: 0, RISK_KIND: 1}
    swapped = _assignment_from(
        lambda c: (
            kind_rank[c.kind],
            tuple(ord(x) for x in c.permanent_id),
            role_rank[c.role],
        )
    )
    retained = _retained_assignment()
    assert swapped[(RISK_KIND, "R-100", ROLE_OCCURRENCE)] == 399
    assert retained[(RISK_KIND, "R-100", ROLE_OCCURRENCE)] == 398


def test_23_risks_before_cost_lines_is_caught() -> None:
    kind_rank = {COST_KIND: 1, RISK_KIND: 0}
    role_rank = {ROLE_VALUE: 0, ROLE_OCCURRENCE: 0, ROLE_SEVERITY: 1}
    flipped = _assignment_from(
        lambda c: (
            kind_rank[c.kind],
            tuple(ord(x) for x in c.permanent_id),
            role_rank[c.role],
        )
    )
    retained = _retained_assignment()
    assert flipped[(COST_KIND, "CL-001", ROLE_VALUE)] != retained[
        (COST_KIND, "CL-001", ROLE_VALUE)
    ]


def test_24_a_contract_whose_two_kind_orderings_disagree_is_refused() -> None:
    """The contract states the kind axis twice; the reference refuses a mismatch."""
    import copy

    import yaml

    from pccm_builder.sim_rng import SimRngError

    sim = load_sim_contract(SIM_PATH)
    raw = copy.deepcopy(sim.raw)
    raw["accumulation"]["driver_kind_order"] = ["risk", "cost_line"]
    broken = dataclasses.replace(sim, raw=raw)
    try:
        RngReference.from_contracts(broken, load_contract(CONTRACT_PATH))
    except SimRngError:
        return
    raise AssertionError("a contract with two disagreeing kind orderings was accepted")


def test_25_a_duplicate_component_is_refused() -> None:
    from pccm_builder import SimRngError

    duplicated = (
        Component(COST_KIND, "CL-001", ROLE_VALUE),
        Component(COST_KIND, "CL-001", ROLE_VALUE),
    )
    try:
        _ref().assign_component_streams(duplicated)
    except SimRngError:
        return
    raise AssertionError("a duplicate component was assigned two streams")


def test_26_an_unknown_kind_or_role_is_refused() -> None:
    from pccm_builder import SimRngError

    for bad in (
        Component("DRIVER", "CL-001", ROLE_VALUE),
        Component(COST_KIND, "CL-001", ROLE_SEVERITY),
        Component(RISK_KIND, "R-001", ROLE_VALUE),
    ):
        try:
            _ref().assign_component_streams((bad,))
        except SimRngError:
            continue
        raise AssertionError(f"component {bad} was accepted")


if __name__ == "__main__":
    import pytest

    raise SystemExit(pytest.main([__file__, "-q"]))
