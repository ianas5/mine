#!/usr/bin/env python3
"""The Phase-6 conformance corpus: `build/phase6_cases.json`.

TEST DATA, not a runtime contract. Later VBA implementation steps consume this
file to check themselves against the accepted Python oracle. It is not a second
simulation engine, not a performance benchmark, not a stochastic recommendation
and not workbook data.

--------------------------------------------------------------------------------
NO SECOND IMPLEMENTATION LIVES HERE
--------------------------------------------------------------------------------
Every expected value is produced by calling the ACCEPTED public reference:
`RngReference` for states, jumps and seeds, the Step-3 samplers for draws,
`prepare_simulation` / `run_simulation` for whole runs, `result_digest` for the
digest, `describe` for statistics and `contingency_at` for the reporting lookup.
There is no RNG, no sampler, no accumulation and no statistic of this module's
own - a duplicate would be free to drift from the oracle it is supposed to be
checking.

It also imports nothing from `tests/`. Tests are consumers of authority, not
sources of it, so the fixtures are built here from the accepted public
`CalculationModel` vocabulary and materialised through the accepted Phase-5
`to_model` adapter.

--------------------------------------------------------------------------------
THE COMPARISON POLICY IS THE POINT
--------------------------------------------------------------------------------
The accepted plan (section 15) is explicit that cross-language identity is NOT
uniform across the engine, and this corpus must not quietly strengthen it. Every
case therefore carries an explicit policy, and a case may additionally carry an
`expected_exact` block for the fields that are exact WHATEVER the case policy is
- a draw count and a post-sample RNG state are exact even where the transformed
sample they produced is only tolerance-bounded.

The numeric tolerance itself is deliberately absent. `sim_contract.yaml` forbids
a comparison tolerance outright, so a number invented here would be a new
authority; the class is stated and the value is left to the evidence policy that
owns it.
"""

from __future__ import annotations

import math
from typing import Any

from .calc_cases import reference_stream, to_model, tolerances_from
from .calc_fingerprint import (
    canonical_number,
    encode_section,
    fingerprint,
    integer_field,
    number_field,
    text_field,
    utf16_length,
)
from .calc_loader import CalcContract
from .calc_numeric import CalculationRefusal
from .contract_loader import ContractError, InputContract
from .sim_loader import SimContract
from .sim_oracle import (
    SimOracleError,
    contingency_at,
    deterministic_base_of,
    prepare_simulation,
    resolve_percentile_ladder,
    result_digest,
    result_digest_stream,
    run_simulation,
)
from .sim_rng import RngReference
from .sim_sample import (
    bernoulli_occurs,
    prepare_beta_pert,
    sample_distribution,
    sample_prepared_beta,
)
from .sim_stats import (
    SimStatsError,
    describe,
    percentile_type7,
    sample_mean,
    sample_standard_deviation,
)

SCHEMA_VERSION = 2
"""Bumped from 1: semantic numbers are JSON numbers, every case carries its own
version identity, and an exact case additionally carries the accepted Phase-5
canonical encoding of each of its numbers."""

CANONICAL_SUFFIX = "_canonical"
"""The sidecar key. `x` is the semantic JSON number; `x_canonical` is the same
value in the ACCEPTED Phase-5 canonical-number encoding, so binary64 identity has
an exact textual form without the number itself being stringified. No new
floating encoding is invented here - `canonical_number` is Phase-5 authority."""

# --- comparison policies ----------------------------------------------------
EXACT = "EXACT"
TOLERANCE_BOUNDED = "TOLERANCE_BOUNDED"
STATISTICAL = "STATISTICAL"
SAME_RUNTIME_ONLY = "SAME_RUNTIME_ONLY"
RUNTIME_ONLY = "RUNTIME_ONLY"

POLICIES: dict[str, str] = {
    EXACT: (
        "Bit-for-bit equality is required across implementations. Used for RNG "
        "states and uniforms, jump-ahead states, stream assignment, Bernoulli "
        "decisions, draw and consumption counts, the locked result-digest "
        "framing vectors, and individual results that are exactly representable "
        "by construction."
    ),
    TOLERANCE_BOUNDED: (
        "Agreement is required only within a bounded numeric tolerance. Used for "
        "transformed floating samples and for general whole-engine comparison. "
        "THE TOLERANCE VALUE IS NOT STATED HERE: sim_contract.yaml forbids a "
        "comparison tolerance, so a number in this corpus would be a new "
        "authority. The evidence policy that owns it supplies it."
    ),
    STATISTICAL: (
        "Sample-for-sample identity is NOT claimed. Confidence comes from "
        "distributional agreement at a stated sample size. Used where a "
        "rejection sampler can legitimately desynchronise two implementations."
    ),
    SAME_RUNTIME_ONLY: (
        "The expectation is a RELATION between two runs inside one runtime - "
        "equal or unequal - not a value another implementation must reproduce. "
        "Used for replay identity, row-order identity and seed divergence."
    ),
    RUNTIME_ONLY: (
        "The behaviour cannot be checked without workbook publication or state "
        "machinery that does not exist yet. Recorded so it is not forgotten; it "
        "carries no numeric expectation."
    ),
}

_PYTHON_REFERENCE_NOTE = (
    "Values under 'python_reference' are what the accepted Python oracle "
    "produced. They are context for a failing comparison, NEVER a "
    "cross-language expectation; the case policy governs what must match."
)


# ===========================================================================
# numbers
# ===========================================================================
def _n(value) -> float:
    """One semantic number, as a JSON number and never as text.

    Stringifying a finite float to preserve `repr` would put it beyond
    `allow_nan=False`: `repr(float("nan"))` is the ordinary string `"nan"`, which
    serialises happily. Keeping the value a Python float all the way to
    `json.dumps` makes that guard a real boundary, and this function refuses a
    non-finite value here as well so the failure names the corpus rather than the
    serialiser.
    """
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SimOracleError(f"not a semantic number: {value!r}")
    number = float(value)
    if not math.isfinite(number):
        raise SimOracleError(
            f"a non-finite value {value!r} reached the corpus; it is refused here "
            "rather than emitted"
        )
    return number


def _canonical(value: float) -> str:
    """The accepted Phase-5 canonical encoding of the same number."""
    return canonical_number(float(value))


def _with_canonical(node):
    """Add a canonical sidecar beside every float leaf, recursively.

    Applied to the blocks whose comparison is EXACT. An integer needs no sidecar
    - JSON integers are exact - and a tolerance-bounded block gets none, because
    an exact textual form there would invite exactly the comparison the accepted
    evidence model refuses.
    """
    if isinstance(node, list):
        return [_with_canonical(item) for item in node]
    if not isinstance(node, dict):
        return node
    out: dict[str, Any] = {}
    for key, value in node.items():
        if key.endswith(CANONICAL_SUFFIX):
            out[key] = value
            continue
        # A NUMERIC CONTAINER gets exactly one sidecar and is not descended
        # into: recursing would also hang a sidecar off every label inside a
        # quantile mapping, so the mapping would hold numbers and text at once.
        if isinstance(value, float):
            out[key] = value
            out[key + CANONICAL_SUFFIX] = _canonical(value)
        elif (
            isinstance(value, list)
            and value
            and all(isinstance(item, float) for item in value)
        ):
            out[key] = value
            out[key + CANONICAL_SUFFIX] = [_canonical(item) for item in value]
        elif (
            isinstance(value, dict)
            and value
            and all(isinstance(item, float) for item in value.values())
        ):
            out[key] = value
            out[key + CANONICAL_SUFFIX] = {
                label: _canonical(item) for label, item in value.items()
            }
        else:
            out[key] = _with_canonical(value)
    return out


# ===========================================================================
# fixture vocabulary
# ===========================================================================
def _cost(
    permanent_id: str,
    distribution: str,
    minimum: Any,
    most_likely: Any,
    maximum: Any,
    quantity: Any = 1.0,
    weights: tuple = (1.0,),
) -> dict[str, Any]:
    return {
        "permanent_id": permanent_id,
        "distribution": distribution,
        "currency": "SAR",
        "inflation_profile": "Standard",
        "min_value": minimum,
        "most_likely": most_likely,
        "max_value": maximum,
        "profile_weights": list(weights),
        "quantity": quantity,
    }


def _risk(
    permanent_id: str,
    distribution: str,
    minimum: Any,
    most_likely: Any,
    maximum: Any,
    probability: Any = 0.5,
    weights: tuple = (1.0,),
) -> dict[str, Any]:
    return {
        "permanent_id": permanent_id,
        "distribution": distribution,
        "currency": "SAR",
        "inflation_profile": "Standard",
        "min_value": minimum,
        "most_likely": most_likely,
        "max_value": maximum,
        "profile_weights": list(weights),
        "probability": probability,
    }


def _payload(
    cost_lines: list | None = None,
    risks: list | None = None,
    duration: int = 1,
    start_year: int = 2026,
    inflation: dict | None = None,
    discount_rate: Any = 0.10,
) -> dict[str, Any]:
    """A model fixture in the accepted Phase-5 payload vocabulary.

    One applied project year with weight 1, no inflation and FX 1 gives
    `Knom = Kpv = 1` exactly, so a contribution is exactly `sample * Quantity`
    and a later implementation can check the arithmetic it is meant to check
    rather than the escalation path Phase 5 already owns.
    """
    return {
        "timeline": {"base_year": 2026, "start_year": start_year, "duration": duration},
        "discount_rate": discount_rate,
        "fx": [{"currency": "SAR", "rate": 1}],
        "inflation": inflation if inflation is not None else {"Standard": {}},
        "cost_lines": cost_lines or [],
        "risks": risks or [],
    }


def _case(
    identifier: str,
    layer: str,
    title: str,
    comparison: str,
    inputs: dict[str, Any],
    expected: dict[str, Any] | None = None,
    expected_exact: dict[str, Any] | None = None,
    expected_refusal: dict[str, Any] | None = None,
    python_reference: dict[str, Any] | None = None,
    note: str | None = None,
) -> dict[str, Any]:
    if comparison not in POLICIES:
        raise ValueError(f"unknown comparison policy {comparison!r}")
    if expected is None and expected_exact is None and expected_refusal is None:
        raise ValueError(f"case {identifier!r} carries no expectation")
    record: dict[str, Any] = {
        "id": identifier,
        "layer": layer,
        "title": title,
        "comparison": comparison,
        "inputs": inputs,
    }
    if expected is not None:
        # `expected` is governed by the case policy, so it only earns an exact
        # sidecar when that policy IS exact.
        record["expected"] = _with_canonical(expected) if comparison == EXACT else expected
    if expected_exact is not None:
        record["expected_exact"] = _with_canonical(expected_exact)
    if expected_refusal is not None:
        record["expected_refusal"] = expected_refusal
    if python_reference is not None:
        record["python_reference"] = python_reference
    if note is not None:
        record["note"] = note
    return record


# ===========================================================================
# A - RNG and stream layer  (plan layers A and B: EXACT)
# ===========================================================================
_ACCEPTED_SEEDS = (1, 2, 12345, 2147483646)
_ACCEPTED_STREAMS = (0, 1, 7, 399, 401)


def _state_words(state) -> list[int]:
    return [int(word) for word in state.words]


def _rng_cases(reference: RngReference) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []

    for seed in _ACCEPTED_SEEDS:
        state = reference.fixed_seed_to_state(seed)
        uniforms, after = reference.uniforms(state, 5)
        cases.append(_case(
            f"rng.fixed_seed.{seed}",
            "A_rng",
            f"FIXED seed {seed}: scalar-to-state expansion and the first five uniforms",
            EXACT,
            {"seed_mode": "FIXED", "seed": seed, "draws": 5},
            expected_exact={
                "initial_state": _state_words(state),
                "first_uniforms": [_n(value) for value in uniforms],
                "state_after": _state_words(after),
            },
            note="The scalar is repeated into all six words. No mixer, no hash.",
        ))

    base = reference.fixed_seed_to_state(12345)
    walked: dict[int, Any] = {}
    current = base
    for index in range(max(_ACCEPTED_STREAMS) + 1):
        if index in _ACCEPTED_STREAMS:
            walked[index] = current
        if index < max(_ACCEPTED_STREAMS):
            current = reference.jump_to_next_stream(current)

    for index in _ACCEPTED_STREAMS:
        state = walked[index]
        uniforms, after = reference.uniforms(state, 5)
        cases.append(_case(
            f"rng.stream.{index}",
            "B_jump",
            f"Stream {index}: the base state advanced by {index} canonical 2^127 jumps",
            EXACT,
            {"seed": 12345, "stream_index": index, "draws": 5},
            expected_exact={
                "initial_state": _state_words(state),
                "first_uniforms": [_n(value) for value in uniforms],
                "state_after": _state_words(after),
            },
            note=(
                "PCCM stores state oldest-first; the jump matrices operate "
                "newest-first, so each triple is reversed at the matrix boundary."
            ),
        ))

    cases.append(_case(
        "rng.state.illegal_all_zero",
        "A_rng",
        "An all-zero component is an absorbing state and is refused",
        EXACT,
        {"state": [0, 0, 0, 1, 2, 3]},
        expected_refusal={
            "kind": "rng_state",
            "reason": "the first component is all zero, which the recurrence can never leave",
        },
    ))
    cases.append(_case(
        "rng.state.illegal_out_of_range",
        "A_rng",
        "A word at or above its modulus is not an MRG state",
        EXACT,
        {"state": [4294967087, 2, 3, 1, 2, 3]},
        expected_refusal={"kind": "rng_state", "reason": "word 0 is not below m1"},
    ))
    return cases


def _stream_assignment_cases(reference: RngReference) -> list[dict[str, Any]]:
    """The canonical assignment, and its invariance to physical register order.

    200 Cost Lines and 100 Risks give 400 components - the accepted Step-0
    family A. Only the first ten and the last four are carried: the rule is a
    total order, so the interior adds bytes without adding evidence, and the last
    four are where the interleaving of occurrence and severity per Risk is
    visible.
    """
    cost_ids = [f"CL-{index:03d}" for index in range(1, 201)]
    risk_ids = [f"R-{index:03d}" for index in range(1, 101)]

    assignment = reference.assign_component_streams(
        reference.components_for(cost_ids, risk_ids)
    )
    listed = [
        {"component": list(component.as_list()), "stream": index}
        for component, index in assignment
    ]

    shuffled_costs = list(reversed(cost_ids))
    shuffled_risks = risk_ids[50:] + risk_ids[:50]
    reordered = reference.assign_component_streams(
        reference.components_for(shuffled_costs, shuffled_risks)
    )
    reordered_listed = [
        {"component": list(component.as_list()), "stream": index}
        for component, index in reordered
    ]
    if reordered_listed != listed:  # pragma: no cover - the sort makes this impossible
        raise SimOracleError("stream assignment is not invariant to register order")

    return [
        _case(
            "stream.assignment.canonical_400",
            "B_jump",
            "Canonical stream assignment for 200 Cost Lines and 100 Risks",
            EXACT,
            {"cost_line_count": 200, "risk_count": 100},
            expected_exact={
                "total_components": len(listed),
                "first_10": listed[:10],
                "last_4": listed[-4:],
            },
            note=(
                "Cost Lines first, then each Risk interleaved occurrence-then-"
                "severity. Kind and role are separate sort keys, so the three "
                "component kinds do NOT form three global blocks."
            ),
        ),
        _case(
            "stream.assignment.row_reorder_invariant",
            "B_jump",
            "Physical register order does not reach stream assignment",
            EXACT,
            {
                "cost_line_order": "descending",
                "risk_order": "rotated_by_50",
                "compare_with": "stream.assignment.canonical_400",
            },
            expected_exact={
                "assignment_identical": True,
                "first_10": reordered_listed[:10],
                "last_4": reordered_listed[-4:],
            },
        ),
    ]


# ===========================================================================
# B - seed semantics
# ===========================================================================
_ACCEPTED_NONCES = (0, 1, 2, 3, 10, 1000, 2147483645)


def _seed_cases(reference: RngReference, inputs: InputContract) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = [
        _case(
            "seed.fixed.domain",
            "B_seed",
            "The admissible FIXED seed domain, owned by input_contract.yaml",
            EXACT,
            {"owner": "input_contract.yaml"},
            expected_exact={
                "seed_min": reference.seed_min,
                "seed_max": reference.seed_max,
            },
        )
    ]
    for nonce in _ACCEPTED_NONCES:
        cases.append(_case(
            f"seed.auto.nonce.{nonce}",
            "B_seed",
            f"AUTO nonce {nonce} maps to its effective seed by modular exponentiation",
            EXACT,
            {"seed_mode": "AUTO", "auto_nonce": nonce},
            expected_exact={
                "effective_seed": reference.auto_seed_from_nonce(nonce),
            },
            note="multiplier ^ nonce mod modulus, evaluated in O(log nonce).",
        ))

    exhausted = reference.nonce_exhausted
    cases.append(_case(
        "seed.auto.nonce.exhausted",
        "B_seed",
        f"AUTO nonce {exhausted} is exhausted and is refused, not wrapped",
        EXACT,
        {"seed_mode": "AUTO", "auto_nonce": exhausted},
        expected_refusal={
            "kind": "auto_nonce_exhausted",
            "reason": (
                f"the cycle has period {exhausted}, so allocating it would silently "
                f"reissue the seed for nonce 0"
            ),
        },
        note="Not executed as a mapping: the refusal precedes any derivation.",
    ))
    for bad, reason in ((-1, "negative"), (exhausted + 1, "beyond the cycle")):
        cases.append(_case(
            f"seed.auto.nonce.refused.{'negative' if bad < 0 else 'beyond'}",
            "B_seed",
            f"AUTO nonce {bad} is refused ({reason})",
            EXACT,
            {"seed_mode": "AUTO", "auto_nonce": bad},
            expected_refusal={"kind": "auto_nonce_domain", "reason": reason},
        ))

    for bad in (0, reference.seed_max + 1):
        cases.append(_case(
            f"seed.fixed.refused.{bad}",
            "B_seed",
            f"FIXED seed {bad} lies outside the admissible domain",
            EXACT,
            {"seed_mode": "FIXED", "seed": bad},
            expected_refusal={
                "kind": "fixed_seed_domain",
                "reason": (
                    f"outside [{reference.seed_min}, {reference.seed_max}] owned by "
                    "input_contract.yaml"
                ),
            },
        ))
    return cases


# ===========================================================================
# C - the sampler layer
# ===========================================================================
_INJECTED_UNIFORMS = (0.001, 0.25, 0.3, 0.5, 0.75, 0.999)
"""Plan layer C works on INJECTED uniforms, not generated ones, so a transform
can be checked without a generator in the way."""


def _sampler_cases(reference: RngReference) -> list[dict[str, Any]]:
    from .sim_sample import _triangular_from_u, _uniform_from_u

    cases: list[dict[str, Any]] = []

    # --- Uniform, injected --------------------------------------------------
    uniform_rows = [
        {"u": _n(u), "value": _n(_uniform_from_u(u, 0.0, 100.0))}
        for u in _INJECTED_UNIFORMS
    ]
    cases.append(_case(
        "sampler.uniform.injected",
        "C_transform",
        "Uniform transform on injected uniforms, convex form",
        TOLERANCE_BOUNDED,
        {"family": "Uniform", "a": 0.0, "b": 100.0, "injected_uniforms": list(_INJECTED_UNIFORMS)},
        expected={"rows": uniform_rows},
        expected_exact={"uniforms_per_sample": 1, "transform": "x = (1 - u) * a + u * b"},
        note=(
            "The convex form is required: a + u * (b - a) overflows at "
            "a = -MAX, b = +MAX where every convex result is finite."
        ),
    ))

    extreme = _uniform_from_u(0.25, -1.0e308, 1.0e308)
    cases.append(_case(
        "sampler.uniform.extreme_span",
        "H_domain",
        "Uniform across the widest authorised span stays finite",
        TOLERANCE_BOUNDED,
        {"family": "Uniform", "a": -1.0e308, "b": 1.0e308, "injected_uniforms": [0.25]},
        expected={"value": _n(extreme)},
        expected_exact={"finite": True, "uniforms_per_sample": 1},
    ))

    # --- Uniform degeneracy, with an ignored Most Likely ---------------------
    start = reference.fixed_seed_to_state(12345)
    for label, most_likely in (("absent", None), ("populated_and_ignored", 500.0)):
        result = sample_distribution(reference, start, "Uniform", 100.0, most_likely, 100.0)
        cases.append(_case(
            f"sampler.uniform.degenerate.{label}",
            "C_transform",
            f"Degenerate Uniform (a == b) with Most Likely {label}",
            EXACT,
            {"family": "Uniform", "a": 100.0, "most_likely": most_likely, "b": 100.0,
             "start_state": _state_words(start)},
            expected_exact={
                "value": _n(result.value),
                "uniforms_consumed": 0,
                "state_unchanged": True,
                "state_after": _state_words(result.state),
            },
            note=(
                "D1: a Uniform ignores Most Likely entirely, so a populated one "
                "cannot make the driver non-degenerate and cannot change RNG "
                "consumption."
            ),
        ))

    # --- Triangular, injected, both branches and the boundary ---------------
    triangular_rows = []
    for u in (0.001, 0.29999999999999993, 0.3, 0.30000000000000004, 0.5, 0.999):
        triangular_rows.append({"u": _n(u), "value": _n(_triangular_from_u(u, 0.0, 30.0, 100.0))})
    cases.append(_case(
        "sampler.triangular.injected_branches",
        "C_transform",
        "Triangular inverse CDF: lower branch, the branch point c, upper branch",
        TOLERANCE_BOUNDED,
        {"family": "Triangular", "a": 0.0, "m": 30.0, "b": 100.0,
         "branch_point_c": 0.3, "comparison_operator": "u <= c takes the lower branch"},
        expected={"rows": triangular_rows},
        expected_exact={"uniforms_per_sample": 1, "value_at_branch_point": _n(30.0)},
        note="At u == c exactly the result is the mode. `<` instead of `<=` moves it.",
    ))
    for label, mode in (("m_equals_a", 0.0), ("m_equals_b", 100.0)):
        cases.append(_case(
            f"sampler.triangular.{label}",
            "C_transform",
            f"Triangular with {label.replace('_', ' ')}",
            TOLERANCE_BOUNDED,
            {"family": "Triangular", "a": 0.0, "m": mode, "b": 100.0,
             "injected_uniforms": [0.5]},
            expected={"value": _n(_triangular_from_u(0.5, 0.0, mode, 100.0))},
            expected_exact={"uniforms_per_sample": 1},
        ))
    degenerate_tri = sample_distribution(reference, start, "Triangular", 7.0, 7.0, 7.0)
    cases.append(_case(
        "sampler.triangular.degenerate",
        "C_transform",
        "Degenerate Triangular (a == m == b) consumes nothing",
        EXACT,
        {"family": "Triangular", "a": 7.0, "m": 7.0, "b": 7.0,
         "start_state": _state_words(start)},
        expected_exact={
            "value": _n(degenerate_tri.value),
            "uniforms_consumed": 0,
            "state_unchanged": True,
        },
    ))

    # --- Cheng, from the accepted stream ------------------------------------
    cheng_shapes = (
        ("bb_interior", 0.25, "BB"),
        ("bb_symmetric", 0.5, "BB"),
        ("bb_near_boundary", 0.01, "BB"),
        ("bc_alpha_1", 0.0, "BC"),
        ("bc_beta_1", 1.0, "BC"),
    )
    for index, (label, ratio, dispatch) in enumerate(cheng_shapes):
        shape = prepare_beta_pert(0.0, ratio, 1.0)
        if shape.dispatch != dispatch:  # pragma: no cover - the rule is fixed
            raise SimOracleError(f"{label}: dispatch is {shape.dispatch}, expected {dispatch}")
        state = reference.fixed_seed_to_state(12345)
        for _ in range(index):
            state = reference.jump_to_next_stream(state)
        initial = state
        samples = []
        attempts = 0
        uniforms = 0
        for ordinal in range(1, 25):
            drawn = sample_prepared_beta(reference, state, shape)
            state = drawn.state
            attempts += drawn.proposal_attempts
            uniforms += drawn.uniforms_consumed
            samples.append({
                "index": ordinal,
                "value": _n(drawn.value),
                "proposal_attempts": drawn.proposal_attempts,
                "uniforms_for_this_sample": drawn.uniforms_consumed,
                "cumulative_uniforms": uniforms,
                "state_after": _state_words(state),
            })
        cases.append(_case(
            f"sampler.beta.cheng.{label}",
            "E_cheng",
            f"Cheng {dispatch}: alpha {shape.alpha}, beta {shape.beta}, 24 samples",
            TOLERANCE_BOUNDED,
            {
                "family": "Beta-PERT", "a": 0.0, "m": ratio, "b": 1.0,
                "seed": 12345, "stream_index": index, "sample_count": 24,
            },
            expected={"samples": [
                {"index": row["index"], "value": row["value"]} for row in samples
            ]},
            expected_exact={
                "dispatch": shape.dispatch,
                "alpha": _n(shape.alpha),
                "beta": _n(shape.beta),
                "total_proposal_attempts": attempts,
                "total_uniforms": uniforms,
                "uniforms_per_attempt": 2,
                "initial_state": _state_words(initial),
                "final_state": _state_words(state),
                "per_sample": [
                    {
                        "index": row["index"],
                        "proposal_attempts": row["proposal_attempts"],
                        "uniforms_for_this_sample": row["uniforms_for_this_sample"],
                        "cumulative_uniforms": row["cumulative_uniforms"],
                        "state_after": row["state_after"],
                    }
                    for row in samples
                ],
            },
            note=(
                "Draw counts and states are EXACT; the transformed values are "
                "tolerance-bounded (plan layer E). a = 0 and b = 1 make the "
                "convex rescale the identity, so the value IS the Beta variate. "
                "A rejected proposal consumes both uniforms and the retry "
                "continues from the resulting state - there is no rewind."
            ),
        ))

    degenerate_beta = sample_distribution(reference, start, "Beta-PERT", -2.0, -2.0, -2.0)
    cases.append(_case(
        "sampler.beta.degenerate",
        "E_cheng",
        "Degenerate Beta-PERT forms no shape ratio and consumes nothing",
        EXACT,
        {"family": "Beta-PERT", "a": -2.0, "m": -2.0, "b": -2.0,
         "start_state": _state_words(start)},
        expected_exact={
            "value": _n(degenerate_beta.value),
            "uniforms_consumed": 0,
            "state_unchanged": True,
            "shape_ratio_formed": False,
        },
        note="r is never formed, so 0/0 cannot arise.",
    ))
    for label, mode, dispatch in (("endpoint_low", 0.0, "BC"), ("endpoint_high", 100.0, "BC")):
        shape = prepare_beta_pert(0.0, mode, 100.0)
        cases.append(_case(
            f"sampler.beta.{label}",
            "E_cheng",
            f"Beta-PERT with the mode at an endpoint reaches {dispatch} by the rule",
            EXACT,
            {"family": "Beta-PERT", "a": 0.0, "m": mode, "b": 100.0},
            expected_exact={
                "dispatch": shape.dispatch,
                "alpha": _n(shape.alpha),
                "beta": _n(shape.beta),
                "alpha_plus_beta": _n(shape.alpha + shape.beta),
            },
            note="Equality belongs to BC; the endpoints are not special-cased.",
        ))

    # --- Bernoulli ----------------------------------------------------------
    from .sim_sample import _bernoulli_from_u

    rows = []
    for probability in (0.0, 0.25, 0.5, 1.0):
        for u in (0.0000001, 0.24999999, 0.25, 0.25000001, 0.5, 0.9999999):
            rows.append({
                "u": _n(u),
                "probability": _n(probability),
                "occurred": _bernoulli_from_u(u, probability),
            })
    cases.append(_case(
        "sampler.bernoulli.decision_table",
        "D_bernoulli",
        "Bernoulli occurrence on injected uniforms, both sides of p",
        EXACT,
        {"rule": "occurred = u < probability", "comparison_operator": "strictly_less_than"},
        expected_exact={"rows": rows},
        note=(
            "p = 0 never occurs and p = 1 always occurs, carried by strictness "
            "alone because raw MRG output is strictly inside (0, 1)."
        ),
    ))
    occurrence = bernoulli_occurs(reference, start, 0.5)
    cases.append(_case(
        "sampler.bernoulli.stream_consumption",
        "D_bernoulli",
        "One Bernoulli draw consumes exactly one uniform",
        EXACT,
        {"probability": 0.5, "start_state": _state_words(start)},
        expected_exact={
            "uniform": _n(occurrence.uniform),
            "occurred": occurrence.occurred,
            "uniforms_consumed": 1,
            "state_after": _state_words(occurrence.state),
        },
    ))
    cases.append(_case(
        "sampler.bernoulli.probability_refused",
        "D_bernoulli",
        "A probability outside [0, 1] is refused, not clamped",
        EXACT,
        {"probabilities": [-0.1, 1.1]},
        expected_refusal={"kind": "probability_domain", "reason": "outside [0, 1]; not clamped"},
    ))
    return cases


# ===========================================================================
# D - the whole engine
# ===========================================================================
ENGINE_ITERATIONS = 1000
"""The business minimum, and the smallest run the accepted vector families admit.

Deliberately not larger. Stage A is built often; the Step-0 feasibility
benchmark, the 20,000-iteration analytical cross-checks and the 100,000-iteration
performance model are RETAINED evidence and are not regenerated here.
"""

_HEAD = 6
_TAIL = 3


def _window(values: tuple) -> dict[str, Any]:
    return {
        "count": len(values),
        "head": [_n(value) for value in values[:_HEAD]],
        "tail": [_n(value) for value in values[-_TAIL:]],
        "distinct_count": len(set(values)),
    }


def _components(run) -> list[dict[str, Any]]:
    return [
        {
            "kind": record.kind,
            "permanent_id": record.permanent_id,
            "role": record.role,
            "stream_index": record.stream_index,
            "initial_state": _state_words(record.initial_state),
            "final_state": _state_words(record.final_state),
            "uniforms_consumed": record.uniforms_consumed,
        }
        for record in run.diagnostics
    ]


def _statistics(stats) -> dict[str, Any]:
    return {
        "count": stats.count,
        "mean": _n(stats.mean),
        "sample_standard_deviation": _n(stats.sample_standard_deviation),
        "minimum": _n(stats.minimum),
        "maximum": _n(stats.maximum),
        "quantiles": {label: _n(value) for label, value in stats.percentiles.items()},
    }


class _Engine:
    """A thin binding of the accepted oracle entry points to one set of contracts."""

    def __init__(self, reference, sim, inputs, tolerances):
        self.reference = reference
        self.sim = sim
        self.inputs = inputs
        self.tolerances = tolerances

    def run(self, payload: dict[str, Any], seed: int,
            iterations: int = ENGINE_ITERATIONS):
        prepared, result = prepare_simulation(
            self.reference, self.sim, self.inputs, to_model(payload), self.tolerances,
            effective_seed=seed, iterations=iterations,
        )
        return run_simulation(self.reference, prepared), prepared, result


def _engine_cases(engine: _Engine) -> list[dict[str, Any]]:
    cases: list[dict[str, Any]] = []

    # --- F2: exact by construction, and genuinely stochastic ----------------
    exact_payload = _payload(cost_lines=[_cost("CL-001", "Uniform", 0.0, None, 1.0, 1.0)])
    run, prepared, _ = engine.run(exact_payload, 12345)
    cases.append(_case(
        "engine.exact_friendly.unit_interval",
        "F2_engine_exact",
        "One Uniform(0,1) Cost Line with Quantity 1: each total IS the drawn uniform",
        EXACT,
        {"model": exact_payload, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_exact={
            "result_digest": run.result_digest,
            "total_nominal": _window(run.total_nominal),
            "total_pv": _window(run.total_pv),
            "components": _components(run),
            "statistics": {"nominal": _statistics(run.summary.nominal)},
        },
        note=(
            "Exact at every step and therefore an F2 fixture: (1-u)*0 + u*1 is u "
            "bit for bit, Knom = Kpv = 1, Quantity is 1, and a one-term "
            "accumulation cannot round. The digest is exact because the uniform "
            "it hashes is exact (plan layer A). No rejection path exists here."
        ),
    ))

    # --- F2: dyadic, several drivers, engine plumbing exact ------------------
    dyadic_payload = _payload(
        cost_lines=[
            _cost("CL-001", "Uniform", 256.0, 999.0, 256.0, 2.0),
            _cost("CL-002", "Triangular", 0.5, 0.5, 0.5, 8.0),
            _cost("CL-003", "Uniform", -32.0, None, -32.0, 0.25),
        ],
        risks=[_risk("R-001", "Uniform", 64.0, None, 64.0, 0.5)],
    )
    run, prepared, result = engine.run(dyadic_payload, 12345)
    cases.append(_case(
        "engine.exact_friendly.dyadic_mixed",
        "F2_engine_exact",
        "Dyadic degenerate drivers and one certain-severity Risk: two exact totals",
        EXACT,
        {"model": dyadic_payload, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_exact={
            "result_digest": run.result_digest,
            "total_nominal": _window(run.total_nominal),
            "distinct_totals": sorted(_n(value) for value in set(run.total_nominal)),
            "deterministic_base_a_nominal": _n(result.totals.a_nom),
            "components": _components(run),
        },
        note=(
            "512 + 4 - 8 = 508 without the Risk, 572 with it. The Uniform's "
            "populated Most Likely of 999 reaches nothing. Only the occurrence "
            "stream draws: every severity and value stream consumes zero."
        ),
    ))

    # --- Quantity applied exactly once --------------------------------------
    quantity_rows = []
    for quantity in (1.0, 2.0, 8.0):
        payload = _payload(cost_lines=[_cost("CL-001", "Uniform", 250.0, None, 250.0, quantity)])
        one, _, _ = engine.run(payload, 12345)
        quantity_rows.append({
            "quantity": _n(quantity),
            "total": _n(one.total_nominal[0]),
            "applied_twice_would_be": _n(250.0 * quantity * quantity),
        })
    cases.append(_case(
        "engine.cost_line.quantity_applied_once",
        "F2_engine_exact",
        "The retained total is LINEAR in Quantity, not quadratic and not flat",
        EXACT,
        {"unit_cost": 250.0, "quantities": [1.0, 2.0, 8.0],
         "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_exact={"rows": quantity_rows},
        note=(
            "The sample is UNIT cost. Quantity is deterministic, sits outside "
            "the distribution and is applied exactly once. Probability never "
            "appears on a Cost Line."
        ),
    ))

    # --- D6-18b -------------------------------------------------------------
    d618_runs = []
    for probability in (0.2, 0.8):
        payload = _payload(
            cost_lines=[_cost("CL-001", "Uniform", 10.0, None, 10.0, 1.0)],
            risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, probability)],
        )
        one, _, _ = engine.run(payload, 999)
        severity = next(r for r in one.diagnostics if r.role == "severity")
        occurrence = next(r for r in one.diagnostics if r.role == "occurrence")
        d618_runs.append({
            "probability": _n(probability),
            "model": payload,
            "occurrences": sum(1 for total in one.total_nominal if total != 10.0),
            "occurrence_uniforms_consumed": occurrence.uniforms_consumed,
            "severity_uniforms_consumed": severity.uniforms_consumed,
            "severity_initial_state": _state_words(severity.initial_state),
            "severity_final_state": _state_words(severity.final_state),
            "python_reference_result_digest": one.result_digest,
        })
    cases.append(_case(
        "engine.risk.d6_18b_unconditional_severity",
        "D_bernoulli",
        "Severity is sampled every Risk iteration, whatever the occurrence decided",
        EXACT,
        {"seed": 999, "iterations": ENGINE_ITERATIONS,
         "probabilities": [0.2, 0.8], "severity": "Triangular(100, 200, 400)"},
        expected_exact={
            "runs": [
                {key: value for key, value in row.items()
                 if key != "python_reference_result_digest"}
                for row in d618_runs
            ],
            "severity_consumption_equals_iterations": True,
            "severity_final_state_identical_across_probabilities": (
                d618_runs[0]["severity_final_state"] == d618_runs[1]["severity_final_state"]
            ),
            "occurrence_counts_differ": (
                d618_runs[0]["occurrences"] != d618_runs[1]["occurrences"]
            ),
        },
        python_reference={
            "result_digests": [row["python_reference_result_digest"] for row in d618_runs],
            "note": _PYTHON_REFERENCE_NOTE,
        },
        note=(
            "Consumption is a property of the DISTRIBUTION, not of the "
            "occurrence. Changing only Probability leaves the severity stream "
            "consuming identically and ending in the same state, which is what "
            "makes two such runs comparable at all."
        ),
    ))

    degenerate_severity = _payload(
        risks=[_risk("R-001", "Triangular", 90.0, 90.0, 90.0, 0.4)]
    )
    run, _, _ = engine.run(degenerate_severity, 12345)
    cases.append(_case(
        "engine.risk.degenerate_severity_zero_consumption",
        "D_bernoulli",
        "A degenerate severity is still invoked, consumes zero and leaves its stream alone",
        EXACT,
        {"model": degenerate_severity, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_exact={
            "result_digest": run.result_digest,
            "distinct_totals": sorted(_n(value) for value in set(run.total_nominal)),
            "components": _components(run),
        },
    ))

    # --- F1: general whole-engine, no rejection path -------------------------
    f1_payload = _payload(
        cost_lines=[
            _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, 2.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, 1.0),
        ],
        risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, 0.3)],
    )
    run, prepared, result = engine.run(f1_payload, 12345)
    cases.append(_case(
        "engine.general.no_beta",
        "F1_engine_bounded",
        "General seeded simulation with no Beta driver: whole engine, end to end",
        TOLERANCE_BOUNDED,
        {"model": f1_payload, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected={
            "result_digest": run.result_digest,
            "total_nominal": _window(run.total_nominal),
            "total_pv": _window(run.total_pv),
            "statistics": {
                "nominal": _statistics(run.summary.nominal),
                "pv": _statistics(run.summary.pv),
            },
        },
        expected_exact={
            "components": _components(run),
            "deterministic_base_a_nominal": _n(result.totals.a_nom),
            "analytical_expected_nominal": _n(result.totals.e_nom),
        },
        note=(
            "No rejection path, so the streams cannot desynchronise - but a "
            "transformed sample may still differ by an ULP, so the digest is "
            "NOT an exact cross-language expectation here (plan layer F1)."
        ),
    ))

    # --- G: with a Beta driver, distributional only -------------------------
    beta_payload = _payload(
        cost_lines=[
            _cost("CL-001", "Beta-PERT", 5.0, 7.0, 20.0, 3.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, 1.0),
        ],
        risks=[_risk("R-001", "Beta-PERT", 100.0, 200.0, 400.0, 0.5)],
    )
    run, prepared, result = engine.run(beta_payload, 12345)
    cases.append(_case(
        "engine.general.with_beta",
        "G_engine_statistical",
        "Seeded simulation containing Beta-PERT drivers",
        STATISTICAL,
        {"model": beta_payload, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected={
            "statistics": {
                "nominal": _statistics(run.summary.nominal),
                "pv": _statistics(run.summary.pv),
            },
            "analytical_expected_nominal": _n(result.totals.e_nom),
        },
        python_reference={
            "result_digest": run.result_digest,
            "components": _components(run),
            "note": _PYTHON_REFERENCE_NOTE,
        },
        note=(
            "Sample-for-sample cross-language identity is NOT claimed: a "
            "rejection sampler can legitimately desynchronise two "
            "implementations. Confidence comes from the exact layers plus a "
            "distributional check against the analytical expectation."
        ),
    ))

    # --- G2 / G3 / seed scope: same-runtime relations ------------------------
    replay_a, _, _ = engine.run(f1_payload, 12345)
    replay_b, _, _ = engine.run(f1_payload, 12345)
    cases.append(_case(
        "engine.replay.same_seed_identical",
        "G2_same_runtime",
        "Same prepared model, same seed, same versions: identical retained output",
        SAME_RUNTIME_ONLY,
        {"model": f1_payload, "seed": 12345, "iterations": ENGINE_ITERATIONS, "runs": 2},
        expected={
            "relation": "equal",
            "fields": ["total_nominal", "total_pv", "result_digest"],
            "tolerance": None,
        },
        python_reference={
            "result_digest": replay_a.result_digest,
            "identical": replay_a.result_digest == replay_b.result_digest,
            "note": _PYTHON_REFERENCE_NOTE,
        },
    ))

    reordered = _payload(
        cost_lines=[f1_payload["cost_lines"][1], f1_payload["cost_lines"][0]],
        risks=list(f1_payload["risks"]),
    )
    reorder_run, reorder_prepared, _ = engine.run(reordered, 12345)
    cases.append(_case(
        "engine.row_order.invariant",
        "G3_same_runtime",
        "A physically reordered register produces the same run",
        SAME_RUNTIME_ONLY,
        {"model_a": f1_payload, "model_b": reordered, "seed": 12345,
         "iterations": ENGINE_ITERATIONS},
        expected={
            "relation": "equal",
            "fields": ["prepared_driver_order", "component_assignment",
                       "total_nominal", "total_pv", "result_digest"],
        },
        expected_exact={
            "canonical_driver_order": [
                driver.permanent_id for driver in reorder_prepared.drivers
            ],
        },
        python_reference={
            "result_digest": reorder_run.result_digest,
            "identical_to_canonical_order": (
                reorder_run.result_digest == replay_a.result_digest
            ),
            "note": _PYTHON_REFERENCE_NOTE,
        },
    ))

    seeded = {}
    for seed in (1, 2, 12345):
        one, _, _ = engine.run(f1_payload, seed)
        seeded[seed] = one.result_digest
    cases.append(_case(
        "engine.seed.non_degenerate_divergence",
        "G2_same_runtime",
        "Different seeds diverge ON THIS FIXTURE, whose uncertainty reaches the total",
        SAME_RUNTIME_ONLY,
        {"model": f1_payload, "seeds": [1, 2, 12345], "iterations": ENGINE_ITERATIONS},
        expected={
            "relation": "all_different",
            "fields": ["result_digest"],
            "scope": "this fixture only",
        },
        python_reference={
            "result_digests": {str(seed): digest for seed, digest in seeded.items()},
            "note": _PYTHON_REFERENCE_NOTE,
        },
        note=(
            "The universal claim 'different seed -> different digest' is "
            "WITHDRAWN and is not encoded anywhere in this corpus. What is "
            "universal is that different accepted seeds give different initial "
            "stream states."
        ),
    ))

    # A FULLY degenerate fixture: every distribution degenerate AND the Risk
    # certain, so even the occurrence decision cannot vary. The p = 0.5 fixture
    # above is deliberately NOT reused here - its occurrence sequence does
    # depend on the seed, and using it would prove the opposite of the point.
    fully_degenerate = _payload(
        cost_lines=list(dyadic_payload["cost_lines"]),
        risks=[_risk("R-001", "Uniform", 64.0, None, 64.0, 1.0)],
    )
    degenerate_digests = {}
    degenerate_totals = set()
    for seed in (1, 2, 2147483646):
        one, _, _ = engine.run(fully_degenerate, seed)
        degenerate_digests[seed] = one.result_digest
        degenerate_totals.update(one.total_nominal)
    cases.append(_case(
        "engine.seed.degenerate_equal_digest",
        "F2_engine_exact",
        "A fully degenerate fixture gives the SAME digest for every seed",
        EXACT,
        {"model": fully_degenerate, "seeds": [1, 2, 2147483646],
         "iterations": ENGINE_ITERATIONS},
        expected_exact={
            "result_digests": {str(seed): digest
                               for seed, digest in degenerate_digests.items()},
            "all_equal": len(set(degenerate_digests.values())) == 1,
            "distinct_totals": sorted(_n(value) for value in degenerate_totals),
        },
        note=(
            "ACCEPTED behaviour, not a defect. A model with no uncertainty has "
            "nothing for the seed to vary: every distribution is degenerate and "
            "the Risk is certain, so the occurrence stream still consumes its "
            "uniform per iteration but the decision - and therefore the total - "
            "cannot change."
        ),
    ))
    return cases


# ===========================================================================
# E - the result digest (D6-17)
# ===========================================================================
_DIGEST_NOMINAL = (1234567.8901234567, 0.0, -0.5, 1e300, 5e-324)
_DIGEST_PV = (1000000.0, 0.0, -0.25, 1e299, 1e-320)


def _digest_case(identifier: str, title: str, version: int,
                 nominal: tuple, pv: tuple, note: str | None = None) -> dict[str, Any]:
    return _case(
        identifier,
        "E_digest",
        title,
        EXACT,
        {
            "sim_method_version": version,
            "total_nominal": [_n(value) for value in nominal],
            "total_pv": [_n(value) for value in pv],
        },
        expected_exact={
            "canonical_stream": result_digest_stream(version, nominal, pv),
            "result_digest": result_digest(version, nominal, pv),
        },
        note=note,
    )


def _digest_cases(sim: SimContract) -> list[dict[str, Any]]:
    reversed_nominal = tuple(reversed(_DIGEST_NOMINAL))
    reversed_pv = tuple(reversed(_DIGEST_PV))
    perturbed = (math.nextafter(_DIGEST_NOMINAL[0], math.inf),) + _DIGEST_NOMINAL[1:]

    cases = [
        _digest_case("digest.base", "The base five-record framing vector",
                     1, _DIGEST_NOMINAL, _DIGEST_PV),
        _digest_case("digest.reversed_iteration_order",
                     "Reversing the retained order is a different digest",
                     1, reversed_nominal, reversed_pv,
                     "Why the samples are never sorted for the digest."),
        _digest_case("digest.nominal_and_pv_swapped",
                     "Swapping the two measures is a different digest",
                     1, _DIGEST_PV, _DIGEST_NOMINAL),
        _digest_case("digest.one_iteration_dropped",
                     "A dropped record is a different digest",
                     1, _DIGEST_NOMINAL[:-1], _DIGEST_PV[:-1]),
        _digest_case("digest.one_ulp_perturbation",
                     "One ULP in one total is a different digest",
                     1, perturbed, _DIGEST_PV),
        _digest_case("digest.version_2",
                     "The version field is hashed, so a version change is visible",
                     2, _DIGEST_NOMINAL, _DIGEST_PV),
        _digest_case("digest.empty",
                     "The empty framing vector, which no real run can produce",
                     1, (), (),
                     "The business minimum is at least 1000 iterations; this "
                     "exercises the framing alone."),
    ]

    block = sim.raw["result_digest"]
    cases.append(_case(
        "digest.grammar",
        "E_digest",
        "The locked D6-17 grammar",
        EXACT,
        {"owner": "sim_contract.yaml", "section": "result_digest"},
        expected_exact={
            "stream_tag": block["stream_tag"],
            "section_name": block["section_name"],
            "record_field_count": block["record_field_count"],
            "record_fields": list(block["record_fields"]),
            "field_types": list(block["field_types"]),
            "iteration_index_origin": block["iteration_index_origin"],
            "version_field_source": block["version_field_source"],
            "samples_sorted_for_digest": block["samples_sorted_for_digest"],
            "equality": block["equality"],
            "grammar": dict(block["grammar"]),
        },
    ))
    return cases


# ===========================================================================
# I - the request fingerprint
#
# The analytical prefix is the ACCEPTED Phase-5 case-26 reference stream, taken
# from `calc_cases.reference_stream` unchanged: its HEADER/COST/RISK bytes are
# not regenerated, re-encoded or hashed as a field here. The SIM extension is
# appended to it, and the whole thing is handed to the accepted
# `calc_fingerprint` hash. THERE IS NO SECOND HASH IMPLEMENTATION.
# ===========================================================================
_REQUEST_ENCODERS = {"F_I": integer_field, "F_S": text_field, "F_N": number_field}
"""Contract field-type name to the accepted Phase-5 encoder. Nothing here decides
which encoder a field uses - `sim_contract.yaml` does, and this maps its answer
onto the primitive that implements it."""


def request_sim_section(
    sim: SimContract, iterations: int, seed_mode: str, supplied_seed: int | None = None
) -> str:
    """The SIM extension, built from the contract's own locked grammar.

    Every shape decision - the record count, which fields the mode carries, the
    order they carry them in and the encoder each one uses - is READ from
    `sim_contract.yaml`. This function chooses none of them, so a contract change
    moves the bytes and a test pinning the literals notices.
    """
    block = sim.raw["request_fingerprint"]["sim_section"]
    shape = block["effective_records"].get(seed_mode)
    if shape is None:
        raise SimOracleError(f"unknown seed mode {seed_mode!r}")

    supply = {
        "iterations": iterations,
        "seed_mode": seed_mode,
        "supplied_seed": supplied_seed,
        "rng_version": sim.rng_version,
        "sim_method_version": sim.sim_method_version,
    }
    fields = []
    for name in shape["fields"]:
        value = supply[name]
        if value is None:
            raise SimOracleError(
                f"the {seed_mode} record needs {name}, and none was supplied"
            )
        fields.append(_REQUEST_ENCODERS[block["field_types"][name]](value))
    if len(fields) != int(shape["field_count"]):  # pragma: no cover - read from one source
        raise SimOracleError(f"the {seed_mode} record is not its declared length")
    if supplied_seed is not None and "supplied_seed" not in shape["fields"]:
        raise SimOracleError(
            f"{seed_mode} carries no supplied seed; passing one would silently drop it"
        )

    records = [tuple(fields)] * int(block["record_count"])
    return encode_section(str(block["name"]), records)


def request_fingerprint_stream(
    sim: SimContract, calc: CalcContract, iterations: int, seed_mode: str,
    supplied_seed: int | None = None,
) -> str:
    """`accepted analytical prefix` + `SIM extension`, byte for byte."""
    return reference_stream(calc.fingerprint_version) + request_sim_section(
        sim, iterations, seed_mode, supplied_seed
    )


def request_fingerprint(
    sim: SimContract, calc: CalcContract, iterations: int, seed_mode: str,
    supplied_seed: int | None = None,
) -> str:
    """The 16-character request fingerprint. Equality exact."""
    return fingerprint(
        request_fingerprint_stream(sim, calc, iterations, seed_mode, supplied_seed)
    )


_REQUEST_ITERATIONS = 1000
_REQUEST_ITERATIONS_ALT = 1001


def _request_case(
    sim: SimContract, calc: CalcContract, identifier: str, title: str,
    iterations: int, seed_mode: str, supplied_seed: int | None = None,
    note: str | None = None,
) -> dict[str, Any]:
    suffix = request_sim_section(sim, iterations, seed_mode, supplied_seed)
    prefix = reference_stream(calc.fingerprint_version)
    inputs: dict[str, Any] = {
        "analytical_prefix_digest": fingerprint(prefix),
        "iterations": iterations,
        "seed_mode": seed_mode,
    }
    if supplied_seed is not None:
        inputs["supplied_seed"] = supplied_seed
    return _case(
        identifier,
        "I_request_fingerprint",
        title,
        EXACT,
        inputs,
        expected_exact={
            "sim_suffix": suffix,
            "sim_suffix_code_units": utf16_length(suffix),
            "request_fingerprint": fingerprint(prefix + suffix),
        },
        note=note,
    )


def _request_fingerprint_cases(
    sim: SimContract, inputs: InputContract, calc: CalcContract,
    reference: RngReference,
) -> list[dict[str, Any]]:
    block = sim.raw["request_fingerprint"]
    section = block["sim_section"]
    cases = [_case(
        "request_fingerprint.grammar",
        "I_request_fingerprint",
        "The locked SIM-extension grammar",
        EXACT,
        {"owner": "sim_contract.yaml", "section": "request_fingerprint.sim_section"},
        expected_exact={
            "section_order": list(block["section_order"]),
            "analytical_prefix": list(block["analytical_prefix"]),
            "section_name": section["name"],
            "record_count": int(section["record_count"]),
            "fields": list(section["fields"]),
            # Parallel to `fields`, exactly as the result-digest case pairs its
            # `record_fields` with its `field_types`.
            "field_types": [section["field_types"][name] for name in section["fields"]],
            "encoded_field_names": bool(section["encoded_field_names"]),
            "auto_field_count": int(section["effective_records"]["AUTO"]["field_count"]),
            "auto_fields": list(section["effective_records"]["AUTO"]["fields"]),
            "fixed_field_count": int(section["effective_records"]["FIXED"]["field_count"]),
            "fixed_fields": list(section["effective_records"]["FIXED"]["fields"]),
            "auto_supplied_seed_representation": section["auto_supplied_seed_representation"],
            "excluded_fields": list(section["excluded_fields"]),
            "stream_tag_repeated_in_extension": bool(
                section["stream_tag_repeated_in_extension"]),
            "analytical_fingerprint_hashed_as_a_field": bool(
                section["analytical_fingerprint_hashed_as_a_field"]),
            "grammar": dict(section["grammar"]),
        },
        note=(
            "The SIM extension carries no stream tag and no stream version of "
            "its own: it is a section of the accepted PCCM-FP stream."
        ),
    )]
    cases.append(_request_case(
        sim, calc, "request_fingerprint.auto.1000",
        "AUTO at the business minimum", _REQUEST_ITERATIONS, "AUTO",
        note=("AUTO carries FOUR fields. The supplied seed does not exist here, "
              "which is why two AUTO runs of the same question share one request "
              "fingerprint and stay CURRENT while C21 is blank."),
    ))
    cases.append(_request_case(
        sim, calc, "request_fingerprint.fixed.seed_1",
        "FIXED at the lowest accepted seed", _REQUEST_ITERATIONS, "FIXED", 1,
    ))
    cases.append(_request_case(
        sim, calc, "request_fingerprint.fixed.seed_max",
        "FIXED at the highest accepted seed", _REQUEST_ITERATIONS, "FIXED",
        reference.seed_max,
        note=("The seed bound is the input contract's, read not restated; this "
              "grammar owns type and presence, never admissibility."),
    ))
    cases.append(_request_case(
        sim, calc, "request_fingerprint.auto.1001",
        "One more iteration is a different request", _REQUEST_ITERATIONS_ALT, "AUTO",
        note="Proves the iteration count participates in the stream.",
    ))

    distinct = {
        case["expected_exact"]["request_fingerprint"]
        for case in cases if "request_fingerprint" in case.get("expected_exact", {})
    }
    if len(distinct) != 4:  # pragma: no cover - a collision would be a defect
        raise SimOracleError("the request-fingerprint vectors are not pairwise distinct")
    return cases


# ===========================================================================
# F - statistics
# ===========================================================================
_HAND_SAMPLE = (2.0, 4.0, 4.0, 4.0, 5.0, 5.0, 7.0, 9.0)
_TYPE7_VECTORS = (
    ((5.0,), (0.0, 0.5, 1.0)),
    ((10.0, 20.0), (0.0, 0.25, 0.5, 1.0)),
    ((10.0, 20.0, 60.0), (0.25, 0.5, 0.75)),
    ((1.0, 2.0, 3.0, 4.0), (0.0, 1.0 / 3.0, 0.5, 0.9, 1.0)),
    (tuple(float(value) for value in range(1, 11)), (0.1, 1.0 / 3.0, 0.5, 0.9)),
)
_CONSTANT_SAMPLES = (0.1, 1.1, 1.0e100, 1.5e308, 5e-324, -12345.678, 0.0)


def _type7_rows(values: tuple, points: tuple) -> list[dict[str, Any]]:
    rows = []
    ordered = sorted(values)
    count = len(ordered)
    for p in points:
        h = (count - 1) * p
        lo = math.floor(h)
        hi = min(lo + 1, count - 1)
        exact_by_construction = (h - lo) == 0.0 or ordered[lo] == ordered[hi]
        rows.append({
            "p": _n(p),
            "h": _n(h),
            "lo": lo,
            "hi": hi,
            "f": _n(h - lo),
            "value": _n(percentile_type7(values, p)),
            "comparison": EXACT if exact_by_construction else TOLERANCE_BOUNDED,
        })
    return rows


def _statistics_cases(sim: SimContract, inputs: InputContract) -> list[dict[str, Any]]:
    ladder = resolve_percentile_ladder(sim, inputs)
    cases: list[dict[str, Any]] = [
        _case(
            "statistics.ladder.resolved",
            "F_statistics",
            "The reported quantile ladder, resolved from its two owners",
            EXACT,
            {
                "selectable_owner": "input_contract.yaml: config_tables.confidence_levels",
                "fixed_owner": "sim_contract.yaml: statistics.fixed_nonselectable_percentiles",
            },
            expected_exact={
                "ordered": list(ladder.ordered),
                "count": len(ladder.ordered),
                "fixed_non_selectable": list(ladder.fixed),
                "selectable": list(ladder.selectable),
                "headline": list(ladder.headline),
                "points": [{"label": label, "p": _n(p)} for label, p in ladder.points],
            },
            note="P10 is reported and is NOT selectable. No ladder is restated here.",
        ),
        _case(
            "statistics.mean.hand_vector",
            "F_statistics",
            "Sample mean of an exactly representable hand vector",
            EXACT,
            {"values": [_n(value) for value in _HAND_SAMPLE]},
            expected_exact={"mean": _n(sample_mean(list(_HAND_SAMPLE)))},
        ),
        _case(
            "statistics.sd.hand_vector",
            "F_statistics",
            "Sample standard deviation with divisor n - 1",
            TOLERANCE_BOUNDED,
            {"values": [_n(value) for value in _HAND_SAMPLE], "divisor": "n_minus_1"},
            expected={
                "sample_standard_deviation": _n(
                    sample_standard_deviation(list(_HAND_SAMPLE))
                ),
            },
            expected_exact={
                "population_standard_deviation_would_be": _n(2.0),
                "divisor": "n - 1",
            },
            note="sqrt(32/7) is not exactly representable, so the value is bounded; "
                 "the population divisor gives 2.0 and is a different number.",
        ),
        _case(
            "statistics.sd.zero_variance",
            "F_statistics",
            "A two-point sample with no spread has exactly zero deviation",
            EXACT,
            {"values": [_n(10.0)] * 3},
            expected_exact={
                "mean": _n(10.0),
                "sample_standard_deviation": _n(0.0),
            },
        ),
        _case(
            "statistics.sd.refused_below_two",
            "F_statistics",
            "A sample standard deviation of fewer than two observations is refused",
            EXACT,
            {"counts": [0, 1]},
            expected_refusal={
                "kind": "sample_size",
                "reason": "the divisor is n - 1, which does not exist; no value is invented",
            },
        ),
    ]

    for index, (values, points) in enumerate(_TYPE7_VECTORS, start=1):
        cases.append(_case(
            f"statistics.quantile.type7.n{len(values)}",
            "F_statistics",
            f"Hyndman-Fan type 7 hand vectors, n = {len(values)}",
            TOLERANCE_BOUNDED,
            {
                "values": [_n(value) for value in values],
                "formula": {
                    "h": "(n - 1) * p", "lo": "floor(h)", "hi": "min(lo + 1, n - 1)",
                    "f": "h - lo", "value": "(1 - f) * x[lo] + f * x[hi]",
                },
            },
            expected={"rows": _type7_rows(values, points)},
            expected_exact={
                "interpolation": "convex",
                "sorting": "on copies only",
            },
            note=(
                "Each row carries its own comparison: an integral h selects an "
                "order statistic outright and is EXACT, and so is a row whose "
                "two bracketing order statistics are equal."
            ),
        ))

    constant_rows = []
    for value in _CONSTANT_SAMPLES:
        sample = [value] * 1000
        stats = describe(sample, ladder.points, "constant")
        constant_rows.append({
            "value": _n(value),
            "count": 1000,
            "mean": _n(stats.mean),
            "sample_standard_deviation": _n(stats.sample_standard_deviation),
            "minimum": _n(stats.minimum),
            "maximum": _n(stats.maximum),
            "all_quantiles_equal_the_value": (
                set(stats.percentiles.values()) == {value}
            ),
        })
    cases.append(_case(
        "statistics.constant_sample.zero_dispersion",
        "F_statistics",
        "A constant retained sample has an exact mean and exactly zero dispersion",
        EXACT,
        {"repeated_values": [_n(value) for value in _CONSTANT_SAMPLES], "count": 1000},
        expected_exact={"rows": constant_rows},
        note=(
            "A distribution with one distinct value has no dispersion. "
            "Accumulating the value a thousand times to rediscover it would "
            "report spread that does not exist."
        ),
    ))

    extreme = [-1.7e308, 1.7e308, 1.7e308, 1.7e308]
    cases.append(_case(
        "statistics.scale_safety.near_maximum",
        "H_domain",
        "Opposite-sign totals near Double maximum: the statistics exist",
        TOLERANCE_BOUNDED,
        {"values": [_n(value) for value in extreme]},
        expected={
            "mean": _n(sample_mean(extreme)),
            "sample_standard_deviation": _n(sample_standard_deviation(extreme)),
        },
        expected_exact={
            "naive_sum_overflows": True,
            "naive_sum_of_squares_overflows": True,
            "unguarded_deviation_overflows": True,
            "finite": True,
        },
        note=(
            "A naive accumulation, a naive sum of squares and an unguarded "
            "x - mean all leave Double range here; the accepted helpers do not."
        ),
    ))

    subnormal = [5e-324] * 999 + [1e-323]
    cases.append(_case(
        "statistics.scale_safety.unrepresentable_dispersion",
        "H_domain",
        "A varying sample whose deviation has no Double is refused, not reported as zero",
        EXACT,
        {"values_summary": "5e-324 x 999 then 1e-323", "count": 1000},
        expected_refusal={
            "kind": "numerical_range",
            "stage": "sample standard deviation: rescale",
            "reason": (
                "the true deviation is about 1.6e-325, below the smallest "
                "subnormal; returning 0.0 would claim the sample has no dispersion"
            ),
        },
        expected_exact={"mean_is_representable": True,
                        "mean": _n(sample_mean(subnormal))},
    ))
    return cases


# ===========================================================================
# G - contingency, reporting only
# ===========================================================================
def _contingency_cases(engine: _Engine, sim: SimContract,
                       inputs: InputContract) -> list[dict[str, Any]]:
    ladder = resolve_percentile_ladder(sim, inputs)
    cases: list[dict[str, Any]] = []

    payload = _payload(
        cost_lines=[
            _cost("CL-001", "Triangular", 80.0, 100.0, 150.0, 2.0),
            _cost("CL-002", "Uniform", 10.0, None, 20.0, 1.0),
        ],
        risks=[_risk("R-001", "Triangular", 100.0, 200.0, 400.0, 0.3)],
    )
    run, prepared, result = engine.run(payload, 12345)
    base = deterministic_base_of(result)

    rows = []
    for level in ("P50", "P70", "P80", "P95"):
        contingency = contingency_at(run.summary, level, base)
        rows.append({
            "selected_confidence_level": level,
            "selected_nominal": _n(contingency.selected_nominal),
            "selected_pv": _n(contingency.selected_pv),
            "contingency_nominal": _n(contingency.nominal),
            "contingency_pv": _n(contingency.pv),
        })
    cases.append(_case(
        "contingency.selected_levels",
        "G_contingency",
        "Contingency is the selected quantile total minus the deterministic base A",
        TOLERANCE_BOUNDED,
        {"model": payload, "seed": 12345, "iterations": ENGINE_ITERATIONS,
         "levels": ["P50", "P70", "P80", "P95"]},
        expected={"rows": rows},
        expected_exact={
            "formula": "selected_px_total - deterministic_base_estimate_a",
            "base_nominal": _n(base.nominal),
            "base_pv": _n(base.pv),
            "forbidden_baselines": ["simulation_mean", "analytical_expected_total",
                                    "a_plus_emv"],
            "simulation_mean_is_a_different_number": (
                run.summary.nominal.mean != base.nominal
            ),
            "analytical_expected_is_a_different_number": (
                result.totals.e_nom != base.nominal
            ),
        },
        note=(
            "Reporting only. Changing the Selected Confidence Level reruns no "
            "RNG, alters no retained sample, no digest, no mean and no stored "
            "quantile - every level was computed during the run."
        ),
    ))

    cases.append(_case(
        "contingency.p10_not_selectable",
        "G_contingency",
        "P10 is reported and is refused as a contingency selector",
        EXACT,
        {"selector": "P10", "reported": True},
        expected_refusal={
            "kind": "confidence_level_not_selectable",
            "reason": f"the selectable levels are {list(ladder.selectable)}",
        },
    ))

    dyadic = _payload(cost_lines=[_cost("CL-001", "Uniform", 8.0, None, 8.0, 4.0)])
    dyadic_run, _, dyadic_result = engine.run(dyadic, 12345)
    dyadic_base = deterministic_base_of(dyadic_result)
    negative = contingency_at(dyadic_run.summary, "P50", type(dyadic_base)(1024.0, 1024.0))
    cases.append(_case(
        "contingency.negative_not_clamped",
        "G_contingency",
        "A selected quantile below A gives a negative contingency, reported as such",
        EXACT,
        {"model": dyadic, "seed": 12345, "iterations": ENGINE_ITERATIONS,
         "deterministic_base_nominal": 1024.0, "level": "P50"},
        expected_exact={
            "selected_nominal": _n(negative.selected_nominal),
            "contingency_nominal": _n(negative.nominal),
            "is_negative": negative.nominal < 0.0,
            "clamped": False,
        },
        note="Exact by construction: every value here is a small dyadic rational.",
    ))

    cases.append(_case(
        "contingency.unrepresentable_subtraction",
        "H_domain",
        "A contingency with no representable Double is refused, naming the stage",
        EXACT,
        {"selected_nominal": 1.0e308, "deterministic_base_nominal": -1.0e308},
        expected_refusal={
            "kind": "numerical_range",
            "stage": "contingency nominal",
            "reason": "the difference 2e308 has no Double; it is not returned as infinity",
        },
        expected_exact={
            "representable_neighbour": {
                "selected_nominal": _n(1.0e308),
                "deterministic_base_nominal": _n(-5.0e307),
                "contingency_nominal": _n(1.5e308),
            },
        },
    ))
    return cases


# ===========================================================================
# H - the numerical domain, through the whole pipeline
# ===========================================================================
def _domain_cases(engine: _Engine) -> list[dict[str, Any]]:
    families = (
        ("large_positive", "Uniform", 1.0e307, None, 1.5e307, 1.0),
        ("large_negative", "Uniform", -1.5e307, None, -1.0e307, 1.0),
        ("crossing_zero", "Triangular", -1.0e307, 0.0, 1.0e307, 1.0),
        ("subnormal_scale", "Uniform", 5e-324, None, 1e-320, 1.0),
        ("degenerate", "Uniform", 42.0, None, 42.0, 1.0),
    )
    cases: list[dict[str, Any]] = []
    for label, family, minimum, most_likely, maximum, quantity in families:
        payload = _payload(
            cost_lines=[_cost("CL-001", family, minimum, most_likely, maximum, quantity)]
        )
        run, _, _ = engine.run(payload, 12345)
        stats = run.summary.nominal
        cases.append(_case(
            f"domain.{label}",
            "H_domain",
            f"Extreme-domain family: {label.replace('_', ' ')}",
            EXACT if label == "degenerate" else TOLERANCE_BOUNDED,
            {"model": payload, "seed": 12345, "iterations": ENGINE_ITERATIONS},
            expected={"statistics": {"nominal": _statistics(stats)}},
            expected_exact={
                "all_finite": True,
                "minimum_within_support": True,
                "components": _components(run),
            },
        ))

    rescue = _payload(cost_lines=[
        _cost("CL-001", "Uniform", 1.5e308, None, 1.5e308, 1.0),
        _cost("CL-002", "Uniform", 1.5e308, None, 1.5e308, 1.0),
        _cost("CL-003", "Uniform", -1.5e308, None, -1.5e308, 1.0),
    ])
    run, _, _ = engine.run(rescue, 12345)
    cases.append(_case(
        "domain.accumulation_partial_sum_out_of_range",
        "H_domain",
        "A partial sum leaves Double range while the total does not",
        EXACT,
        {"model": rescue, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_exact={
            "distinct_totals": sorted(_n(value) for value in set(run.total_nominal)),
            "naive_left_to_right_has_no_finite_result": True,
            "mean": _n(run.summary.nominal.mean),
            "sample_standard_deviation": _n(
                run.summary.nominal.sample_standard_deviation
            ),
        },
        note=(
            "1.5e308 + 1.5e308 is infinite, so a naive accumulation refuses an "
            "answer the model has. The accepted signed sum returns 1.5e308, and "
            "because every retained total is then the same Double the statistics "
            "are exact with zero dispersion."
        ),
    ))

    overflow = _payload(cost_lines=[
        _cost("CL-001", "Uniform", 5.0, None, 5.0, 1.0),
        _cost("CL-002", "Uniform", -1.0e308, None, 1.0e308, 1.0e10),
    ])
    try:
        engine.run(overflow, 12345)
    except CalculationRefusal as error:
        refusal_message = str(error)
    else:  # pragma: no cover - the product cannot be represented
        raise SimOracleError("the overflowing contribution was accepted")
    cases.append(_case(
        "domain.contribution_unrepresentable",
        "H_domain",
        "A contribution with no representable Double refuses, naming where",
        EXACT,
        {"model": overflow, "seed": 12345, "iterations": ENGINE_ITERATIONS},
        expected_refusal={
            "kind": "numerical_range",
            "names_iteration_index": "iteration" in refusal_message,
            "names_permanent_id": "CL-002" in refusal_message,
            "names_driver_kind": "Cost Line" in refusal_message,
            "names_stage": "contribution" in refusal_message,
            "no_partial_result_returned": True,
        },
        note=(
            "Phase 5 accepts this model outright - the stochastic line is "
            "symmetric about zero, so every analytical total is zero. Only a "
            "sampled value far enough from the centre overflows, which puts the "
            "failure inside the iteration loop."
        ),
    ))
    return cases


# ===========================================================================
# runtime-only behaviours
# ===========================================================================
def _runtime_only_cases(sim: SimContract) -> list[dict[str, Any]]:
    labels = sim.raw["label_sets"]
    layout = sim.layout
    return [
        _case(
            "runtime.sim_data.geometry",
            "R_runtime",
            "_SimData geometry and the technical iteration ceiling it determines",
            RUNTIME_ONLY,
            {"owner": "sim_contract.yaml: sim_data"},
            expected_exact={
                "sheet": layout.sheet,
                "required_visibility": layout.required_visibility,
                "header_row": layout.header_row,
                "first_iteration_row": layout.first_iteration_row,
                "reserved_rows": layout.reserved_row_count,
                "max_iterations_representable": layout.max_iterations_representable,
            },
            note="No _SimData row is written in Step 5; the sheet stays structurally empty.",
        ),
        _case(
            "runtime.state_labels",
            "R_runtime",
            "The two orthogonal state axes and the seed modes",
            RUNTIME_ONLY,
            {"owner": "sim_contract.yaml: label_sets"},
            expected_exact={
                "sim_state": list(labels["sim_state"]),
                "attempt_result": list(labels["attempt_result"]),
                "seed_mode": list(labels["seed_mode"]),
            },
        ),
        _case(
            "runtime.publication_transaction",
            "R_runtime",
            "Persistence, run-id allocation and attempt metadata are later work",
            RUNTIME_ONLY,
            {"scope": "not implemented in Step 5"},
            expected_exact={
                "run_id_initial": sim.raw["run_id"]["initial"],
                "run_id_first_successful": sim.raw["run_id"]["first_successful_value"],
                "run_id_maximum": sim.raw["run_id"]["maximum"],
                "failure_consumes_run_id": sim.raw["run_id"]["failure_consumes"],
                "auto_nonce_initial": sim.raw["seeding"]["nonce_lifecycle"]["initial"],
            },
            note="Recorded so the transactional boundary is not forgotten. No expectation "
                 "here can be checked until the publication step exists.",
        ),
    ]


# ===========================================================================
# the artefact-shape validator
#
# NOT another production contract. It pins the JSON SEMANTIC TYPES of the
# emitted corpus so that a number cannot quietly become text - which is what
# `allow_nan=False` needs in order to mean anything, because `repr(float("nan"))`
# is the ordinary string "nan" and would serialise without complaint.
#
# It runs at the end of every build, so a typing regression fails the Stage-A
# build rather than waiting for a reader to notice.
# ===========================================================================
INTEGER_KEYS = frozenset({
    "schema_version", "rng_version", "sim_method_version", "case_count",
    "engine_iterations", "iterations", "count", "distinct_count", "draws",
    "sample_count", "index", "lo", "hi", "seed", "auto_nonce", "effective_seed",
    "seed_min", "seed_max", "stream_index", "stream", "total_components",
    "cost_line_count", "risk_count", "base_year", "start_year", "duration",
    "uniforms_consumed", "uniforms_per_sample", "uniforms_per_attempt",
    "uniforms_for_this_sample", "cumulative_uniforms", "proposal_attempts",
    "total_proposal_attempts", "total_uniforms", "occurrences",
    "occurrence_uniforms_consumed", "severity_uniforms_consumed",
    "sim_method_version", "selected_nominal_row", "field_count",
    "record_field_count", "iteration_index_origin", "header_row",
    "first_iteration_row", "reserved_rows", "max_iterations_representable",
    "run_id_initial", "run_id_first_successful", "run_id_maximum",
    "auto_nonce_initial", "record_count", "supplied_seed", "auto_field_count",
    "fixed_field_count", "sim_suffix_code_units",
})

NUMBER_KEYS = frozenset({
    "u", "p", "h", "f", "value", "probability", "alpha", "beta",
    "alpha_plus_beta", "mean", "sample_standard_deviation", "minimum",
    "maximum", "total", "applied_twice_would_be", "quantity",
    "selected_nominal", "selected_pv", "contingency_nominal", "contingency_pv",
    "base_nominal", "base_pv", "deterministic_base_nominal",
    "deterministic_base_a_nominal", "analytical_expected_nominal",
    "population_standard_deviation_would_be", "a", "b", "m", "most_likely",
    "min_value", "max_value", "discount_rate", "rate", "uniform",
    "branch_point_c", "value_at_branch_point",
})

NUMBER_LIST_KEYS = frozenset({
    "first_uniforms", "initial_state", "state_after", "final_state", "state",
    "head", "tail", "distinct_totals", "injected_uniforms", "values",
    "profile_weights", "seeds", "probabilities", "counts", "repeated_values",
    "total_nominal", "total_pv", "quantiles",
})

_NULLABLE_NUMBER_KEYS = frozenset({"most_likely"})

TEXT_SUBTREE_KEYS = frozenset({"grammar", "formula"})
"""Subtrees that are textual by declaration - grammar productions and written
formulas. Their inner key names collide with numeric ones by coincidence
(`stream`, `h`, `lo`), so they are checked as text rather than as numbers."""


def _looks_numeric(text: str) -> bool:
    try:
        float(text)
    except (TypeError, ValueError):
        return False
    return True


def _validate_node(node: Any, path: str, problems: list[str]) -> None:
    if isinstance(node, bool):
        return
    if isinstance(node, float):
        if not math.isfinite(node):
            problems.append(f"{path}: non-finite number {node!r}")
        return
    if isinstance(node, int) or node is None:
        return
    if isinstance(node, str):
        return
    if isinstance(node, list):
        for index, item in enumerate(node):
            _validate_node(item, f"{path}[{index}]", problems)
        return
    if not isinstance(node, dict):
        problems.append(f"{path}: {type(node).__name__} is not a JSON type")
        return

    for key, value in node.items():
        where = f"{path}.{key}"
        if key in TEXT_SUBTREE_KEYS:
            for label, leaf in (
                value.items() if isinstance(value, dict) else [(key, value)]
            ):
                if not isinstance(leaf, str):
                    problems.append(f"{where}.{label}: a declared text subtree must be text")
            continue
        if key.endswith(CANONICAL_SUFFIX):
            # The sidecar is TEXT on purpose: it is an encoding, not a number.
            leaves = value if isinstance(value, list) else (
                list(value.values()) if isinstance(value, dict) else [value]
            )
            for leaf in leaves:
                if not isinstance(leaf, str):
                    problems.append(f"{where}: a canonical sidecar must be text")
            continue
        if key in INTEGER_KEYS:
            for leaf in (value if isinstance(value, list) else [value]):
                if leaf is None:
                    continue
                if isinstance(leaf, bool) or not isinstance(leaf, int):
                    problems.append(
                        f"{where}: expected an integer, got {type(leaf).__name__} "
                        f"{leaf!r}"
                    )
        elif key in NUMBER_KEYS:
            for leaf in (value if isinstance(value, list) else [value]):
                if leaf is None and key in _NULLABLE_NUMBER_KEYS:
                    continue
                if isinstance(leaf, bool) or not isinstance(leaf, (int, float)):
                    problems.append(
                        f"{where}: expected a number, got {type(leaf).__name__} "
                        f"{leaf!r}"
                    )
        elif key in NUMBER_LIST_KEYS:
            leaves = value.values() if isinstance(value, dict) else (
                value if isinstance(value, list) else [value]
            )
            for leaf in leaves:
                if isinstance(leaf, (dict, list)):
                    continue
                if isinstance(leaf, bool) or not isinstance(leaf, (int, float)):
                    problems.append(
                        f"{where}: expected numbers, got {type(leaf).__name__} "
                        f"{leaf!r}"
                    )
        elif isinstance(value, str) and _looks_numeric(value):
            # THE GENERIC RULE. A string that parses as a number is a number
            # wearing a costume; only an explicitly named encoding may do that.
            problems.append(
                f"{where}: the string {value!r} is a semantic number and must be "
                "emitted as a JSON number"
            )
        _validate_node(value, where, problems)


def validate_corpus(document: dict[str, Any]) -> None:
    """Refuse a corpus whose semantic types have drifted. Raises, never repairs."""
    problems: list[str] = []
    _validate_node(document, "$", problems)

    top = {
        "model_version": document["model_version"],
        "sim_contract_version": document["sim_contract_version"],
        "rng_version": document["rng_version"],
        "sim_method_version": document["sim_method_version"],
    }
    for group in document["groups"]:
        for case in group["cases"]:
            versions = case.get("versions")
            if versions is None:
                problems.append(f"case {case['id']}: no version identity")
                continue
            if versions != top:
                problems.append(
                    f"case {case['id']}: version identity {versions} does not "
                    f"project the artefact's {top}"
                )
    if problems:
        raise SimOracleError(
            "the emitted corpus failed artefact-shape validation:\n  "
            + "\n  ".join(problems[:20])
        )


# ===========================================================================
# assembly
# ===========================================================================
def build_sim_cases(
    sim: SimContract, inputs: InputContract, calc: CalcContract, model_version: str
) -> dict[str, Any]:
    """The whole corpus, as plain deterministic data.

    Ordered throughout: the groups are a tuple, each builder appends in a fixed
    sequence, and every mapping is built in a fixed key order. Nothing here reads
    a clock, a path, an environment variable or a set iteration order.
    """
    reference = RngReference.from_contracts(sim, inputs)
    engine = _Engine(reference, sim, inputs, tolerances_from(calc))

    groups = (
        ("A_rng", "MRG backbone: seeds, states and uniforms", _rng_cases(reference)),
        ("B_jump", "Stream jumps and canonical component assignment",
         _stream_assignment_cases(reference)),
        ("B_seed", "Seed semantics: FIXED domain and the AUTO cycle",
         _seed_cases(reference, inputs)),
        ("C_sampler", "Distribution transforms and Bernoulli occurrence",
         _sampler_cases(reference)),
        ("D_engine", "Whole seeded simulations", _engine_cases(engine)),
        ("E_digest", "The result digest and its framing", _digest_cases(sim)),
        ("F_statistics", "Statistics and the reported quantile ladder",
         _statistics_cases(sim, inputs)),
        ("G_contingency", "Contingency, a reporting lookup",
         _contingency_cases(engine, sim, inputs)),
        ("H_domain", "The numerical domain end to end", _domain_cases(engine)),
        ("I_request_fingerprint", "The request fingerprint and its SIM extension",
         _request_fingerprint_cases(sim, inputs, calc, reference)),
        ("R_runtime", "Behaviours that need machinery Step 5 does not build",
         _runtime_only_cases(sim)),
    )

    # Every case carries its own version identity, so a case lifted out of the
    # file still says what it applies to. These are PROJECTIONS of the four
    # top-level values - literally the same objects - not independent copies.
    identity = {
        "model_version": model_version,
        "sim_contract_version": sim.version,
        "rng_version": sim.rng_version,
        "sim_method_version": sim.sim_method_version,
    }

    seen: set[str] = set()
    rendered_groups = []
    for key, title, cases in groups:
        for case in cases:
            if case["id"] in seen:
                raise SimOracleError(f"duplicate case id {case['id']!r}")
            seen.add(case["id"])
            if case["comparison"] not in POLICIES:  # pragma: no cover - _case checks
                raise SimOracleError(f"case {case['id']!r} has no comparison policy")
            case["versions"] = dict(identity)
        rendered_groups.append({"group": key, "title": title, "cases": cases})

    document = {
        "schema_version": SCHEMA_VERSION,
        "model_version": model_version,
        "sim_contract_version": sim.version,
        "rng_version": sim.rng_version,
        "sim_method_version": sim.sim_method_version,
        "purpose": (
            "Machine-readable conformance corpus for the later Phase-6 VBA "
            "implementation steps. Test data only: not a runtime contract, not a "
            "second engine, not a benchmark, not workbook data."
        ),
        "comparison_policies": dict(POLICIES),
        "expectation_blocks": {
            "expected": "governed by the case's own comparison policy",
            "expected_exact": (
                "EXACT whatever the case policy is - draw counts, consumption "
                "and RNG states are exact even where the transformed value they "
                "produced is only tolerance-bounded"
            ),
            "expected_refusal": "the case must be refused, not answered",
            "python_reference": _PYTHON_REFERENCE_NOTE,
        },
        "canonical_sidecar": (
            "A key ending in '" + CANONICAL_SUFFIX + "' holds the accepted "
            "Phase-5 canonical-number encoding of the sibling number. It exists "
            "so binary64 identity has an exact textual form WITHOUT the semantic "
            "value being stringified, and it is emitted only where the "
            "comparison is EXACT."
        ),
        "engine_iterations": ENGINE_ITERATIONS,
        "case_count": len(seen),
        "groups": rendered_groups,
    }
    validate_corpus(document)
    return document
