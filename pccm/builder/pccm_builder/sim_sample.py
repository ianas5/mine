"""PCCM Phase-6 Step-3 reference: the stochastic transforms.

Turns raw MRG32k3a uniforms into distribution samples and Bernoulli decisions:
Uniform, Triangular, Beta-PERT through the exact locked Cheng BB/BC formulation,
and occurrence. That is the whole of Step 3.

WHAT THIS IS NOT. There is no Monte Carlo loop here, no Cost Line or Risk
contribution, no `Quantity`, `Knom` or `Kpv`, no retained arrays, no
`result_digest`, no statistic and no contingency. Step 4 pairs occurrence with
unconditional severity in the simulation oracle; Step 3 provides the pieces.

--------------------------------------------------------------------------------
CONSUMPTION IS PART OF THE CONTRACT
--------------------------------------------------------------------------------
Every sampler returns the RNG state it produced along with how many uniforms it
consumed and, for a rejection sampler, how many proposal attempts it made. None
of that is hidden behind a mutable generator, because under an
acceptance/rejection sampler the number of uniforms consumed is not a fixed
property of the call - it depends on the values drawn - and a later
implementation has to reproduce it exactly, not merely reproduce the value.

    Uniform, non-degenerate       1 uniform
    Triangular, non-degenerate    1 uniform
    Beta-PERT, non-degenerate     2 x proposal_attempts
    Bernoulli                     1 uniform
    ANY degenerate distribution   0 uniforms, state unchanged

A rejected Cheng proposal consumes both of its uniforms and the retry continues
from the resulting state. There is no rewind.

--------------------------------------------------------------------------------
NUMERICS
--------------------------------------------------------------------------------
The accepted domain is any finite, correctly ordered triple: negative, crossing
zero, near `Double` maximum, subnormal-scale. Phase 6 may not narrow it, so the
rescales use the convex form and the shape arithmetic is done in a conditioned
space. A representable result is never refused because a NAIVE intermediate
would have overflowed; a result that genuinely cannot be represented raises,
naming the family and the numerical stage. Nothing here silently returns `inf`
or `NaN`, and nothing clips.

Every uniform comes from the accepted `RngReference`. No global RNG, no hidden
draw, no `random`, no NumPy.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any

from .contract_loader import ContractError
from .sim_rng import RngReference, RngState

FAMILY_UNIFORM = "Uniform"
FAMILY_TRIANGULAR = "Triangular"
FAMILY_BETA_PERT = "Beta-PERT"

ACCEPTED_FAMILIES = (FAMILY_UNIFORM, FAMILY_TRIANGULAR, FAMILY_BETA_PERT)
"""Exactly three. No aliases, no fourth family, no silent fallback."""

PERT_LAMBDA = 4.0

# The locked Cheng literals, as LITERALS. `1.3862944` is those eight digits and
# is NOT evaluated as `log(4)`; `2.609438` is not `1 + log(5)`; `0.0138889`,
# `0.0416667` and `0.777778` are not `1/72`, `3/72` and `7/9`. Step 0 measured
# what they control: the squeeze literals never change the value of an accepted
# proposal - they change WHICH proposals are accepted, and therefore consumption
# and every draw after it.
_BB_LOG4 = 1.3862944
_BB_ONE_PLUS_LOG5 = 2.609438
_BC_K1_A = 0.0138889
_BC_K1_B = 0.0416667
_BC_K1_C = 0.777778
_BC_LOG4 = 1.3862944


class SimSampleError(ContractError):
    """Raised when a sampler is asked for something the contract forbids.

    Subclasses ContractError so a specification-level refusal reports the same
    way every other PCCM authority failure does.
    """


# ---------------------------------------------------------------------------
# Values
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class SampleResult:
    """One sample, and everything a later implementation must reproduce."""

    value: float
    state: RngState
    uniforms_consumed: int
    proposal_attempts: int = 0


@dataclass(frozen=True)
class BernoulliResult:
    """One occurrence decision. Always exactly one uniform."""

    occurred: bool
    state: RngState
    uniform: float
    uniforms_consumed: int = 1


@dataclass(frozen=True)
class PreparedBetaPert:
    """Per-DRIVER Beta-PERT shape constants, computed once.

    `alpha` and `beta` are fixed for the life of a driver, and so is everything
    Cheng derives from them. A simulation samples one driver 100,000 times, so
    recomputing a square root and two logarithms per iteration would be work the
    shape already settled. Immutable, and holding no RNG state - preparing a
    shape draws nothing.
    """

    a: float
    m: float
    b: float
    alpha: float
    beta: float
    dispatch: str
    degenerate: bool
    # Cheng per-driver terms. Named as the contract names them; `cheng_a` and
    # `cheng_b` are the ORIENTED pair, which is min/max for BB and max/min for
    # BC - opposite, and inverting one returns the mirrored distribution.
    cheng_a: float = 0.0
    cheng_b: float = 0.0
    cheng_alpha: float = 0.0
    cheng_beta: float = 0.0
    cheng_gamma: float = 0.0
    cheng_delta: float = 0.0
    cheng_k1: float = 0.0
    cheng_k2: float = 0.0
    first_parameter_is_oriented_a: bool = True


# ---------------------------------------------------------------------------
# Parameter validation
# ---------------------------------------------------------------------------
def _finite(value: Any, label: str, family: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise SimSampleError(f"{family}: {label} must be a finite number, got {value!r}")
    number = float(value)
    if not math.isfinite(number):
        raise SimSampleError(f"{family}: {label} must be finite, got {number!r}")
    return number


def _check_ordering(family: str, a: float, m: float | None, b: float) -> None:
    """Refuse, never repair. Silently swapping endpoints would turn a data entry
    error into a plausible distribution nobody asked for."""
    if family == FAMILY_UNIFORM:
        if a > b:
            raise SimSampleError(
                f"{family}: Min {a!r} exceeds Max {b!r}. Endpoints are never swapped silently."
            )
        return
    if m is None:
        raise SimSampleError(f"{family}: Most Likely is required")
    if not a <= m <= b:
        raise SimSampleError(
            f"{family}: requires Min <= Most Likely <= Max, got {a!r}, {m!r}, {b!r}. "
            "The ordering is refused, not repaired."
        )


def is_degenerate(family: str, a: float, m: float | None, b: float) -> bool:
    """FAMILY-SPECIFIC, and it has to be.

    A common `a == m == b` predicate made a degenerate Uniform depend on Most
    Likely, which accepted Phase-5 D1 IGNORES numerically. A Uniform with
    `Min = Max` and a populated, unrelated Most Likely would then enter the
    sampler and consume a uniform - so an ignored input would decide RNG
    consumption and every later draw on that component.
    """
    if family == FAMILY_UNIFORM:
        return a == b
    return a == m == b


def _conditioning_scale(*values: float) -> float:
    """Plan section 4.6: `s = max(|a|, |m|, |b|)`, never zero.

    Working on `a/s, m/s, b/s` is what keeps `(b-a)(m-a)` finite for endpoints
    near `Double` maximum, where the naive product overflows long before the
    result does.
    """
    scale = max(abs(v) for v in values)
    return scale if scale > 0.0 else 1.0


def _checked_result(value: float, family: str, stage: str) -> float:
    if not math.isfinite(value):
        raise SimSampleError(
            f"{family}: the {stage} stage produced {value!r}, which is not representable as a "
            "finite Double. The refusal is explicit and names the stage; a silent inf or NaN "
            "would travel into a published total."
        )
    return value


# ---------------------------------------------------------------------------
# Injected-uniform transforms - pure, and NOT a second RNG
# ---------------------------------------------------------------------------
# The caller supplies `u`, so these consume nothing and hold nothing. They exist
# so a hand-derived transform test can pin a branch or a boundary without also
# depending on which uniform the stream happens to produce. Production paths
# obtain `u` only from `RngReference`.
def _uniform_from_u(u: float, a: float, b: float) -> float:
    """The ACCEPTED stable convex form.

    NOT `a + u*(b - a)`: for `a = -MAX, b = +MAX` the difference overflows while
    every convex result is finite, and the accepted numerical-domain authority
    refuses to lose a representable answer to a naive intermediate.
    """
    return _checked_result((1.0 - u) * a + u * b, FAMILY_UNIFORM, "convex rescale")


def _triangular_from_u(u: float, a: float, m: float, b: float) -> float:
    """Inverse CDF, evaluated in the conditioned space and rescaled after."""
    scale = _conditioning_scale(a, m, b)
    an, mn, bn = a / scale, m / scale, b / scale
    width = bn - an
    if width <= 0.0:  # pragma: no cover - degeneracy is handled before dispatch
        raise SimSampleError(f"{FAMILY_TRIANGULAR}: zero width reached the transform")
    c = (mn - an) / width
    if u <= c:
        conditioned = an + math.sqrt(u * width * (mn - an))
    else:
        conditioned = bn - math.sqrt((1.0 - u) * width * (bn - mn))
    return _checked_result(conditioned * scale, FAMILY_TRIANGULAR, "conditioned rescale")


def _bernoulli_from_u(u: float, probability: float) -> bool:
    """`occurred = u < Probability`. STRICT.

    Because raw MRG output is strictly inside `(0,1)`, strictness is what makes
    `p = 0` never occur and `p = 1` always occur - both exactly, with no special
    case anywhere.
    """
    return u < probability


# ---------------------------------------------------------------------------
# Uniform and Triangular
# ---------------------------------------------------------------------------
def sample_uniform(
    reference: RngReference, state: RngState, a: Any, b: Any, m: Any = None
) -> SampleResult:
    """One uniform consumed. Most Likely is ignored, whatever it holds."""
    a = _finite(a, "Min", FAMILY_UNIFORM)
    b = _finite(b, "Max", FAMILY_UNIFORM)
    _check_ordering(FAMILY_UNIFORM, a, None, b)
    if is_degenerate(FAMILY_UNIFORM, a, None, b):
        return SampleResult(value=a, state=state, uniforms_consumed=0)
    draw = reference.next_uniform(state)
    return SampleResult(
        value=_uniform_from_u(draw.uniform, a, b), state=draw.state, uniforms_consumed=1
    )


def sample_triangular(
    reference: RngReference, state: RngState, a: Any, m: Any, b: Any
) -> SampleResult:
    """One uniform consumed."""
    a = _finite(a, "Min", FAMILY_TRIANGULAR)
    m = _finite(m, "Most Likely", FAMILY_TRIANGULAR)
    b = _finite(b, "Max", FAMILY_TRIANGULAR)
    _check_ordering(FAMILY_TRIANGULAR, a, m, b)
    if is_degenerate(FAMILY_TRIANGULAR, a, m, b):
        return SampleResult(value=a, state=state, uniforms_consumed=0)
    draw = reference.next_uniform(state)
    return SampleResult(
        value=_triangular_from_u(draw.uniform, a, m, b), state=draw.state, uniforms_consumed=1
    )


# ---------------------------------------------------------------------------
# Beta-PERT
# ---------------------------------------------------------------------------
def prepare_beta_pert(a: Any, m: Any, b: Any) -> PreparedBetaPert:
    """Per-driver shape constants. Draws nothing and holds no RNG state."""
    a = _finite(a, "Min", FAMILY_BETA_PERT)
    m = _finite(m, "Most Likely", FAMILY_BETA_PERT)
    b = _finite(b, "Max", FAMILY_BETA_PERT)
    _check_ordering(FAMILY_BETA_PERT, a, m, b)

    if is_degenerate(FAMILY_BETA_PERT, a, m, b):
        # r is NEVER formed, so 0/0 cannot arise. That is why degeneracy is
        # detected before parameterisation and not inside it.
        return PreparedBetaPert(
            a=a, m=m, b=b, alpha=0.0, beta=0.0, dispatch="", degenerate=True
        )

    scale = _conditioning_scale(a, m, b)
    an, mn, bn = a / scale, m / scale, b / scale
    r = (mn - an) / (bn - an)
    alpha = 1.0 + PERT_LAMBDA * r
    beta = 1.0 + PERT_LAMBDA * (1.0 - r)

    # Dispatch, and the boundary: EQUALITY BELONGS TO BC. `m = a` gives alpha 1
    # and `m = b` gives beta 1, so both endpoints dispatch to BC - not as a
    # special case bolted on afterwards, but because the rule says so.
    dispatch = "BB" if min(alpha, beta) > 1.0 else "BC"

    if dispatch == "BB":
        ca, cb = min(alpha, beta), max(alpha, beta)
        c_alpha = ca + cb
        c_beta = math.sqrt((c_alpha - 2.0) / (2.0 * ca * cb - c_alpha))
        c_gamma = ca + 1.0 / c_beta
        return PreparedBetaPert(
            a=a, m=m, b=b, alpha=alpha, beta=beta, dispatch="BB", degenerate=False,
            cheng_a=ca, cheng_b=cb, cheng_alpha=c_alpha, cheng_beta=c_beta,
            cheng_gamma=c_gamma, first_parameter_is_oriented_a=(alpha == ca),
        )

    # BC orients the OPPOSITE way to BB. Inverting it is a silent defect: the
    # sampler still returns a valid Beta variate, just of the mirrored
    # distribution.
    ca, cb = max(alpha, beta), min(alpha, beta)
    c_alpha = ca + cb
    c_beta = 1.0 / cb
    c_delta = 1.0 + ca - cb
    c_k1 = c_delta * (_BC_K1_A + _BC_K1_B * cb) / (ca * c_beta - _BC_K1_C)
    c_k2 = 0.25 + (0.5 + 0.25 / c_delta) * cb
    return PreparedBetaPert(
        a=a, m=m, b=b, alpha=alpha, beta=beta, dispatch="BC", degenerate=False,
        cheng_a=ca, cheng_b=cb, cheng_alpha=c_alpha, cheng_beta=c_beta,
        cheng_delta=c_delta, cheng_k1=c_k1, cheng_k2=c_k2,
        first_parameter_is_oriented_a=(alpha == ca),
    )


def _cheng_bb(
    reference: RngReference, state: RngState, shape: PreparedBetaPert
) -> tuple[float, RngState, int]:
    """Cheng BB, exactly as the contract writes it. Returns (y, state, attempts)."""
    a, b = shape.cheng_a, shape.cheng_b
    alpha, beta, gamma = shape.cheng_alpha, shape.cheng_beta, shape.cheng_gamma
    attempts = 0
    current = state
    while True:
        attempts += 1
        first = reference.next_uniform(current)
        second = reference.next_uniform(first.state)
        current = second.state
        u1, u2 = first.uniform, second.uniform
        vlog = math.log(u1 / (1.0 - u1))
        v = beta * vlog
        w = a * math.exp(v)
        z = u1 * u1 * u2
        rr = gamma * v - _BB_LOG4
        s = a + rr - w
        if s + _BB_ONE_PLUS_LOG5 >= 5.0 * z:
            break
        t = math.log(z)
        if s >= t:
            break
        if rr + alpha * math.log(alpha / (b + w)) >= t:
            break
    y = (w / (b + w)) if shape.first_parameter_is_oriented_a else (b / (b + w))
    return y, current, attempts


def _cheng_bc(
    reference: RngReference, state: RngState, shape: PreparedBetaPert
) -> tuple[float, RngState, int]:
    """Cheng BC, exactly as the contract writes it. Returns (y, state, attempts)."""
    a, b = shape.cheng_a, shape.cheng_b
    alpha, beta = shape.cheng_alpha, shape.cheng_beta
    k1, k2 = shape.cheng_k1, shape.cheng_k2
    attempts = 0
    current = state
    while True:
        attempts += 1
        first = reference.next_uniform(current)
        second = reference.next_uniform(first.state)
        current = second.state
        u1, u2 = first.uniform, second.uniform
        if u1 < 0.5:
            y0 = u1 * u2
            z = u1 * y0
            if 0.25 * u2 + z - y0 >= k1:
                continue
        else:
            z = u1 * u1 * u2
            if z <= 0.25:
                vlog = math.log(u1 / (1.0 - u1))
                v = beta * vlog
                w = a * math.exp(v)
                break
            if z >= k2:
                continue
        vlog = math.log(u1 / (1.0 - u1))
        v = beta * vlog
        w = a * math.exp(v)
        if alpha * (math.log(alpha / (b + w)) + v) - _BC_LOG4 >= math.log(z):
            break
    y = (w / (b + w)) if shape.first_parameter_is_oriented_a else (b / (b + w))
    return y, current, attempts


def sample_prepared_beta(
    reference: RngReference, state: RngState, shape: PreparedBetaPert
) -> SampleResult:
    """Sample a prepared shape. Two uniforms per proposal attempt, no rewind."""
    if not isinstance(shape, PreparedBetaPert):
        raise SimSampleError(f"expected a PreparedBetaPert, got {type(shape).__name__}")
    if shape.degenerate:
        return SampleResult(value=shape.a, state=state, uniforms_consumed=0)
    reference.validate_state(state)
    if shape.dispatch == "BB":
        y, advanced, attempts = _cheng_bb(reference, state, shape)
    elif shape.dispatch == "BC":
        y, advanced, attempts = _cheng_bc(reference, state, shape)
    else:  # pragma: no cover - dispatch is set by prepare_beta_pert
        raise SimSampleError(f"unknown Cheng dispatch {shape.dispatch!r}")
    if not 0.0 < y < 1.0:
        raise SimSampleError(
            f"{FAMILY_BETA_PERT}: the Cheng stage produced a Beta variate {y!r} outside (0,1)"
        )
    value = _checked_result(
        (1.0 - y) * shape.a + y * shape.b, FAMILY_BETA_PERT, "convex rescale"
    )
    return SampleResult(
        value=value, state=advanced, uniforms_consumed=2 * attempts, proposal_attempts=attempts
    )


def sample_beta_pert(
    reference: RngReference, state: RngState, a: Any, m: Any, b: Any
) -> SampleResult:
    """Prepare and sample in one call. Equivalent to preparing once per driver."""
    return sample_prepared_beta(reference, state, prepare_beta_pert(a, m, b))


# ---------------------------------------------------------------------------
# Dispatch and Bernoulli
# ---------------------------------------------------------------------------
def sample_distribution(
    reference: RngReference, state: RngState, family: Any, a: Any, m: Any, b: Any
) -> SampleResult:
    """The single entry point. Exactly three families; anything else is refused."""
    if family not in ACCEPTED_FAMILIES:
        raise SimSampleError(
            f"unknown distribution family {family!r}; the accepted families are exactly "
            f"{list(ACCEPTED_FAMILIES)}. There is no alias and no fallback."
        )
    if family == FAMILY_UNIFORM:
        return sample_uniform(reference, state, a, b, m)
    if family == FAMILY_TRIANGULAR:
        return sample_triangular(reference, state, a, m, b)
    return sample_beta_pert(reference, state, a, m, b)


def bernoulli_occurs(
    reference: RngReference, state: RngState, probability: Any
) -> BernoulliResult:
    """One uniform, strict `<`. A separate primitive from severity sampling.

    Step 3 provides the decision; it does not orchestrate D6-18. Pairing
    occurrence with unconditional severity is Step 4's.
    """
    if isinstance(probability, bool) or not isinstance(probability, (int, float)):
        raise SimSampleError(f"Probability must be a number, got {probability!r}")
    probability = float(probability)
    if not math.isfinite(probability):
        raise SimSampleError(f"Probability must be finite, got {probability!r}")
    if not 0.0 <= probability <= 1.0:
        raise SimSampleError(
            f"Probability {probability!r} is outside [0, 1]. It is refused, not clamped."
        )
    draw = reference.next_uniform(state)
    return BernoulliResult(
        occurred=_bernoulli_from_u(draw.uniform, probability),
        state=draw.state,
        uniform=draw.uniform,
    )
