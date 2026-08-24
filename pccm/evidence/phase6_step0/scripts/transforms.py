"""The accepted plan section 4.0-4.3 transforms, and the degenerate rule.

Two jobs:
  1. the D6-18 + degenerate evidence - a degenerate driver returns `a`, enters
     no sampler and consumes ZERO uniforms, so its component stream is unchanged
     after any number of iterations;
  2. the tolerance model - each transform is written in the ACCEPTED expression
     order and in an algebraically equivalent alternate order, and the two are
     compared, so the tolerance rests on measurement rather than on a number
     somebody liked.
"""
import math

import beta_ref as BR


# --- plan section 4.0: the degenerate rule, one rule for all three families ---
def is_degenerate(a, m, b):
    """Under the accepted ordering a <= m <= b, `a == b` forces `a == m == b`."""
    return a == b


def _scale(a, m, b):
    """Plan section 4.6 conditioning scale."""
    s = max(abs(a), abs(m), abs(b))
    return s if s > 0.0 else 1.0


# --- plan section 4.1: Uniform ------------------------------------------------
def uniform_accepted(u, a, b):
    """The ACCEPTED stable convex form."""
    return (1.0 - u) * a + u * b


def uniform_alternate(u, a, b):
    """Algebraically equivalent, different expression order. NOT the contract."""
    return a + u * (b - a)


# --- plan section 4.2: Triangular --------------------------------------------
def triangular_accepted(u, a, m, b):
    s = _scale(a, m, b)
    an, mn, bn = a / s, m / s, b / s
    c = (mn - an) / (bn - an)
    if u <= c:
        xn = an + math.sqrt(u * (bn - an) * (mn - an))
    else:
        xn = bn - math.sqrt((1.0 - u) * (bn - an) * (bn - mn))
    return xn * s


def triangular_alternate(u, a, m, b):
    """Same mathematics, different grouping and a different scale. NOT the contract."""
    s = _scale(a, m, b)
    an, mn, bn = a / s, m / s, b / s
    width = bn - an
    c = (mn - an) / width
    if u <= c:
        xn = an + math.sqrt(((bn - an) * u) * (mn - an))
    else:
        xn = bn - math.sqrt(((bn - mn) * (1.0 - u)) * (bn - an))
    return xn * s


# --- plan section 4.3: Beta-PERT rescale --------------------------------------
def pert_rescale_accepted(y, a, b):
    return (1.0 - y) * a + y * b


def pert_rescale_alternate(y, a, b):
    return a + y * (b - a)


# --- one driver sample, with the degenerate short-circuit --------------------
def sample_driver(kind, a, m, b, rng, counter):
    """Returns (value, uniforms_consumed_by_this_call).

    The degenerate test happens BEFORE dispatch and before any parameterisation,
    so `r = (m-a)/(b-a)` is never formed and 0/0 cannot arise.
    """
    before = counter["uniforms"]
    if is_degenerate(a, m, b):
        return a, 0                       # no sampler, no uniform, no r
    if kind == "uniform":
        v = uniform_accepted(rng(), a, b)
        counter["uniforms"] += 1
    elif kind == "triangular":
        v = triangular_accepted(rng(), a, m, b)
        counter["uniforms"] += 1
    elif kind == "pert":
        s = _scale(a, m, b)
        an, mn, bn = a / s, m / s, b / s
        r = (mn - an) / (bn - an)
        alpha, beta = BR.pert_shape(r)
        y = BR.cheng_sample(alpha, beta, rng, counter)
        v = pert_rescale_accepted(y, a, b)
    else:
        raise ValueError(kind)
    return v, counter["uniforms"] - before
