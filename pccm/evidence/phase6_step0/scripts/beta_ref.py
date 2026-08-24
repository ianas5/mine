"""Beta reference: incomplete beta, inverse CDF (candidate A), Cheng BB/BC.

The PCCM Beta-PERT family is ONE-DIMENSIONAL:
    alpha = 1 + 4r,  beta = 5 - 4r,  r in [0,1]   =>   alpha + beta = 6
Nothing outside that line can arise from a PCCM driver, so nothing outside it
is measured here.
"""
import math

OPS = {"betacf_iters": 0}

# --- one counting method, applied to EVERY candidate ---------------------
# Transcendental calls are counted, not estimated. They are the dominant cost
# in both candidate A and candidate B, and counting them the same way in both
# is what makes the two numbers comparable. Cheap arithmetic (add, multiply,
# divide, compare) is counted separately as loop iterations / attempts; no
# flops-per-iteration multiplier is applied anywhere in the retained numbers.
TR = {"log": 0, "log1p": 0, "exp": 0, "sqrt": 0, "lgamma": 0}


def _log(x):
    TR["log"] += 1
    return math.log(x)


def _log1p(x):
    TR["log1p"] += 1
    return math.log1p(x)


def _exp(x):
    TR["exp"] += 1
    return math.exp(x)


def _sqrt(x):
    TR["sqrt"] += 1
    return math.sqrt(x)


def _lgamma(x):
    TR["lgamma"] += 1
    return math.lgamma(x)


def tr_reset():
    for k in TR:
        TR[k] = 0


def tr_total():
    return sum(TR.values())


def tr_snapshot():
    return dict(TR)


# --- libm sensitivity probe (settlement section 7) ------------------------
# A cross-language difference in `log` or `exp` of one ULP does not stay one
# ULP: `exp` turns an ABSOLUTE error in its argument into a RELATIVE error in
# its result, so the amplification is |v|. This measures the resulting relative
# change in the RETURNED SAMPLE, holding the acceptance path fixed, so the
# tolerance policy rests on a measurement rather than on a chosen number.
#
# Off by default. Enabling it changes no arithmetic on the accepted path: the
# perturbed twin is computed alongside and never fed back.
SENS = {"on": False, "max_rel_out": 0.0, "max_abs_v": 0.0, "samples": 0}


def sens_reset(on=True):
    SENS["on"] = on
    SENS["max_rel_out"] = 0.0
    SENS["max_abs_v"] = 0.0
    SENS["samples"] = 0


def sens_snapshot():
    return {k: v for k, v in SENS.items() if k != "on"}


def _sens_record(out, out_pert, v):
    SENS["samples"] += 1
    if abs(v) > SENS["max_abs_v"]:
        SENS["max_abs_v"] = abs(v)
    denom = abs(out) if out != 0.0 else 1.0
    rel = abs(out_pert - out) / denom
    if rel > SENS["max_rel_out"]:
        SENS["max_rel_out"] = rel


# --- floating acceptance-path margins (authorisation section 5.3) --------
# Every accept/reject predicate in Cheng BB and BC is a comparison of two
# Doubles. This records how close each evaluated comparison came to its own
# boundary. It is DIAGNOSTIC: it says how fragile the branch is under a small
# perturbation. It says nothing whatever about VBA, and no cross-language
# conclusion may be drawn from it.
MARGIN = {"evaluated": 0, "min_abs": float("inf"), "min_rel": float("inf"),
          "rel_lt_1e_3": 0, "rel_lt_1e_6": 0, "rel_lt_1e_9": 0,
          "rel_lt_1e_12": 0, "rel_lt_1e_15": 0}


def margin_reset():
    MARGIN["evaluated"] = 0
    MARGIN["min_abs"] = float("inf")
    MARGIN["min_rel"] = float("inf")
    for k in ("rel_lt_1e_3", "rel_lt_1e_6", "rel_lt_1e_9",
              "rel_lt_1e_12", "rel_lt_1e_15"):
        MARGIN[k] = 0


def margin_snapshot():
    return dict(MARGIN)


def _ge(lhs, rhs):
    """`lhs >= rhs`, recording the margin. Every Cheng branch goes through it."""
    d = lhs - rhs
    a = abs(d)
    rel = a / max(abs(lhs), abs(rhs), 1.0)
    MARGIN["evaluated"] += 1
    if a < MARGIN["min_abs"]:
        MARGIN["min_abs"] = a
    if rel < MARGIN["min_rel"]:
        MARGIN["min_rel"] = rel
    if rel < 1e-3:
        MARGIN["rel_lt_1e_3"] += 1
        if rel < 1e-6:
            MARGIN["rel_lt_1e_6"] += 1
            if rel < 1e-9:
                MARGIN["rel_lt_1e_9"] += 1
                if rel < 1e-12:
                    MARGIN["rel_lt_1e_12"] += 1
                    if rel < 1e-15:
                        MARGIN["rel_lt_1e_15"] += 1
    return d >= 0.0


def pert_shape(r):
    return 1.0 + 4.0 * r, 5.0 - 4.0 * r


def betacf(a, b, x, tol=3e-16, maxit=400):
    qab, qap, qam = a + b, a + 1.0, a - 1.0
    c, d = 1.0, 1.0 - qab * x / qap
    if abs(d) < 1e-300:
        d = 1e-300
    d = 1.0 / d
    h = d
    for m in range(1, maxit + 1):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1.0 + aa * d
        c = 1.0 + aa / c
        if abs(d) < 1e-300: d = 1e-300
        if abs(c) < 1e-300: c = 1e-300
        d = 1.0 / d
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1.0 + aa * d
        c = 1.0 + aa / c
        if abs(d) < 1e-300: d = 1e-300
        if abs(c) < 1e-300: c = 1e-300
        d = 1.0 / d
        delta = d * c
        h *= delta
        if abs(delta - 1.0) < tol:
            return h, m
    return h, maxit


def betai(a, b, x):
    if x <= 0.0: return 0.0
    if x >= 1.0: return 1.0
    lbeta = _lgamma(a) + _lgamma(b) - _lgamma(a + b)
    if x < (a + 1.0) / (a + b + 2.0):
        cf, it = betacf(a, b, x)
        OPS["betacf_iters"] = it
        return _exp(a * _log(x) + b * _log1p(-x) - lbeta) * cf / a
    cf, it = betacf(b, a, 1.0 - x)
    OPS["betacf_iters"] = it
    return 1.0 - _exp(b * _log1p(-x) + a * _log(x) - lbeta) * cf / b


def beta_ppf(u, a, b, halvings=60):
    """Candidate A: bisection on the regularised incomplete beta."""
    lo, hi, evals, worst = 0.0, 1.0, 0, 0
    for _ in range(halvings):
        mid = 0.5 * (lo + hi)
        v = betai(a, b, mid)
        evals += 1
        worst = max(worst, OPS["betacf_iters"])
        if v < u: lo = mid
        else: hi = mid
    return 0.5 * (lo + hi), evals, worst


# --- Cheng BB / BC -------------------------------------------------------
# Dispatch, LOCKED at the boundary: BB when min(a,b) > 1, else BC.
def cheng_dispatch(a, b):
    return "BB" if min(a, b) > 1.0 else "BC"


def cheng_bb(a0, b0, rng, counter):
    """Cheng (1978) BB, for min(a,b) > 1. Returns x in (0,1)."""
    a = min(a0, b0); b = max(a0, b0)
    alpha = a + b
    beta = _sqrt((alpha - 2.0) / (2.0 * a * b - alpha))
    gamma = a + 1.0 / beta
    while True:
        counter["attempts"] += 1
        u1 = rng(); u2 = rng()
        counter["uniforms"] += 2
        vlog = _log(u1 / (1.0 - u1))
        v = beta * vlog
        w = a * _exp(v)
        z = u1 * u1 * u2
        rr = gamma * v - 1.3862944
        s = a + rr - w
        if _ge(s + 2.609438, 5.0 * z):
            break
        t = _log(z)
        if _ge(s, t):
            break
        if _ge(rr + alpha * _log(alpha / (b + w)), t):
            break
    out = (w / (b + w)) if a0 == a else (b / (b + w))
    if SENS["on"]:
        wp = a * math.exp(beta * math.nextafter(vlog, math.inf))
        _sens_record(out, (wp / (b + wp)) if a0 == a else (b / (b + wp)), v)
    return out


def cheng_bc(a0, b0, rng, counter):
    """Cheng (1978) BC, for min(a,b) <= 1. Returns x in (0,1)."""
    a = max(a0, b0); b = min(a0, b0)
    alpha = a + b
    beta = 1.0 / b
    delta = 1.0 + a - b
    k1 = delta * (0.0138889 + 0.0416667 * b) / (a * beta - 0.777778)
    k2 = 0.25 + (0.5 + 0.25 / delta) * b
    while True:
        counter["attempts"] += 1
        u1 = rng(); u2 = rng()
        counter["uniforms"] += 2
        if not _ge(u1, 0.5):
            y = u1 * u2
            z = u1 * y
            if _ge(0.25 * u2 + z - y, k1):
                continue
        else:
            z = u1 * u1 * u2
            if _ge(0.25, z):
                vlog = _log(u1 / (1.0 - u1))
                v = beta * vlog
                w = a * _exp(v)
                break
            if _ge(z, k2):
                continue
        vlog = _log(u1 / (1.0 - u1))
        v = beta * vlog
        w = a * _exp(v)
        if _ge(alpha * (_log(alpha / (b + w)) + v) - 1.3862944, _log(z)):
            break
    # a = max(a0,b0), b = min(a0,b0); w/(b+w) ~ Beta(a, b). So the caller's
    # first parameter gets w/(b+w) when it WAS the max. Inverting this is a
    # silent defect: the sampler still returns a valid Beta variate, just of
    # the mirrored distribution. It was caught by the theoretical-mean check.
    return (w / (b + w)) if a0 == a else (b / (b + w))


def cheng_sample(a, b, rng, counter):
    if cheng_dispatch(a, b) == "BB":
        return cheng_bb(a, b, rng, counter)
    return cheng_bc(a, b, rng, counter)


# --- candidate A2: a PRACTICAL inverse CDF -------------------------------
# The bisection form above (`beta_ppf`) is deliberately the naive one: 60
# unconditional halvings, and a fresh `lgamma` triple on every evaluation.
# Rejecting the whole inverse-CDF family on that implementation would be
# rejecting an implementation, not a family. A2 is the same family done
# competently, and it is measured by the same counter.
#
# Two facts drive it:
#   1. alpha and beta are PER-DRIVER CONSTANTS. log B(a,b) is computed once
#      per driver, not once per sample, so the three lgamma calls amortise
#      over every sample that driver ever draws.
#   2. Newton on I_x(a,b) has a closed-form derivative - the Beta density -
#      so each step costs one extra exp/log pair, not a second CF expansion.
# Safeguarded: the bracket is maintained and a Newton step that leaves it is
# replaced by a bisection step, so it cannot diverge.
def log_beta(a, b):
    """Per-DRIVER constant. Computed once, reused for every sample."""
    return _lgamma(a) + _lgamma(b) - _lgamma(a + b)


def betai_with(a, b, x, lbeta):
    """Regularised incomplete beta, taking the per-driver log B(a,b)."""
    if x <= 0.0:
        return 0.0
    if x >= 1.0:
        return 1.0
    if x < (a + 1.0) / (a + b + 2.0):
        cf, it = betacf(a, b, x)
        OPS["betacf_iters"] = it
        return _exp(a * _log(x) + b * _log1p(-x) - lbeta) * cf / a
    cf, it = betacf(b, a, 1.0 - x)
    OPS["betacf_iters"] = it
    return 1.0 - _exp(b * _log1p(-x) + a * _log(x) - lbeta) * cf / b


def beta_pdf_with(a, b, x, lbeta):
    return _exp((a - 1.0) * _log(x) + (b - 1.0) * _log1p(-x) - lbeta)


def beta_ppf_newton(u, a, b, lbeta, tol=1e-13, maxit=20):
    """Safeguarded Newton. Returns (x, newton_iterations, cf_iterations_worst)."""
    lo, hi = 0.0, 1.0
    x = a / (a + b)                      # the mean: on this family, always in (0,1)
    cf_worst, it = 0, 0
    for it in range(1, maxit + 1):
        f = betai_with(a, b, x, lbeta) - u
        cf_worst = max(cf_worst, OPS["betacf_iters"])
        if f > 0.0:
            hi = x
        else:
            lo = x
        d = beta_pdf_with(a, b, x, lbeta)
        step = f / d if d > 0.0 else 0.0
        nx = x - step
        if nx == x:
            # The step is below the ULP of x: x IS the root to the precision a
            # Double can carry. Falling through to the safeguard here is a real
            # defect and it fired in testing - at convergence the just-updated
            # bracket end equals x, the strict test `lo < nx < hi` is false, and
            # the converged value is thrown away for a bisection of a bracket
            # whose far end was never tightened. Newton then re-approaches from
            # scratch and the iteration limit truncates it mid-descent, which is
            # what produced the 1.5e-05 .. 3.9e-05 errors before this branch
            # existed. Detected by comparing against the bisection reference.
            break
        if not (lo < nx < hi):           # safeguard: fall back to bisection
            nx = 0.5 * (lo + hi)
        delta = abs(nx - x)
        x = nx
        if delta < tol or (hi - lo) < tol:
            break
    return x, it, cf_worst
