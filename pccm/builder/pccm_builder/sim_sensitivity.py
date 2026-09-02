#!/usr/bin/env python3
"""PCCM Phase 7 Step-2 sensitivity mathematics.

Mid-ranks, Spearman as Pearson-on-ranks, the undefined case, and the ranking of
finished driver results. Pure: nothing here knows about RNG state, component
streams, drivers-as-worksheet-rows, `_SimData`, run identity or presentation. It
receives finished sequences of Doubles and returns statistics over them, exactly
as `sim_stats` does for the distribution.

--------------------------------------------------------------------------------
WHY THE SCALE-NORMALISATION OF `sim_stats` IS NOT REPEATED HERE
--------------------------------------------------------------------------------
`sim_stats` normalises because an iteration total may legally sit near `Double`
maximum. RANKS CANNOT. A mid-rank lies in [1, n] and n is bounded by the
accepted technical ceiling of 1 048 543 iterations, so the largest quantity any
correlation below can form is a centred sum of squares under n^3 / 4, about
2.9e17 - seventeen orders below the `Double` maximum. Copying the normalisation
machinery would be ceremony that guards against nothing, and ceremony that
guards against nothing is how a reader loses track of what the real guards are.

The CONTRIBUTIONS being ranked are unbounded, and they are checked for finiteness
before anything else happens. Only the ranks reach the arithmetic.

--------------------------------------------------------------------------------
WHY THE NO-TIES SHORTCUT IS ABSENT
--------------------------------------------------------------------------------
`1 - 6 * sum(d^2) / (n * (n^2 - 1))` is not an optimisation of what follows; on
tied data it is a different, wrong number. A Risk at probability 0.2 puts roughly
80% of its contribution column on one tied value, which is the ordinary case
here rather than a corner. The shortcut appears nowhere, including as a fast
path, and `sim_contract.yaml` forbids it.
"""

from __future__ import annotations

import math

# The undefined case is a STATUS, not a magic rho. Presentation - the string a
# reader sees - belongs to the reporting layer, not to the mathematics.
SENSITIVITY_DEFINED = 0
SENSITIVITY_NO_VARIANCE = 1


class SensitivityRefusal(Exception):
    """The kernel cannot answer, and says why rather than returning a number."""


def _checked(values, where: str) -> list[float]:
    if not values:
        raise SensitivityRefusal(f"sensitivity: an empty sequence has no {where}")
    out = []
    for value in values:
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            raise SensitivityRefusal(
                f"sensitivity: the {where} sequence carries a value that is not a "
                "finite Double")
        out.append(number)
    return out


def mid_ranks(values) -> list[float]:
    """Average ranks, 1-based, in the ORIGINAL order of `values`.

    A tie block occupying ordinal positions p..q takes (p + q) / 2, so
    [10, 20, 20, 20, 50] ranks as [1, 3, 3, 3, 5].

    THE INPUT IS NOT SORTED. A private sorted COPY is taken and each original
    value is located in it, which is why observation j keeps its rank at
    position j: the correlation downstream pairs contribution j with total j,
    and a rank vector that had been permuted would silently pair neither.

    TIES ARE EXACT EQUALITY, with no epsilon. Two contributions that differ by
    one ulp are two different outcomes of the model, and grouping them would be
    this module inventing a numerical policy the project has not settled.
    """
    checked = _checked(values, "mid-rank")
    ordered = sorted(checked)
    count = len(ordered)
    # One pass over the sorted copy records where each distinct value's tie
    # block starts and ends, so the per-observation lookup below is a dictionary
    # hit rather than a scan - the VBA implementation reaches the same answer
    # with a pair of binary searches.
    span: dict[float, tuple[int, int]] = {}
    start = 0
    for position in range(1, count + 1):
        if position == count or ordered[position] != ordered[start]:
            span[ordered[start]] = (start, position - 1)
            start = position
    return [(span[value][0] + 1 + span[value][1] + 1) / 2.0 for value in checked]


def rank_correlation(driver_ranks, total_ranks) -> tuple[float, int]:
    """Pearson over two ALREADY-RANKED series. Returns (rho, status).

    Taking ranks as the input rather than raw observations is what lets the
    total's ranks be computed once and reused for every driver: ranking the
    total D times would be D sorts of the same vector for the same answer.
    """
    x = _checked(driver_ranks, "driver rank")
    y = _checked(total_ranks, "total rank")
    if len(x) != len(y):
        raise SensitivityRefusal(
            f"sensitivity: {len(x)} driver observations against {len(y)} total "
            "observations; the pairing by iteration is lost")
    count = len(x)
    if count < 2:
        raise SensitivityRefusal(
            "sensitivity: a correlation needs at least two observations")

    mean_x = math.fsum(x) / count
    mean_y = math.fsum(y) / count
    sxx = math.fsum((v - mean_x) ** 2 for v in x)
    syy = math.fsum((v - mean_y) ** 2 for v in y)
    # UNDEFINED, NOT ZERO. A constant series has no dispersion to associate
    # with anything, so there is no monotone relationship to find - and rho = 0
    # would assert that one was looked for and not found.
    if sxx == 0.0 or syy == 0.0:
        return 0.0, SENSITIVITY_NO_VARIANCE
    sxy = math.fsum((a - mean_x) * (b - mean_y) for a, b in zip(x, y))
    rho = sxy / math.sqrt(sxx * syy)
    # Rounding can carry a perfect monotone association a hair outside [-1, 1];
    # clamping states the bound rather than publishing an impossible number.
    return max(-1.0, min(1.0, rho)), SENSITIVITY_DEFINED


def spearman(driver_values, total_ranks) -> tuple[float, int]:
    """One driver's Spearman rho against a PRECOMPUTED total-rank vector."""
    return rank_correlation(mid_ranks(driver_values), total_ranks)


def rank_drivers(results) -> list[int]:
    """The eligible drivers, ordered, as INDICES into `results`.

    `results` is a sequence of (permanent_id, rho, status). Indices are returned
    rather than a reordered copy, so the caller's sequence is untouched and each
    entry stays findable at the position it arrived in.

    ORDER: |rho| descending. A driver with no variance is not ranked at all -
    it has no rho to rank - and stays reportable through its status.

    THE TIE-BREAK IS THE PERMANENT ID, ascending, compared as ordinal UTF-16
    code units - the comparison this project already uses wherever driver order
    must be reproducible. Without it two drivers at equal |rho| would come out
    in whatever order the sort happened to produce, which is a different report
    from the same numbers. Worksheet row position is NOT available for this: a
    driver's row moves when an unrelated driver is added.
    """
    eligible = [
        (index, permanent_id, rho)
        for index, (permanent_id, rho, status) in enumerate(results)
        if status == SENSITIVITY_DEFINED
    ]
    ordered = sorted(
        eligible,
        key=lambda entry: (-abs(entry[2]), [ord(c) for c in entry[1]]),
    )
    return [index for index, _, _ in ordered]
