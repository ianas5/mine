#!/usr/bin/env python3
"""P7-5: the annual path end to end, through the REAL replay.

This is the suite the P7-5 correctness anchors live in. Nothing here is a
re-implementation: `modSimEngine`, `modSimAnnual`, `modSimStats`, `modSimRng`
and `modSimSample` are executed as their real `.bas` text through the accepted
Phase-6 transcriber, and the totals every annual sum is checked against come
from the REAL `SimEngineRun` in the same transcription.

WHAT MAKES THE RECONCILIATION MEAN SOMETHING. `TotalNom(j)` is produced by the
accepted simulation path with the scalar factor. `AnnualNom(j, y)` is produced
by the annual replay with a per-year factor. They share a seed, the component
streams, the canonical driver order and the contribution rule, and they share
nothing else - so agreement is evidence that the decomposition is the same
arithmetic regrouped, and disagreement is a real defect rather than two copies
of one mistake.

THE ALLOWANCE IS THE PROJECT'S OWN: `docs/phase5_plan.md` section 15, identities
I3c/I4c, with the ERRATUM C1 conditioning scale summed over CONTRIBUTIONS. No
result is scaled, nudged or normalised to make a sum come out.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

PCCM_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PCCM_ROOT / "builder"))
sys.path.insert(0, str(PCCM_ROOT / "tests"))

from pccm_builder import load_contract, load_sim_contract  # noqa: E402
from pccm_builder import sim_annual as annual  # noqa: E402
from pccm_builder.calc_loader import load_calc_contract  # noqa: E402
from pccm_builder.calc_numeric import (  # noqa: E402
    CalculationRefusal,
    safe_product,
    safe_signed_sum,
)
from pccm_builder.calc_cases import tolerances_from  # noqa: E402
from pccm_builder.calc_oracle import (  # noqa: E402
    AppliedTimeline,
    CalculationModel,
    CostDriver,
    FxRow,
    RiskDriver,
    calculate,
)
from pccm_builder.contract_loader import ContractError  # noqa: E402
from pccm_builder.sim_emit import render_sim_contract_module  # noqa: E402
from pccm_builder.sim_stats import percentile_type7  # noqa: E402
from pccm_builder.spec_loader import load_spec  # noqa: E402
from phase6_vba_transcribe import _Ref, _val, build as _build_transcription  # noqa: E402

SRC_VBA = PCCM_ROOT / "src" / "vba"
SPEC = PCCM_ROOT / "spec"
_CACHE: dict = {}

ABSOLUTE_FLOOR = 1e-6
RELATIVE_COEFFICIENT = 1e-12
SCALE_FLOOR = 1.0

SEED = 20260904
# The accepted business minimum. A smaller run is refused by the engine, and
# a suite that lowered it would be testing a configuration production forbids.
ITERATIONS = 1000
YEARS = 4
BLOCK_WIDTH = 12          # SIM_ANNUAL_BLOCK_WIDTH
DIST_OF = {"Triangular": 1, "Beta-PERT": 2, "Uniform": 3}


# ---------------------------------------------------------------------------
# Transcription
# ---------------------------------------------------------------------------
def _constants() -> dict:
    if "consts" not in _CACHE:
        out: dict = {}
        rendered = render_sim_contract_module(
            load_spec(SPEC / "workbook.yaml"),
            load_sim_contract(SPEC / "sim_contract.yaml"),
            load_contract(SPEC / "input_contract.yaml"))
        sources = [rendered]
        for name in ("modCalcFactors", "modSimAnnual", "modSimEngine"):
            sources.append((SRC_VBA / f"{name}.bas").read_text(encoding="utf-8"))
        for text in sources:
            for line in text.splitlines():
                match = re.match(r"^Public Const (\w+) As (\w+) = (.*)$", line)
                if not match:
                    continue
                name, kind, rest = match.groups()
                literal = rest.split("    '")[0].rstrip()
                out[name] = (literal[1:-1] if kind == "String"
                             else (float(literal) if kind == "Double" else int(literal)))
        _CACHE["consts"] = out
    return _CACHE["consts"]


def _safe_product(factors, count, result, *rest):
    try:
        result.v = safe_product([float(v) for v in factors[: int(_val(count))]], "annual")
    except (CalculationRefusal, ContractError):
        return False
    return True


def _safe_signed_sum(terms, count, result, *rest):
    try:
        result.v = safe_signed_sum([float(v) for v in terms[: int(_val(count))]], "annual")
    except (CalculationRefusal, ContractError):
        return False
    return True


def _vba() -> dict:
    if "vba" not in _CACHE:
        _CACHE["vba"] = _build_transcription(
            {
                "modSimRng": SRC_VBA / "modSimRng.bas",
                "modSimSample": SRC_VBA / "modSimSample.bas",
                "modSimStats": SRC_VBA / "modSimStats.bas",
                "modSimAnnual": SRC_VBA / "modSimAnnual.bas",
                "modSimEngine": SRC_VBA / "modSimEngine.bas",
                "modCalcFactors": SRC_VBA / "modCalcFactors.bas",
            },
            _constants(),
            only={
                "modCalcFactors": {"IsUsableDouble"},
                # modSimStats carries one procedure with an error handler the
                # transcriber cannot execute (SimStatsLadderExtent). The annual
                # path never reaches it, so only what the annual path DOES reach
                # is transcribed - the percentile and the position, and the
                # helpers each of those depends on.
                "modSimStats": {
                    "SimStatsQuantileType7", "SimStatsQuantileSorted",
                    "SimStatsQuantileLadder",
                    "SimStatsSortedCopy", "SimStatsSortAscending",
                    "SimStatsQuantilePosition", "SimStatsPositionOf",
                    "SimStatsOrderedIndices", "SimStatsSortIndices",
                    "SimStatsUsableSequence", "SimStatsUsableProbability",
                    "SimStatsConstantValue", "SimStatsUnitScale",
                },
            },
            signature_only={"modCalcFactors": {"SafeProduct", "SafeSignedSum",
                                               "SafeMultiply", "SafeDivide",
                                               "SafeSubtract"}},
            extra={
                "MAX_DOUBLE": sys.float_info.max,
                "SafeProduct": _safe_product,
                "SafeSignedSum": _safe_signed_sum,
                "SafeMultiply": lambda a, b, r, *x: _safe_product([a, b], _Ref(2), r),
                "SafeDivide": _unsupported("SafeDivide"),
                "SafeSubtract": _unsupported("SafeSubtract"),
            },
        )
    return _CACHE["vba"]


def _unsupported(name):
    def shim(*args):
        raise AssertionError(f"{name} is not exercised by the annual path")
    return shim


# ---------------------------------------------------------------------------
# THE MODEL
# ---------------------------------------------------------------------------
# Four project years, two currencies, two inflation profiles with real rates, a
# non-zero discount rate, Cost Lines and Risks, and profiles that put nothing in
# some years. Nominal and PV cannot coincide, and a driver absent from a year is
# a real case rather than a hypothetical.
def _model() -> CalculationModel:
    costs = (
        CostDriver("CL-001", "Triangular", "SAR", "Standard",
                   80.0, 100.0, 150.0, (0.25, 0.25, 0.25, 0.25), quantity=3.0),
        CostDriver("CL-002", "Beta-PERT", "USD", "Escalated",
                   200.0, 260.0, 400.0, (0.10, 0.20, 0.30, 0.40), quantity=1.0),
        CostDriver("CL-003", "Uniform", "SAR", "Escalated",
                   50.0, None, 90.0, (0.00, 0.50, 0.50, 0.00), quantity=2.0),
    )
    risks = (
        RiskDriver("R-001", "Triangular", "USD", "Standard",
                   1000.0, 2000.0, 4500.0, (0.50, 0.50, 0.00, 0.00), probability=0.3),
        RiskDriver("R-002", "Uniform", "SAR", "Escalated",
                   500.0, None, 2500.0, (0.00, 0.00, 0.40, 0.60), probability=0.7),
    )
    return CalculationModel(
        timeline=AppliedTimeline(2026, 2026, YEARS), discount_rate=0.08,
        fx_rows=(FxRow("SAR", 1), FxRow("USD", 3.75)),
        inflation_rates={
            "Standard": {2027: 0.03, 2028: 0.03, 2029: 0.03},
            "Escalated": {2027: 0.06, 2028: 0.055, 2029: 0.05},
        },
        cost_drivers=costs, risk_drivers=risks,
    )


def _tolerances():
    if "tol" not in _CACHE:
        _CACHE["tol"] = tolerances_from(load_calc_contract(SPEC / "calc_contract.yaml"))
    return _CACHE["tol"]


def _cumulative(rates: dict[int, float]) -> tuple[float, ...]:
    """The cumulative inflation factor per project year.

    Year 1 is the base year and carries 1.0; each later year multiplies by
    (1 + rate) for its calendar year. If this is ever wrong, `test_01` fails on
    the factor reconciliation - the reconstruction is CHECKED against Knom
    rather than assumed to match it.
    """
    out = [1.0]
    for offset in range(1, YEARS):
        out.append(out[-1] * (1.0 + rates[2026 + offset]))
    return tuple(out)


def _discount() -> tuple[float, ...]:
    return tuple(1.0 / (1.08 ** offset) for offset in range(YEARS))


def _records():
    """Resolved DriverFactors, with the per-year inputs Phase 5 now carries."""
    if "records" in _CACHE:
        return _CACHE["records"]
    model = _model()
    resolved = calculate(model, _tolerances())
    by_id = {}
    for record in resolved.drivers:
        kind = str(getattr(record.driver_kind, "value", record.driver_kind)).upper()
        by_id[("RISK" in kind, record.permanent_id)] = record
    fx = {"SAR": 1.0, "USD": 3.75}
    out = []
    for is_risk, drivers in ((False, model.cost_drivers), (True, model.risk_drivers)):
        for driver in drivers:
            record = by_id[(is_risk, driver.permanent_id)]
            out.append({
                "PermanentId": driver.permanent_id,
                "IsRisk": is_risk,
                "Knom": float(record.knom),
                "Kpv": float(record.kpv),
                "Quantity": 0.0 if is_risk else float(record.quantity),
                "Probability": float(record.probability) if is_risk else 0.0,
                "DistKind": DIST_OF[driver.distribution],
                "CentralBasis": "",
                "MinValue": float(driver.min_value),
                "MostLikely": 0.0 if driver.most_likely is None else float(driver.most_likely),
                "MaxValue": float(driver.max_value),
                "Central": 0.0,
                "MeanValue": 0.0,
                "FxRate": fx[driver.currency],
                "Weights": list(driver.profile_weights),
                "Inflation": list(_cumulative(model.inflation_rates[driver.inflation_profile])),
            })
    _CACHE["records"] = out
    return out


def _run_totals():
    """TotalNom and TotalPV from the REAL accepted simulation path."""
    if "totals" in _CACHE:
        return _CACHE["totals"]
    nominal, pv, detail = [], [], _Ref("")
    ok = _vba()["SimEngineRun"](
        _records(), _Ref(len(_records())), _Ref(SEED), _Ref(ITERATIONS),
        nominal, pv, detail)
    assert ok, f"the accepted run refused: {detail.v}"
    _CACHE["totals"] = (list(nominal), list(pv))
    return _CACHE["totals"]


def _year_factors(measure: str):
    """Per-year factors in SUPPLY order, stride YEARS, built by the VBA owner."""
    key = f"factors:{measure}"
    if key in _CACHE:
        return _CACHE[key]
    flat = []
    for record in _records():
        result, detail = [], _Ref("")
        ok = _vba()["SimAnnualFactors"](
            _Ref(record["FxRate"]), list(record["Weights"]), list(record["Inflation"]),
            list(_discount()), _Ref(measure == "PV"), result, detail)
        assert ok, f"{record['PermanentId']}: {detail.v}"
        assert len(result) == YEARS
        flat.extend(float(v) for v in result)
    _CACHE[key] = flat
    return flat


def _annual_block(measure: str, first_year: int, year_count: int):
    column, detail = [], _Ref("")
    ok = _vba()["SimEngineReplayAnnualBlock"](
        _records(), _Ref(len(_records())), _Ref(SEED), _Ref(ITERATIONS),
        _year_factors(measure), _Ref(YEARS), _Ref(first_year), _Ref(year_count),
        _Ref(measure), column, detail)
    assert ok, f"annual replay refused: {detail.v}"
    return [float(v) for v in column]


def _annual_matrix(measure: str, block_width: int = BLOCK_WIDTH):
    """The full duration, assembled from blocks. One block per pass."""
    matrix = [[0.0] * YEARS for _ in range(ITERATIONS)]
    first = 0
    while first < YEARS:
        width = min(block_width, YEARS - first)
        column = _annual_block(measure, first, width)
        for iteration in range(ITERATIONS):
            for offset in range(width):
                matrix[iteration][first + offset] = column[iteration * width + offset]
        first += width
    return matrix


def _allowance(measure: str, iterations):
    """I3c / I4c, conditioned on CONTRIBUTIONS."""
    scalar = "Kpv" if measure == "PV" else "Knom"
    factors = _year_factors(measure)
    annual_terms, total_terms = [], []
    matrix = _annual_matrix(measure)
    for j in iterations:
        for index, record in enumerate(_records()):
            total_terms.append(abs(record[scalar]))
            for y in range(YEARS):
                annual_terms.append(abs(factors[index * YEARS + y]))
        annual_terms.extend(abs(v) for v in matrix[j])
    return annual.reconciliation_allowance(
        annual_terms, total_terms, ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR)


# ===========================================================================
# 1. THE FACTOR DECOMPOSITION, THROUGH THE VBA OWNER
# ===========================================================================
@pytest.mark.parametrize("measure,scalar", [("nominal", "Knom"), ("PV", "Kpv")])
def test_01_the_vba_per_year_factors_sum_to_the_accepted_factor(measure, scalar) -> None:
    """`sum_y K_y = K`, with K taken from the accepted Phase-5 resolution.

    This is also what proves the fixture's inflation reconstruction is the one
    the model actually resolved: a wrong cumulative factor could not reconcile.
    """
    factors = _year_factors(measure)
    for index, record in enumerate(_records()):
        parts = factors[index * YEARS:(index + 1) * YEARS]
        whole = record[scalar]
        recombined = safe_signed_sum(list(parts), "factor recombination")
        allowance = annual.reconciliation_allowance(
            list(parts), [whole], ABSOLUTE_FLOOR, RELATIVE_COEFFICIENT, SCALE_FLOOR)
        assert abs(recombined - whole) <= allowance, (
            f"{record['PermanentId']} {scalar}: sum_y = {recombined!r}, accepted = {whole!r}")


def test_02_the_vba_factors_equal_the_python_reference() -> None:
    for measure, with_discount in (("nominal", False), ("PV", True)):
        flat = _year_factors(measure)
        for index, record in enumerate(_records()):
            if with_discount:
                expected = annual.per_year_pv_factors(
                    record["FxRate"], record["Weights"], record["Inflation"], _discount())
            else:
                expected = annual.per_year_nominal_factors(
                    record["FxRate"], record["Weights"], record["Inflation"])
            actual = tuple(flat[index * YEARS:(index + 1) * YEARS])
            assert actual == expected, f"{record['PermanentId']} {measure}"


# ===========================================================================
# 2. THE PRIMARY ANCHOR: EVERY ITERATION RECONCILES
# ===========================================================================
@pytest.mark.parametrize("measure,which", [("nominal", 0), ("PV", 1)])
def test_03_every_iteration_annual_sum_equals_the_run_total(measure, which) -> None:
    """`sum_y Annual(j, y) = Total(j)` for all 200 iterations, nominal and PV.

    The totals come from the accepted SimEngineRun; the annual values from the
    annual replay. Nothing is scaled to make them agree.
    """
    totals = _run_totals()[which]
    matrix = _annual_matrix(measure)
    allowance = _allowance(measure, range(ITERATIONS))
    worst, worst_j = 0.0, -1
    for j in range(ITERATIONS):
        recombined = safe_signed_sum(list(matrix[j]), "annual recombination")
        difference = abs(recombined - totals[j])
        if difference > worst:
            worst, worst_j = difference, j
    assert worst <= allowance, (
        f"{measure}: iteration {worst_j} is out by {worst!r}, allowance {allowance!r}")


def test_04_nominal_and_pv_are_independently_produced_and_differ() -> None:
    """A discount rate of 8% over four years: PV cannot equal nominal.

    A test whose model collapses Knom to Kpv would pass with the two measures
    confused, which is the defect most worth catching here.
    """
    nominal, pv = _run_totals()
    assert any(abs(a - b) > 1.0 for a, b in zip(nominal, pv)), (
        "the fixture cannot tell nominal from PV")
    nominal_matrix = _annual_matrix("nominal")
    pv_matrix = _annual_matrix("PV")
    assert nominal_matrix != pv_matrix
    # Year 1 is undiscounted, so the two agree there and nowhere else.
    for j in range(0, ITERATIONS, 37):
        assert nominal_matrix[j][0] == pytest.approx(pv_matrix[j][0], rel=1e-12)
        assert any(abs(nominal_matrix[j][y] - pv_matrix[j][y]) > 1e-9
                   for y in range(1, YEARS))


# ===========================================================================
# 3. BLOCKING CHANGES MEMORY, NOT ANSWERS
# ===========================================================================
@pytest.mark.parametrize("block_width", [1, 2, 3, YEARS, BLOCK_WIDTH])
def test_05_blocked_execution_equals_the_single_pass_result(block_width) -> None:
    reference = _annual_matrix("nominal", block_width=YEARS)
    blocked = _annual_matrix("nominal", block_width=block_width)
    assert blocked == reference, f"block width {block_width} changed the answer"


def test_06_the_block_count_is_the_contracted_number_of_passes() -> None:
    """`passes = ceil(applied_duration / block_width)`."""
    for duration, width, expected in (
            (1, 12, 1), (12, 12, 1), (13, 12, 2), (24, 12, 2), (25, 12, 3),
            (200, 12, 17), (4, 1, 4), (0, 12, 0), (5, 0, 0)):
        assert _vba()["SimAnnualBlockCount"](_Ref(duration), _Ref(width)) == expected, (
            duration, width)


# ===========================================================================
# 4. THE SELECTED-Px PROFILE USES THE TOTAL'S OWN POSITION
# ===========================================================================
def _profile(measure: str, which: int, p: float):
    totals = _run_totals()[which]
    position = _vba()["_new"]("SimStatsPosition")
    detail = _Ref("")
    ok = _vba()["SimStatsQuantilePosition"](
        list(totals), _Ref(len(totals)), _Ref(p), position, detail)
    assert ok, detail.v
    matrix = _annual_matrix(measure)
    result = []
    ok = _vba()["SimAnnualProfile"](
        list(matrix[position["LoSource"]]), list(matrix[position["HiSource"]]),
        _Ref(YEARS), _Ref(position["Fraction"]), result, detail)
    assert ok, detail.v
    return position, [float(v) for v in result], matrix


@pytest.mark.parametrize("measure,which", [("nominal", 0), ("PV", 1)])
@pytest.mark.parametrize("p", [0.5, 0.8, 0.9])
def test_07_the_profile_reconciles_to_the_reported_percentile(measure, which, p) -> None:
    """`sum_y Profile_Px(y) = Px`, nominal and PV, with f strictly interior."""
    totals = _run_totals()[which]
    position, profile, _ = _profile(measure, which, p)
    assert 0.0 < position["Fraction"] < 1.0, f"p={p} gave f={position['Fraction']}"
    recombined = safe_signed_sum(profile, "profile recombination")
    reported = percentile_type7(totals, p)
    allowance = _allowance(measure, [position["LoSource"], position["HiSource"]])
    assert abs(recombined - reported) <= allowance, (
        f"{measure} p={p}: sum_y Profile = {recombined!r}, Px = {reported!r}")


@pytest.mark.parametrize("p", [0.5, 0.8])
def test_08_the_profile_uses_the_totals_exact_lo_hi_and_f(p) -> None:
    """Not a nearby iteration, not a per-year percentile - those two vectors."""
    position, profile, matrix = _profile("nominal", 0, p)
    reference = annual.percentile_position(_run_totals()[0], p)
    assert position["LoSource"] == reference.lo
    assert position["HiSource"] == reference.hi
    assert position["Fraction"] == reference.fraction
    low, high, f = matrix[reference.lo], matrix[reference.hi], reference.fraction
    for year in range(YEARS):
        expected = low[year] if low[year] == high[year] else (
            (1.0 - f) * low[year] + f * high[year])
        assert profile[year] == expected, f"year {year + 1}"


@pytest.mark.parametrize("p", [0.0, 1.0])
def test_09_at_f_zero_only_the_exact_source_vector_is_used(p) -> None:
    """The profile IS that iteration's own annual vector, bit for bit."""
    position, profile, matrix = _profile("nominal", 0, p)
    assert position["Fraction"] == 0.0
    assert profile == matrix[position["LoSource"]]


def test_10_a_fraction_outside_zero_to_one_is_refused() -> None:
    result, detail = [], _Ref("")
    ok = _vba()["SimAnnualProfile"](
        [1.0] * YEARS, [2.0] * YEARS, _Ref(YEARS), _Ref(1.5), result, detail)
    assert not ok and "outside" in detail.v


# ===========================================================================
# 5. THE LADDER IS NOT THE PROFILE
# ===========================================================================
def test_11_the_per_year_ladder_and_the_profile_stay_distinct() -> None:
    """Different objects, and this model makes the difference visible."""
    p = 0.8
    matrix = _annual_matrix("nominal")
    column = [matrix[j][y] for j in range(ITERATIONS) for y in range(YEARS)]
    ladder, detail = [], _Ref("")
    ok = _vba()["SimAnnualLadder"](
        column, _Ref(ITERATIONS), _Ref(YEARS), [p], _Ref(1), ladder, detail)
    assert ok, detail.v
    per_year = [float(v) for v in ladder]
    _, profile, _ = _profile("nominal", 0, p)
    assert per_year != profile, "the per-year ladder equals the profile"
    reported = percentile_type7(_run_totals()[0], p)
    assert abs(sum(per_year) - reported) > 1.0, (
        "the per-year ladder sums to the total percentile in this model, so it "
        "cannot demonstrate that they are different objects")


def test_12_the_vba_ladder_equals_the_python_reference() -> None:
    probabilities = [0.1, 0.5, 0.9]
    matrix = _annual_matrix("nominal")
    column = [matrix[j][y] for j in range(ITERATIONS) for y in range(YEARS)]
    ladder, detail = [], _Ref("")
    ok = _vba()["SimAnnualLadder"](
        column, _Ref(ITERATIONS), _Ref(YEARS), list(probabilities),
        _Ref(len(probabilities)), ladder, detail)
    assert ok, detail.v
    for year in range(YEARS):
        series = [matrix[j][year] for j in range(ITERATIONS)]
        for rung, p in enumerate(probabilities):
            assert float(ladder[year * len(probabilities) + rung]) == percentile_type7(
                series, p), f"year {year + 1} rung {p}"


# ===========================================================================
# 6. NOTHING UPSTREAM IS DISTURBED
# ===========================================================================
def test_13_the_annual_path_mutates_no_source_and_no_identity() -> None:
    """The replay reads; it does not write.

    The resolved factors, the per-year factors and the published totals are all
    compared before and after a full annual pass. A statistic or a replay that
    reordered or rewrote its source would change the result digest as a side
    effect, and the digest is a published identity.
    """
    before_records = [dict(r, Weights=list(r["Weights"]),
                           Inflation=list(r["Inflation"])) for r in _records()]
    before_factors = list(_year_factors("nominal"))
    before_totals = [list(_run_totals()[0]), list(_run_totals()[1])]

    _annual_matrix("nominal")
    _annual_matrix("PV")
    _profile("nominal", 0, 0.8)

    assert [dict(r, Weights=list(r["Weights"]), Inflation=list(r["Inflation"]))
            for r in _records()] == before_records
    assert list(_year_factors("nominal")) == before_factors
    assert [list(_run_totals()[0]), list(_run_totals()[1])] == before_totals


def test_13b_the_factors_follow_permanent_identity_not_supply_position() -> None:
    """The physical order of the caller's array must reach nothing.

    The engine maps supply order to canonical order by PERMANENT ID. A mapping
    that used position instead would agree with identity on any fixture whose
    supply order happens to be canonical - which the main fixture's is - so this
    supplies the drivers REVERSED and requires the same answer.
    """
    reference = _annual_matrix("nominal")
    records = _records()
    order = list(range(len(records)))[::-1]
    shuffled = [records[i] for i in order]
    factors = _year_factors("nominal")
    shuffled_factors = []
    for i in order:
        shuffled_factors.extend(factors[i * YEARS:(i + 1) * YEARS])

    column, detail = [], _Ref("")
    ok = _vba()["SimEngineReplayAnnualBlock"](
        shuffled, _Ref(len(shuffled)), _Ref(SEED), _Ref(ITERATIONS),
        shuffled_factors, _Ref(YEARS), _Ref(0), _Ref(YEARS),
        _Ref("nominal"), column, detail)
    assert ok, detail.v
    for j in range(ITERATIONS):
        for y in range(YEARS):
            assert float(column[j * YEARS + y]) == reference[j][y], (
                f"supply order changed the answer at iteration {j}, year {y + 1}")


def test_14_the_annual_replay_is_deterministic() -> None:
    """Same seed, same numbers. A second stochastic sequence would not repeat."""
    first = _annual_matrix("nominal")
    _CACHE.pop("records", None)
    _CACHE.pop("factors:nominal", None)
    second = _annual_matrix("nominal")
    assert first == second


def test_15_the_annual_path_touches_no_run_identity() -> None:
    """Structural: the replay carries no run id, nonce or persisted state.

    It cannot consume one - there is nothing in its signature or body to consume
    - and this is what says so rather than a comment claiming it.
    """
    body = (SRC_VBA / "modSimEngine.bas").read_text(encoding="utf-8")
    start = body.index("Public Function SimEngineReplayAnnualBlock")
    end = body.index("' ==========", start)
    replay = body[start:end]
    for forbidden in ("RunId", "run_id", "Nonce", "nonce", "F21", "Worksheet",
                      "Range(", "_SimData", "Digest", "Fingerprint"):
        assert forbidden not in replay, forbidden
    annual_source = (SRC_VBA / "modSimAnnual.bas").read_text(encoding="utf-8")
    for forbidden in ("RunId", "Nonce", "Worksheet", "Range(", "_SimData",
                      "Digest", "Fingerprint", "SimRng", "SimSample", "Randomize",
                      "Sensitivity"):
        assert forbidden not in annual_source, forbidden
